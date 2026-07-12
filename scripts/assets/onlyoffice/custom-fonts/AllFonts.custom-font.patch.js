window["__custom_font_registry__"] = {
  "1002": ["FZXiaoBiaoSong-B05S", "方正小标宋简体"],
  "1003": ["FangSong_GB2312", "仿宋_GB2312"],
  "1004": ["SimHei", "黑体"],
  "1005": ["KaiTi_GB2312", "楷体_GB2312"]

};
/**
 * 根据 window.__custom_font_registry__ 同步写入 __fonts_files / __fonts_infos。
 * 在 AllFonts.js 中紧接 registry 定义之后执行（SDK 初始化前）。
 */
(function () {
  var registry = window.__custom_font_registry__;
  var files = window.__fonts_files;
  var infos = window.__fonts_infos;
  if (!registry || !files || !infos) {
    return;
  }

  function findInfoIndex(name) {
    for (var i = 0; i < infos.length; i++) {
      if (infos[i][0] === name) {
        return i;
      }
    }
    return -1;
  }

  function listFontNames(entry) {
    if (Array.isArray(entry)) {
      return entry;
    }
    if (entry && entry.aliases) {
      return entry.aliases;
    }
    return [];
  }

  for (var id in registry) {
    if (!Object.prototype.hasOwnProperty.call(registry, id)) {
      continue;
    }

    if (files.indexOf(id) < 0) {
      files.push(id);
    }
    var fileIndex = files.indexOf(id);
    var names = [];
    var seen = {};
    var list = listFontNames(registry[id]);

    for (var i = 0; i < list.length; i++) {
      if (list[i]) {
        names.push(list[i]);
      }
    }

    for (var n = 0; n < names.length; n++) {
      var name = names[n];
      if (!name || seen[name]) {
        continue;
      }
      seen[name] = true;
      var row = [name, fileIndex, 0, -1, -1, -1, -1, -1, -1];
      var idx = findInfoIndex(name);
      if (idx >= 0) {
        infos[idx] = row;
      } else {
        infos.push(row);
      }
    }
  }

  // Word SDK 初始化后会 delete __fonts_files；Cell SDK 后加载时需用快照重建 yyc/JPb/KPb。
  window.__custom_font_catalog_snapshot__ = {
    files: files.slice(),
    infos: infos.map(function (row) {
      return row.slice();
    }),
  };
})();

// 9.4 已直接从上面的 __fonts_files / __fonts_infos 构建 XDc / Vbb。
// web-apps 比 SDK 晚加载，单独刷新 ComboBox 的搜索缓存即可；不改写字体
// 加载器、二进制流或文档排版管线，避免影响 OnlyOffice 自带字体渲染。
(function (window) {
  "use strict";
  var registry = window.__custom_font_registry__ || {};

  function patchComboBoxFonts() {
    var proto =
      window.Common &&
      window.Common.UI &&
      window.Common.UI.ComboBoxFonts &&
      window.Common.UI.ComboBoxFonts.prototype;
    if (!proto || proto.__CUSTOM_FONT_REGISTRY_COMBO_PATCHED__) {
      return !!proto;
    }

    if (typeof proto.selectCandidate === "function") {
      var origSelectCandidate = proto.selectCandidate;
      proto.selectCandidate = function (isExact) {
        // 9.4 会把 imgidx 当 sprite index。自定义字体记录有时将文件偏移
        // （例如 134400）写到这里，导致 getImage 分配超大 TypedArray。
        if (this.store && typeof this.store.each === "function") {
          this.store.each(function (model) {
            if (!model || typeof model.get !== "function") return;
            var idx = model.get("imgidx");
            if (typeof idx !== "number" || !isFinite(idx) || idx < 0 || idx > 512) {
              model.set("imgidx", 0);
            }
          });
        }
        if (this.store && typeof this.store.toJSON === "function") {
          this._fontsArray = this.store.toJSON();
        }
        return origSelectCandidate.call(this, isExact);
      };
    }

    proto.__CUSTOM_FONT_REGISTRY_COMBO_PATCHED__ = true;
    return true;
  }

  // 9.4 的 FontPicker（AscFonts.cU）有一份独立的名称映射。未知名称会
  // 回退到“等线”，即使 Vbb 已含有自定义字体。对 registry 名称直接返回
  // Vbb 条目，后续仍完全交由 OnlyOffice 原生 iM 加载器请求 fonts/{id}。
  function patchNativeFontPicker() {
    var asc = window.AscFonts;
    var picker = asc && asc.cU;
    if (!asc || !picker || !asc.Vbb || !asc.ucc) {
      return false;
    }
    if (picker.__CUSTOM_FONT_REGISTRY_PICKER_PATCHED__) {
      return true;
    }

    var original = picker.WD || picker.cSf;
    var originalResolveName = picker.FVc;
    var originalLoadFace = picker.IH;
    if (typeof original !== "function") {
      return false;
    }
    var names = {};
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var aliases = Array.isArray(registry[id]) ? registry[id] : [];
      for (var index = 0; index < aliases.length; index++) {
        names[aliases[index]] = true;
      }
    }

    var resolve = function (name, style, output) {
      var catalogIndex = names[name] ? asc.ucc[name] : undefined;
      if (catalogIndex !== undefined && asc.Vbb[catalogIndex]) {
        if (output) {
          output.xa = name;
          output.AW = null;
        }
        return asc.Vbb[catalogIndex];
      }
      return original.apply(this, arguments);
    };
    picker.cSf = resolve;
    picker.WD = resolve;
    // Word 在把 TextPr 写回文档模型、以及重绘时会单独调用 FVc() 获取
    // 规范化后的字体名。只覆盖 WD/cSf 会让这里再次归一化到 fallback，
    // 因而发生“工具栏是自定义名、画布仍用 Arial”的假成功。
    if (typeof originalResolveName === "function") {
      picker.FVc = function (name) {
        return names[name] ? name : originalResolveName.apply(this, arguments);
      };
    }
    // Canvas 排版不走 WD()，而是 Ay.ila() → FontPicker.IH()。未覆盖此处时
    // IH 会将未知名模糊匹配为 Arial；字体文件虽被请求，却从不进入 FreeType cache。
    if (typeof originalLoadFace === "function") {
      picker.IH = function (name) {
        var catalogIndex = names[name] ? asc.ucc[name] : undefined;
        var entry =
          catalogIndex !== undefined && asc.Vbb[catalogIndex]
            ? asc.Vbb[catalogIndex]
            : null;
        if (entry && typeof entry.IH === "function") {
          var args = Array.prototype.slice.call(arguments, 1);
          return entry.IH.apply(entry, args);
        }
        return originalLoadFace.apply(this, arguments);
      };
    }
    picker.__CUSTOM_FONT_REGISTRY_PICKER_PATCHED__ = true;
    return true;
  }

  // Cell / Slide 的 9.4 picker 不会从追加的 __fonts_infos 建立名称索引，
  // 因而会为自定义名返回 Arial（文件 022）。只克隆返回的文件描述并替换
  // 文件 id，不触碰原生 catalog、内置字体或布局管线。
  function patchRuntimeFontPicker() {
    var asc = window.AscFonts;
    if (!asc || typeof asc.pickFont !== "function" || asc.__CUSTOM_PICK_FONT_PATCHED__) {
      return false;
    }

    var original = asc.pickFont.bind(asc);
    asc.pickFont = function (name) {
      var file = original.apply(this, arguments);
      var customId = null;
      for (var id in registry) {
        if (!Object.prototype.hasOwnProperty.call(registry, id)) continue;
        var aliases = Array.isArray(registry[id])
          ? registry[id]
          : registry[id] && registry[id].aliases
            ? registry[id].aliases
            : [];
        if (aliases.indexOf(name) >= 0) {
          customId = id;
          break;
        }
      }
      if (!customId || !file) return file;

      var result = Object.create(Object.getPrototypeOf(file));
      for (var key in file) result[key] = file[key];
      if ("Xa" in result) result.Xa = customId;
      if ("Na" in result) result.Na = customId;
      return result;
    };
    asc.__CUSTOM_PICK_FONT_PATCHED__ = true;
    return true;
  }

  // Cell / Slide 的异步字体加载器不使用 AscFonts.pickFont，而是直接调用
  // AscFonts.JZ.rM(name)。9.4 在启动时已把 custom file 加进 hOc/y1b，但
  // JZ 的名称索引早于该步骤构建，未知字体会在这里回退 Arial，因而从不请求
  // fonts/{customId}。克隆 Arial 的 family 描述，仅把四种样式指到 custom
  // file；内置字体仍由原始 rM 完整处理。
  function patchCellSlideFontResolver() {
    var asc = window.AscFonts;
    var resolver = asc && asc.JZ;
    if (!asc || !resolver || !asc.hOc || typeof resolver.rM !== "function") {
      return false;
    }
    if (resolver.__CUSTOM_FONT_REGISTRY_RESOLVER_PATCHED__) {
      return true;
    }

    var fileIndexes = {};
    for (var index = 0; index < asc.hOc.length; index++) {
      var file = asc.hOc[index];
      var fileId = file && (file.Xa || file.Na);
      if (fileId && Object.prototype.hasOwnProperty.call(registry, fileId)) {
        fileIndexes[fileId] = index;
      }
    }
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id) || fileIndexes[id] === undefined) {
        return false;
      }
    }

    var aliasesToId = {};
    for (var customId in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, customId)) continue;
      var aliases = Array.isArray(registry[customId])
        ? registry[customId]
        : registry[customId] && registry[customId].aliases
          ? registry[customId].aliases
          : [];
      for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
        aliasesToId[aliases[aliasIndex]] = customId;
      }
    }

    // Cell 的画布排版会直接调用 xMd(name)，并基于其 Rma 再查询
    // y1b[z1b[Rma]]；只 hook rM 无法覆盖这条文档加载后的重绘路径。
    // 保留 Arial 的 matcher 信息，但让 custom 名进入已存在的 y1b catalog。
    var originalResolveName = resolver.xMd.bind(resolver);
    resolver.xMd = function (name) {
      var customId = aliasesToId[name];
      if (!customId) {
        return originalResolveName.apply(this, arguments);
      }
      var fallback = originalResolveName("Arial");
      if (!fallback) {
        return originalResolveName.apply(this, arguments);
      }
      var resolved = Object.create(Object.getPrototypeOf(fallback));
      for (var key in fallback) resolved[key] = fallback[key];
      resolved.aP = name;
      resolved.Rma = name;
      return resolved;
    };

    var original = resolver.rM.bind(resolver);
    resolver.rM = function (name) {
      var customId = aliasesToId[name];
      if (!customId) {
        return original.apply(this, arguments);
      }
      var fallback = original("Arial");
      if (!fallback) {
        return original.apply(this, arguments);
      }
      var resolved = Object.create(Object.getPrototypeOf(fallback));
      for (var key in fallback) resolved[key] = fallback[key];
      var fileIndex = fileIndexes[customId];
      resolved.Fa = name;
      resolved.vja = fileIndex;
      resolved.tua = fileIndex;
      resolved.WE = fileIndex;
      resolved.sua = fileIndex;
      return resolved;
    };
    resolver.__CUSTOM_FONT_REGISTRY_RESOLVER_PATCHED__ = true;
    return true;
  }

  // Slide 的 LU.gLc(name) 会把未知文档字体名规范化为 Arial。虽然 custom
  // family 已由 AllFonts 构建进 $Jb/r3b，但若只在 FG 返回时克隆 Arial，
  // 排版缓存仍会沿用 fallback。直接让 gLc 返回原生 custom family。
  function patchSlideFontResolver() {
    var asc = window.AscFonts;
    var resolver = asc && asc.LU;
    if (!asc || !resolver || !asc.Stc || typeof resolver.FG !== "function") {
      return false;
    }
    if (resolver.__CUSTOM_FONT_REGISTRY_RESOLVER_PATCHED__) {
      return true;
    }

    var fileIndexes = {};
    for (var index = 0; index < asc.Stc.length; index++) {
      var file = asc.Stc[index];
      var fileId = file && (file.Na || file.Xa);
      if (fileId && Object.prototype.hasOwnProperty.call(registry, fileId)) {
        fileIndexes[fileId] = index;
      }
    }
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id) || fileIndexes[id] === undefined) {
        return false;
      }
    }

    var aliasesToId = {};
    for (var customId in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, customId)) continue;
      var aliases = Array.isArray(registry[customId])
        ? registry[customId]
        : registry[customId] && registry[customId].aliases
          ? registry[customId].aliases
          : [];
      for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
        aliasesToId[aliases[aliasIndex]] = customId;
      }
    }

    var original = resolver.gLc && resolver.gLc.bind(resolver);
    if (typeof original !== "function" || !asc.$Jb || !asc.r3b) {
      return false;
    }
    resolver.gLc = function (name) {
      var customId = aliasesToId[name];
      if (!customId) {
        return original.apply(this, arguments);
      }
      if (asc.$Jb[name] === undefined || !asc.r3b[asc.$Jb[name]]) {
        return original.apply(this, arguments);
      }
      return { TL: name, fba: name };
    };
    resolver.__CUSTOM_FONT_REGISTRY_RESOLVER_PATCHED__ = true;
    return true;
  }

  // AllFonts 早于 Slide SDK 加载。仅靠 50ms 轮询可能在首次 TJb 文档字体
  // 收集之后才包住 LU.FG；监听 LU 赋值并在同一同步调用栈中安装 resolver。
  function watchSlideFontResolver() {
    var asc = (window.AscFonts = window.AscFonts || {});
    if (asc.__CUSTOM_FONT_REGISTRY_LU_WATCHED__) {
      return true;
    }
    var current = asc.LU;
    asc.__CUSTOM_FONT_REGISTRY_LU_WATCHED__ = true;
    Object.defineProperty(asc, "LU", {
      configurable: true,
      enumerable: true,
      get: function () {
        return current;
      },
      set: function (value) {
        current = value;
        patchSlideFontResolver();
      },
    });
    return true;
  }

  watchSlideFontResolver();
  var tries = 0;
  var timer = window.setInterval(function () {
    var comboPatched = patchComboBoxFonts();
    var pickerPatched = patchNativeFontPicker();
    var runtimePickerPatched = patchRuntimeFontPicker();
    watchSlideFontResolver();
    var resolverPatched = patchCellSlideFontResolver() || patchSlideFontResolver();
    if ((comboPatched && pickerPatched && runtimePickerPatched && resolverPatched) || ++tries > 1200) {
      window.clearInterval(timer);
    }
  }, 50);
})(window);

/**
 * AllFonts catalog 自定义字体运行时补丁（构建时同步内联到 AllFonts.js 末尾）。
 * catalog 由 __custom_font_registry__ 经 apply-custom-font-registry.js 写入上方数组。
 *
 * Word / Excel / Slide 走三套字体管线，补丁代码按编辑器分区。
 * - Word: dpc / L1b / QQ.sgd / mJ.GPb + Lif
 * - Excel: yyc / KPb / JPb / dW / V_.koc + workbook reload
 * - Slide: jec / hyb / jR.quc / lU.vFb + s7e + TJb(ea.N_) reload
 *
 * 公共部分：registry、二进制同步加载、字体名归一化 resolveDocumentFontName。
 * Excel 打开时正确渲染依赖：catalog 快照 → syncCellEngineFontRefs → 二进制注入 →
 * FontPicker 绕过 → reloadCellDocumentFontsFromWorkbook → wa 就绪后 refresh。
 * Slide 打开时正确渲染依赖：uji/lU hook → prepareSlideFontCatalog → 二进制注入 kec →
 * patch jR.quc/EE → 一次 TJb(ea.N_) reload → JOf 后 Me+Rp（禁止 yd.WN，避免缩略图栏消失）。
 */
(function (window) {
  "use strict";

  // 这份旧实现会改写 9.4 Cell catalog，不能执行；PPT 现由上面的 LU 同步
  // watcher 处理，因此此处仅保留作符号比对。
  return;

  var XOR_KEY = [
    160, 102, 214, 32, 20, 150, 71, 250, 149, 105, 184, 80, 176, 65, 73, 72,
  ];

  var PATCHED_WORD = false;
  var WORD_BINARIES_INJECTED = false;
  var PATCHED_CELL = false;
  var PATCHED_SLIDE = false;
  // --- Slide 状态机（勿随意重置；reload 只允许一次，否则易与 ZDc/JOf 形成死循环）---
  var SLIDE_BINARIES_INJECTED = false; // 1001 等已 sync 写入 kec
  var SLIDE_CATALOG_READY = false; // prepareSlideFontCatalog 已成功
  var SLIDE_FONTS_RELOADED = false; // yga.TJb(ea.N_) 已重跑过一次
  var SLIDE_RELOAD_SCHEDULED = false; // scheduleSlideFontReloadOnceWhenReady 已启动
  var SLIDE_UNPATCHED_TJB = false; // 首次 TJb 时 patch 仍未完成（调试用）
  var SLIDE_CATALOG_PREPARING = false; // prepare 重入保护，防止 vFb→ensure 栈溢出
  var SLIDE_LAYOUT_REFRESHED = false; // JOf 后 Me+Rp 已执行过一次

  // catalog / 二进制 / FontPicker 三者都就绪才算 patch 完成。
  function slidePatchIncomplete() {
    var jR = window.AscFonts && window.AscFonts.jR;
    return (
      !SLIDE_CATALOG_READY ||
      !SLIDE_BINARIES_INJECTED ||
      !(jR && jR.__CATALOG_FONT_PATCHED__)
    );
  }

  // ---------------------------------------------------------------------------
  // 公共：registry / catalog 二进制 / 字体名工具
  // ---------------------------------------------------------------------------

  function plainFontName(name) {
    return name.indexOf("Embedded: ") === 0 ? name.slice(10) : name;
  }

  // Cell 引擎默认 uzg 指向 ../../../../fonts/，在 iframe 里会 404；必须从 AllFonts.js 路径推导。
  function getFontsBaseUrl() {
    if (window.__ONLYOFFICE_FONTS_BASE__) {
      return window.__ONLYOFFICE_FONTS_BASE__;
    }
    var scripts = document.getElementsByTagName("script");
    for (var i = 0; i < scripts.length; i++) {
      var src = scripts[i].src;
      if (!src || src.indexOf("AllFonts.js") < 0) {
        continue;
      }
      window.__ONLYOFFICE_FONTS_BASE__ = new URL("../../fonts/", src).href;
      return window.__ONLYOFFICE_FONTS_BASE__;
    }
    window.__ONLYOFFICE_FONTS_BASE__ = new URL(
      "../../../../fonts/",
      window.location.href
    ).href;
    return window.__ONLYOFFICE_FONTS_BASE__;
  }

  function loadCustomFontRegistry() {
    return window.__custom_font_registry__ || {};
  }

  function listFontNames(entry) {
    if (Array.isArray(entry)) {
      return entry;
    }
    if (entry && entry.aliases) {
      return entry.aliases;
    }
    return [];
  }

  function registryNames(entry) {
    var names = [];
    var seen = {};
    var candidates = listFontNames(entry);
    for (var i = 0; i < candidates.length; i++) {
      var name = candidates[i];
      if (!name || seen[name]) {
        continue;
      }
      seen[name] = true;
      names.push(name);
    }
    return names;
  }

  function expandRegistry(registry) {
    var ids = [];
    var names = [];
    var seen = {};

    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      ids.push(id);
      var entryNames = registryNames(registry[id] || {});
      for (var i = 0; i < entryNames.length; i++) {
        var name = entryNames[i];
        if (!seen[name]) {
          seen[name] = true;
          names.push(name);
        }
      }
    }

    return { ids: ids, names: names };
  }

  function decodeCatalogWire(wire) {
    var plain = new Uint8Array(wire);
    var n = Math.min(32, plain.length);
    for (var i = 0; i < n; i++) {
      plain[i] ^= XOR_KEY[i % XOR_KEY.length];
    }
    return plain;
  }

  function loadCatalogFontSync(fileId) {
    try {
      var url = getFontsBaseUrl() + fileId;
      var xhr = new XMLHttpRequest();
      // 同步 XHR 不能设 responseType（iframe 代理下会 InvalidAccessError），用 responseText 读二进制。
      xhr.open("GET", url, false);
      if (xhr.overrideMimeType) {
        xhr.overrideMimeType("text/plain; charset=x-user-defined");
      }
      xhr.send(null);
      if (xhr.status !== 200 && xhr.status !== 0) {
        return null;
      }
      var text = xhr.responseText || "";
      var buf = new Uint8Array(text.length);
      for (var i = 0; i < text.length; i++) {
        buf[i] = text.charCodeAt(i) & 255;
      }
      return decodeCatalogWire(buf);
    } catch (err) {
      return null;
    }
  }

  function findDpcIndex(dpc, fileId) {
    for (var i = 0; i < dpc.length; i++) {
      if (dpc[i] && dpc[i].Ua === fileId) {
        return i;
      }
    }
    return -1;
  }

  function findYycIndex(yyc, fileId) {
    for (var i = 0; i < yyc.length; i++) {
      if (yyc[i] && yyc[i].Va === fileId) {
        return i;
      }
    }
    return -1;
  }

  function findJecIndex(jec, fileId) {
    for (var i = 0; i < jec.length; i++) {
      if (jec[i] && jec[i].La === fileId) {
        return i;
      }
    }
    return -1;
  }

  function cellFontCatalogHasFamilies(kpb) {
    if (!kpb) {
      return false;
    }
    for (var name in kpb) {
      if (Object.prototype.hasOwnProperty.call(kpb, name)) {
        return true;
      }
    }
    return false;
  }

  function restoreFontCatalogArraysFromSnapshot() {
    var snap = window.__custom_font_catalog_snapshot__;
    if (!snap || !snap.files || !snap.infos || window.__fonts_files !== undefined) {
      return false;
    }
    window.__fonts_files = snap.files.slice();
    window.__fonts_infos = snap.infos.map(function (row) {
      return row.slice();
    });
    return true;
  }

  // Cell SDK 初始化时把 yyc/JPb/KPb 拷进 V_；仅改 AscFonts 不够，绘制仍读 V_ 上的旧引用。
  function syncCellEngineFontRefs() {
    var asc = window.AscFonts;
    var V_ = window.AscCommon && window.AscCommon.V_;
    if (!asc || !V_) {
      return;
    }
    var base = getFontsBaseUrl();
    V_.uzg = base;
    window.__ONLYOFFICE_FONTS_BASE__ = base;
    if (asc.yyc) {
      V_.Rkb = asc.yyc;
    }
    if (asc.JPb) {
      V_.KNd = asc.JPb;
    }
    if (asc.KPb) {
      V_.yEc = asc.KPb;
    }
  }

  // Word 先加载并 delete __fonts_files 后，Cell 需从快照调用 acj() 重建 yyc/JPb/KPb。
  function buildCellFontCatalogFromSnapshot() {
    var asc = window.AscFonts;
    if (!asc) {
      return false;
    }

    if (window.__fonts_files === undefined) {
      restoreFontCatalogArraysFromSnapshot();
    }

    if (window.__fonts_files !== undefined) {
      var inner = asc.__CUSTOM_FONT_ACJ_INNER__;
      if (typeof inner === "function") {
        inner();
      } else if (typeof asc.acj === "function") {
        asc.acj();
      }
    }

    syncCellEngineFontRefs();
    return !!(
      asc.yyc &&
      asc.JPb &&
      cellFontCatalogHasFamilies(asc.KPb)
    );
  }

  function ensureCellFontCatalogFromSnapshot() {
    var asc = window.AscFonts;
    if (
      asc &&
      asc.yyc &&
      asc.JPb &&
      cellFontCatalogHasFamilies(asc.KPb)
    ) {
      syncCellEngineFontRefs();
      return true;
    }
    return buildCellFontCatalogFromSnapshot();
  }

  // ---------------------------------------------------------------------------
  // Word：dpc / L1b / QQ.sgd / mJ.GPb
  // ---------------------------------------------------------------------------

  function registerAliasFamiliesWord(registry) {
    var l1b = window.AscFonts && window.AscFonts.L1b;
    if (!l1b) {
      return false;
    }

    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var entry = registry[id];
      var list = listFontNames(entry);
      var primary = list[0];
      var idx = primary !== undefined ? l1b[primary] : undefined;
      if (idx === undefined) {
        for (var p = 0; p < list.length; p++) {
          if (list[p] !== undefined && l1b[list[p]] !== undefined) {
            primary = list[p];
            idx = l1b[primary];
            break;
          }
        }
      }
      if (primary === undefined || idx === undefined) {
        continue;
      }
      for (var i = 0; i < list.length; i++) {
        var alias = list[i];
        if (alias && alias !== primary) {
          l1b[alias] = idx;
        }
      }
    }

    return true;
  }

  // ---------------------------------------------------------------------------
  // Excel：yyc / KPb / JPb / dW / V_.koc
  // ---------------------------------------------------------------------------

  function registerAliasFamiliesCell(registry) {
    var kpb = window.AscFonts && window.AscFonts.KPb;
    var dW = window.AscFonts && window.AscFonts.dW;
    if (!kpb) {
      return false;
    }

    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var entry = registry[id];
      var list = listFontNames(entry);
      var primary = list[0];
      if (!primary || kpb[primary] === undefined) {
        continue;
      }
      var idx = kpb[primary];
      for (var i = 1; i < list.length; i++) {
        var alias = list[i];
        if (alias && alias !== primary) {
          kpb[alias] = idx;
          if (dW && dW.X0b) {
            dW.X0b[alias] = primary;
          }
        }
      }
    }

    return true;
  }

  // x2t 可能把「仿宋」→「仿宋_GB2312」重复替换，产生 仿宋_GB2312_GB2312。
  function collapsePrimaryFontSuffix(name, primary) {
    if (!name || !primary || name === primary) {
      return name;
    }
    if (name.indexOf(primary) !== 0) {
      return name;
    }
    var tail = name.slice(primary.length);
    if (tail === "_GB2312" || /^(_GB2312)+$/.test(tail)) {
      return primary;
    }
    return name;
  }

  // 统一三端 workbook/样式/文档字体队列里的族名；Slide 的 vFb、hyb Proxy、jR.quc 均依赖此函数。
  function resolveDocumentFontName(name) {
    if (!name) {
      return name;
    }
    var kpb = window.AscFonts && window.AscFonts.KPb;
    var hyb = window.AscFonts && window.AscFonts.hyb;
    var l1b = window.AscFonts && window.AscFonts.L1b;
    if (kpb && Object.prototype.hasOwnProperty.call(kpb, name)) {
      return name;
    }
    if (hyb && Object.prototype.hasOwnProperty.call(hyb, name)) {
      return name;
    }
    if (l1b && Object.prototype.hasOwnProperty.call(l1b, name)) {
      return name;
    }

    var registry = loadCustomFontRegistry();
    for (var rid in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, rid)) {
        continue;
      }
      var rlist = listFontNames(registry[rid]);
      if (rlist[0]) {
        name = collapsePrimaryFontSuffix(name, rlist[0]);
      }
    }
    if (kpb && Object.prototype.hasOwnProperty.call(kpb, name)) {
      return name;
    }
    if (hyb && Object.prototype.hasOwnProperty.call(hyb, name)) {
      return name;
    }
    if (l1b && Object.prototype.hasOwnProperty.call(l1b, name)) {
      return name;
    }

    registry = loadCustomFontRegistry();
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var list = listFontNames(registry[id]);
      var primary = list[0];
      for (var i = 0; i < list.length; i++) {
        if (list[i] === name && primary) {
          return primary;
        }
      }
    }

    return name;
  }

  function registerDocumentFontAliasesCell(registry) {
    var kpb = window.AscFonts && window.AscFonts.KPb;
    var dW = window.AscFonts && window.AscFonts.dW;
    if (!kpb) {
      return false;
    }

    var primary = null;
    var idx;
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var list = listFontNames(registry[id]);
      if (list[0] && kpb[list[0]] !== undefined) {
        primary = list[0];
        idx = kpb[primary];
        break;
      }
    }
    if (primary === null || idx === undefined) {
      return false;
    }

    return true;
  }

  // 单元格读写 font 名时归一化，避免 X5b()/样式里带着重复后缀进 koc。
  function hookCellExcelFontBe() {
    var Excel = window.AscCommonExcel;
    if (!Excel || !Excel.Wm || Excel.Wm.__CUSTOM_FONT_BE_PATCHED__) {
      return false;
    }

    var origBe = Excel.Wm.prototype.be;
    Excel.Wm.prototype.be = function () {
      return resolveDocumentFontName(origBe.call(this));
    };

    var origHk = Excel.Wm.prototype.Hk;
    Excel.Wm.prototype.Hk = function (name) {
      return origHk.call(this, resolveDocumentFontName(name));
    };

    Excel.Wm.__CUSTOM_FONT_BE_PATCHED__ = true;
    return true;
  }

  // 从 workbook 收集实际用到的字体，去重并 resolve 后交给 fi.jyb / V_.koc。
  function buildWorkbookFontList(wd) {
    var raw = wd.X5b();
    var seen = {};
    var list = [];
    for (var name in raw) {
      var resolved = resolveDocumentFontName(name);
      if (!resolved || seen[resolved]) {
        continue;
      }
      seen[resolved] = true;
      list.push(new AscFonts.tJa(resolved));
    }
    return list;
  }

  // fi.PO 注册文档字体时也走 resolveDocumentFontName。
  function hookCellDocumentFontRegistration() {
    var fi = window.AscFonts && window.AscFonts.fi;
    if (!fi || !fi.PO || fi.__CUSTOM_FONTS_PO_PATCHED__) {
      return false;
    }

    var origPo = fi.PO.bind(fi);
    fi.PO = function (name) {
      return origPo(resolveDocumentFontName(name));
    };
    fi.__CUSTOM_FONTS_PO_PATCHED__ = true;
    return true;
  }

  function findFallbackThumbnailIndex(catalog) {
    for (var i = 0; i < catalog.length; i++) {
      if (catalog[i].za === "Arial") {
        return catalog[i].QSh;
      }
    }
    for (var j = 0; j < catalog.length; j++) {
      if (catalog[j].za === "Calibri") {
        return catalog[j].QSh;
      }
    }
    return 0;
  }

  var cachedFontSpriteCount = -2;

  function getFontThumbnailSpritePath() {
    var suffix = "";
    try {
      var lang = (navigator.language || "en").toLowerCase();
      if (/^(zh|ja|ko)/.test(lang)) {
        suffix = "_ea";
      }
    } catch (err) {
      suffix = "";
    }
    // web-apps ComboBoxFonts 加载的是 fonts_thumbnail[_ea].png.bin，不是 .bin。
    return (
      "../../../../sdkjs/common/Images/fonts_thumbnail" +
      suffix +
      ".png.bin"
    );
  }

  function getFontSpriteCountSync() {
    if (cachedFontSpriteCount !== -2) {
      return cachedFontSpriteCount;
    }
    cachedFontSpriteCount = -1;
    try {
      var xhr = new XMLHttpRequest();
      var url = new URL(getFontThumbnailSpritePath(), window.location.href).href;
      xhr.open("GET", url, false);
      if ("responseType" in xhr) {
        xhr.responseType = "arraybuffer";
      }
      xhr.send(null);
      if (!xhr.response || xhr.response.byteLength < 12) {
        return cachedFontSpriteCount;
      }
      var header = new Uint8Array(xhr.response);
      cachedFontSpriteCount =
        ((header[8] << 24) |
          (header[9] << 16) |
          (header[10] << 8) |
          header[11]) >>>
        0;
    } catch (err) {
      cachedFontSpriteCount = -1;
    }
    return cachedFontSpriteCount;
  }

  function buildRegistryNameSet(registry) {
    var set = {};
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var names = registryNames(registry[id] || {});
      for (var i = 0; i < names.length; i++) {
        set[names[i]] = true;
      }
    }
    return set;
  }

  function fixFontPickerThumbnails(registry) {
    var catalog = window.AscFonts && window.AscFonts.$4a;
    if (!catalog) {
      return false;
    }

    var fallback = findFallbackThumbnailIndex(catalog);
    var nameSet = buildRegistryNameSet(registry);
    for (var i = 0; i < catalog.length; i++) {
      if (nameSet[catalog[i].za]) {
        catalog[i].QSh = fallback;
      }
    }

    return true;
  }

  function getFontThumbnailValue(font) {
    if (!font) {
      return undefined;
    }
    if (typeof font.asc_getFontThumbnail === "function") {
      return font.asc_getFontThumbnail();
    }
    if (font.zPg !== undefined) {
      return font.zPg;
    }
    if (font.feh !== undefined) {
      return font.feh;
    }
    if (font.qAg !== undefined) {
      return font.qAg;
    }
    return undefined;
  }

  function setFontThumbnailValue(font, idx) {
    if (!font) {
      return;
    }
    if (font.zPg !== undefined) {
      font.zPg = idx;
    }
    if (font.feh !== undefined) {
      font.feh = idx;
    }
    if (font.qAg !== undefined) {
      font.qAg = idx;
    }
  }

  function getMaxKnownThumbnailIndex(catalog) {
    var max = 0;
    if (!catalog) {
      return max;
    }
    for (var i = 0; i < catalog.length; i++) {
      var q = catalog[i].QSh;
      if (typeof q === "number" && isFinite(q) && q >= 0 && q > max) {
        max = q;
      }
    }
    return max;
  }

  function fontThumbnailNeedsFallback(name, thumb, nameSet, maxKnown, spriteCount) {
    if (nameSet[name]) {
      return true;
    }
    if (typeof thumb !== "number" || !isFinite(thumb) || thumb < 0) {
      return true;
    }
    thumb = Math.floor(thumb);
    if (spriteCount > 0 && thumb >= spriteCount) {
      return true;
    }
    if (maxKnown >= 0 && thumb > maxKnown) {
      return true;
    }
    // 自定义字体二进制元数据常误写成单格字节数（300*28*4=33600），读不到 sprite 时也要拦。
    if (spriteCount <= 0 && thumb > 512) {
      return true;
    }
    return false;
  }

  function getThumbnailSanitizeContext(registry) {
    registry = registry || loadCustomFontRegistry();
    var spriteCount = getFontSpriteCountSync();
    var nameSet = buildRegistryNameSet(registry);
    var catalog = window.AscFonts && window.AscFonts.$4a;

    if (catalog) {
      fixFontPickerThumbnails(registry);
      var maxKnown = getMaxKnownThumbnailIndex(catalog);
      if (spriteCount > 0) {
        maxKnown = Math.min(maxKnown, spriteCount - 1);
      } else if (maxKnown > 512) {
        maxKnown = 512;
      }
      return {
        registry: registry,
        fallback: findFallbackThumbnailIndex(catalog),
        nameSet: nameSet,
        spriteCount: spriteCount,
        maxKnown: maxKnown,
      };
    }

    if (spriteCount > 0) {
      return {
        registry: registry,
        fallback: 0,
        nameSet: nameSet,
        spriteCount: spriteCount,
        maxKnown: spriteCount - 1,
      };
    }

    return {
      registry: registry,
      fallback: 0,
      nameSet: nameSet,
      spriteCount: -1,
      maxKnown: 512,
    };
  }

  function resolveSafeThumbnailIndex(name, rawThumb, ctx) {
    ctx = ctx || getThumbnailSanitizeContext();
    if (fontThumbnailNeedsFallback(name, rawThumb, ctx.nameSet, ctx.maxKnown, ctx.spriteCount)) {
      return ctx.fallback;
    }
    if (typeof rawThumb !== "number" || !isFinite(rawThumb)) {
      return ctx.fallback;
    }
    return Math.floor(rawThumb);
  }

  function sanitizeFontComboStore(store) {
    if (!store || typeof store.each !== "function") {
      return;
    }
    var ctx = getThumbnailSanitizeContext();
    store.each(function (model) {
      if (!model || typeof model.get !== "function") {
        return;
      }
      var name = model.get("name");
      var imgidx = model.get("imgidx");
      var safe = resolveSafeThumbnailIndex(name, imgidx, ctx);
      if (safe !== imgidx) {
        model.set("imgidx", safe);
      }
    });
  }

  // 首段 registry 补丁与本段完整编辑器补丁处于不同 IIFE 作用域；暴露一个
  // 极小入口，供前者在字体输入/复制前安全清理缩略图索引。
  window.__sanitizeCustomFontComboStore = sanitizeFontComboStore;

  function sanitizeFontListThumbnails(fonts, registry) {
    if (!fonts || !fonts.length) {
      return;
    }

    var ctx = getThumbnailSanitizeContext(registry);
    for (var j = 0; j < fonts.length; j++) {
      var font = fonts[j];
      if (!font) {
        continue;
      }
      var name =
        font.name ||
        (typeof font.asc_getFontName === "function" ? font.asc_getFontName() : "");
      var thumb = getFontThumbnailValue(font);
      var safe = resolveSafeThumbnailIndex(name, thumb, ctx);
      if (safe !== thumb) {
        setFontThumbnailValue(font, safe);
      }
    }
  }

  function sanitizeFontCollectionModels(collection) {
    if (!collection) {
      return;
    }
    var ctx = getThumbnailSanitizeContext();
    if (typeof collection.each === "function") {
      collection.each(function (model) {
        if (!model || typeof model.get !== "function") {
          return;
        }
        var name = model.get("name");
        var imgidx = model.get("imgidx");
        var safe = resolveSafeThumbnailIndex(name, imgidx, ctx);
        if (safe !== imgidx) {
          model.set("imgidx", safe);
        }
      });
      return;
    }
    if (Array.isArray(collection)) {
      for (var i = 0; i < collection.length; i++) {
        var row = collection[i];
        if (!row) {
          continue;
        }
        var rowName = row.name;
        var rowIdx = row.imgidx;
        var rowSafe = resolveSafeThumbnailIndex(rowName, rowIdx, ctx);
        if (rowSafe !== rowIdx) {
          row.imgidx = rowSafe;
        }
      }
    }
  }

  function hookFontObjectThumbnailProto(Ctor, prop) {
    if (!Ctor || !Ctor.prototype || Ctor.prototype.__CUSTOM_FONT_THUMB_PROTO__) {
      return false;
    }
    var proto = Ctor.prototype;
    var origGet = proto.asc_getFontThumbnail;
    proto.asc_getFontThumbnail = function () {
      var name = "";
      if (typeof this.asc_getFontName === "function") {
        name = this.asc_getFontName();
      } else if (this.name) {
        name = this.name;
      }
      var raw = origGet ? origGet.call(this) : this[prop];
      return resolveSafeThumbnailIndex(name, raw);
    };
    proto.__CUSTOM_FONT_THUMB_PROTO__ = true;
    return true;
  }

  function hookFontObjectThumbnails() {
    var asc = window.AscFonts;
    if (!asc) {
      return false;
    }
    var patched = false;
    patched = hookFontObjectThumbnailProto(asc.tJa, "feh") || patched;
    patched = hookFontObjectThumbnailProto(asc.nbb, "qAg") || patched;
    patched = hookFontObjectThumbnailProto(asc.aKa, "zPg") || patched;
    return patched;
  }

  /**
   * Cell / Slide 9.4 使用 AscFonts.pickFont() 返回实际字体文件描述。它们的
   * 内置 catalog 不会从后加的 __fonts_infos 自动生成名称索引，未知自定义名会
   * 被解析为 Arial（文件 022）。保留原对象原型，仅将文件 id 改为 registry id，
   * 让原生加载器继续按 fonts/{id} 拉取并缓存字形。
   */
  function hookRuntimeCustomFontPicker() {
    var asc = window.AscFonts;
    if (!asc || typeof asc.pickFont !== "function" || asc.__CUSTOM_PICK_FONT_PATCHED__) {
      return false;
    }

    var original = asc.pickFont.bind(asc);
    asc.pickFont = function (name) {
      var file = original.apply(this, arguments);
      var registry = loadCustomFontRegistry();
      var id = null;
      for (var key in registry) {
        if (!Object.prototype.hasOwnProperty.call(registry, key)) {
          continue;
        }
        var names = listFontNames(registry[key]);
        if (names.indexOf(name) >= 0) {
          id = key;
          break;
        }
      }
      if (!id || !file) {
        return file;
      }

      var resolved = Object.create(Object.getPrototypeOf(file));
      for (var prop in file) {
        resolved[prop] = file[prop];
      }
      // Cell uses Xa; Slide uses Na. Keep both variants for minor SDK builds.
      if ("Xa" in resolved) resolved.Xa = id;
      if ("Na" in resolved) resolved.Na = id;
      return resolved;
    };
    asc.__CUSTOM_PICK_FONT_PATCHED__ = true;
    return true;
  }

  function patchEditorFontListMethod(editor, methodName) {
    if (!editor) {
      return false;
    }
    var orig = editor[methodName];
    if (typeof orig !== "function") {
      var proto = Object.getPrototypeOf(editor);
      while (proto && typeof orig !== "function") {
        orig = proto[methodName];
        proto = Object.getPrototypeOf(proto);
      }
    }
    if (typeof orig !== "function") {
      return false;
    }
    var flag = "__CUSTOM_FONTS_" + methodName + "_PATCHED__";
    if (editor[flag]) {
      return false;
    }
    editor[methodName] = function (fonts) {
      sanitizeFontListThumbnails(fonts, loadCustomFontRegistry());
      return orig.apply(this, arguments);
    };
    editor[flag] = true;
    return true;
  }

  function patchEditorEventMethod(editor, methodName) {
    if (!editor) {
      return false;
    }
    var orig = editor[methodName];
    if (typeof orig !== "function") {
      var proto = Object.getPrototypeOf(editor);
      while (proto && typeof orig !== "function") {
        orig = proto[methodName];
        proto = Object.getPrototypeOf(proto);
      }
    }
    if (typeof orig !== "function") {
      return false;
    }
    var flag = "__CUSTOM_FONTS_" + methodName + "_PATCHED__";
    if (editor[flag]) {
      return false;
    }
    editor[methodName] = function (eventName) {
      if (eventName === "asc_onInitEditorFonts" && arguments[1]) {
        sanitizeFontListThumbnails(arguments[1], loadCustomFontRegistry());
      }
      return orig.apply(this, arguments);
    };
    editor[flag] = true;
    return true;
  }

  function hookEditorFontEventEmit() {
    var editor = window.Asc && window.Asc.editor;
    if (!editor) {
      return false;
    }
    var patched = false;
    patched = patchEditorEventMethod(editor, "fe") || patched;
    patched = patchEditorEventMethod(editor, "qc") || patched;
    return patched;
  }

  // web-apps 在 asc_onInitEditorFonts 里把 imgidx 写入 Backbone store；点选字体复制到「最近使用」会 clone 该值。
  function hookWebAppsFontsLoad() {
    var common = window.Common;
    if (
      !common ||
      !common.NotificationCenter ||
      common.NotificationCenter.__CUSTOM_FONTS_LOAD_PATCHED__
    ) {
      return false;
    }
    var nc = common.NotificationCenter;
    var orig = nc.trigger;
    nc.trigger = function (eventName) {
      if (eventName === "fonts:load" && arguments.length > 1) {
        sanitizeFontCollectionModels(arguments[1]);
      }
      return orig.apply(this, arguments);
    };
    nc.__CUSTOM_FONTS_LOAD_PATCHED__ = true;
    return true;
  }

  // 最后一道防线：在 web-apps ComboBoxFonts 渲染/筛选/复制到最近使用前修正 store.imgidx。
  function hookComboBoxFontsWebApps() {
    var common = window.Common;
    if (!common || !common.UI || !common.UI.ComboBoxFonts) {
      return false;
    }
    var proto = common.UI.ComboBoxFonts.prototype;
    if (proto.__CUSTOM_FONTS_COMBO_PATCHED__) {
      return false;
    }

    if (typeof proto.updateVisibleFontsTiles === "function") {
      var origTiles = proto.updateVisibleFontsTiles;
      proto.updateVisibleFontsTiles = function (t, e) {
        sanitizeFontComboStore(this.store);
        return origTiles.call(this, t, e);
      };
    }

    if (typeof proto.fillFonts === "function") {
      var origFill = proto.fillFonts;
      proto.fillFonts = function (t, e) {
        sanitizeFontCollectionModels(t);
        return origFill.call(this, t, e);
      };
    }

    // 9.4 的 ComboBoxFonts 会在首次 fillFonts 后缓存 _fontsArray；后续由
    // SDK 注册进 store 的自定义字体不会自动更新这份搜索索引，导致下拉项
    // 已存在却无法通过输入框匹配。搜索前从当前 store 重建缓存。
    if (typeof proto.selectCandidate === "function") {
      var origSelectCandidate = proto.selectCandidate;
      proto.selectCandidate = function (t) {
        if (this.store && typeof this.store.toJSON === "function") {
          this._fontsArray = this.store.toJSON();
        }
        return origSelectCandidate.call(this, t);
      };
    }

    if (typeof proto.addItemToRecent === "function") {
      var origRecent = proto.addItemToRecent;
      proto.addItemToRecent = function (t, e) {
        if (t && typeof t.get === "function") {
          var name = t.get("name");
          var imgidx = t.get("imgidx");
          var safe = resolveSafeThumbnailIndex(name, imgidx);
          if (safe !== imgidx) {
            t.set("imgidx", safe);
          }
        }
        var result = origRecent.call(this, t, e);
        sanitizeFontComboStore(this.store);
        return result;
      };
    }

    proto.__CUSTOM_FONTS_COMBO_PATCHED__ = true;
    return true;
  }

  function buildToolbarFontArrayFromCatalog() {
    var asc = window.AscFonts;
    if (!asc) {
      return null;
    }
    var Ctor = asc.tJa || asc.nbb || asc.aKa;
    if (!Ctor) {
      return null;
    }

    // 9.3 使用 $4a（{ za, QSh }）；9.4 Word 改为 Vbb（{ xa, a9 }）。
    // 两者都必须转换为 asc_onInitEditorFonts 所期待的 CFontInfo 列表，
    // 否则 web-apps 的 ComboBoxFonts 不会重建其 Backbone store。
    var catalog = asc.$4a;
    var isV94Catalog = !Array.isArray(catalog) && Array.isArray(asc.Vbb);
    if (isV94Catalog) {
      catalog = asc.Vbb;
    }
    if (!Array.isArray(catalog)) {
      return null;
    }

    var fonts = [];
    var seen = {};
    for (var i = 0; i < catalog.length; i++) {
      var entry = catalog[i];
      var name = isV94Catalog ? entry && entry.xa : entry && entry.za;
      if (!name || name === "ASCW3" || seen[name]) {
        continue;
      }
      seen[name] = true;
      // 9.4 tJa 从字体名构建 CFontInfo；缩略图统一在下方做安全兜底，
      // 不可把 Vbb.a9（字体文件索引）误当作 thumbnail 索引传入。
      fonts.push(new Ctor(name));
    }
    return fonts.length ? fonts : null;
  }

  // catalog/缩略图补丁生效后，强制刷新 toolbar 字体列表（修复 ock/Igj 早于 hook 的 imgidx）。
  function reloadEditorFontListForToolbar() {
    hookFontObjectThumbnails();
    hookRuntimeCustomFontPicker();
    hookFontListInit();
    hookRuntimeCustomFontPicker();
    hookEditorFontEventEmit();
    hookWebAppsFontsLoad();
    hookComboBoxFontsWebApps();
    var fonts = buildToolbarFontArrayFromCatalog();
    if (!fonts) {
      return false;
    }
    sanitizeFontListThumbnails(fonts, loadCustomFontRegistry());
    var editor = window.Asc && window.Asc.editor;
    if (!editor) {
      return false;
    }
    if (typeof editor.Igj === "function") {
      editor.Igj(fonts);
      return true;
    }
    if (typeof editor.ock === "function") {
      editor.ock(fonts);
      return true;
    }
    if (typeof editor.Zwj === "function") {
      editor.Zwj(fonts);
      return true;
    }
    return false;
  }

  function scheduleToolbarFontListReload() {
    window.setTimeout(function () {
      reloadEditorFontListForToolbar();
    }, 0);
  }

  // Word: Zwj；Excel: ock；PPT: Igj — 三端把字体列表交给 web-apps 的入口不同。
  function hookFontListInit() {
    var editor = window.Asc && window.Asc.editor;
    if (!editor) {
      return false;
    }
    hookFontObjectThumbnails();
    hookEditorFontEventEmit();
    hookWebAppsFontsLoad();
    hookComboBoxFontsWebApps();
    var patched = false;
    patched = patchEditorFontListMethod(editor, "Zwj") || patched;
    patched = patchEditorFontListMethod(editor, "ock") || patched;
    patched = patchEditorFontListMethod(editor, "Igj") || patched;
    return patched;
  }

  function hookSlideFontDelivery() {
    var lU = window.AscCommon && window.AscCommon.lU;
    if (!lU || lU.__CUSTOM_FONTS_QDB_PATCHED__ || !lU.qdb) {
      return false;
    }
    var origQdb = lU.qdb.bind(lU);
    lU.qdb = function (fonts) {
      sanitizeFontListThumbnails(fonts, loadCustomFontRegistry());
      return origQdb.apply(this, arguments);
    };
    lU.__CUSTOM_FONTS_QDB_PATCHED__ = true;
    return true;
  }

  function hookEditorForFontList() {
    var asc = (window.Asc = window.Asc || {});
    if (asc.__CUSTOM_FONTS_EDITOR_WATCH__) {
      return false;
    }

    asc.__CUSTOM_FONTS_EDITOR_WATCH__ = true;
    var editor = asc.editor;
    Object.defineProperty(asc, "editor", {
      configurable: true,
      get: function () {
        return editor;
      },
      set: function (value) {
        editor = value;
        hookFontListInit();
        scheduleToolbarFontListReload();
        if (
          value &&
          !isSpreadsheetEditor(value) &&
          !isPresentationEditor(value)
        ) {
          window.setTimeout(function () {
            reloadWordDocumentFontsFromEditor();
          }, 0);
        }
        // 演示稿 editor 就绪后：挂 JOf hook，并等待 ea.N_ 可用后 reload 一次 TJb。
        if (PATCHED_SLIDE && value && value.yga) {
          syncSlideFontBaseUrl();
          hookSlideEditorJOfOnce(value);
          scheduleSlideFontReloadOnceWhenReady();
        }
      },
    });
    hookFontListInit();
    if (PATCHED_SLIDE && asc.editor && asc.editor.yga) {
      syncSlideFontBaseUrl();
      hookSlideEditorJOfOnce(asc.editor);
      scheduleSlideFontReloadOnceWhenReady();
    }
    return true;
  }

  // --- Word 安装 ---

  // Word FontPicker（QQ.sgd）：catalog 已注册字体直接返回，不走模糊匹配。
  function installWordCatalogFontResolution() {
    var qq = window.AscFonts && window.AscFonts.QQ;
    var l1b = window.AscFonts && window.AscFonts.L1b;
    if (!qq || !qq.sgd || !l1b || qq.__CATALOG_FONT_PATCHED__) {
      return false;
    }

    var origSgd = qq.sgd.bind(qq);
    qq.sgd = function (name) {
      var family = resolveDocumentFontName(plainFontName(name));
      if (Object.prototype.hasOwnProperty.call(l1b, family)) {
        var picked = { Yda: family };
        if (qq.n3d) {
          qq.n3d[name] = picked;
          if (family !== name) {
            qq.n3d[family] = picked;
          }
        }
        return picked;
      }
      return origSgd(name);
    };
    qq.__CATALOG_FONT_PATCHED__ = true;
    return true;
  }

  // ---------------------------------------------------------------------------
  // Slide（PPT）：jec / hyb / jR.quc / AscCommon.lU.vFb + s7e + yga.TJb
  //
  // SDK 打开演示稿时序（sdk-all-min.js）：
  //   ra.yd.INd()  → 收集 ea.N_（文档用到的字体）
  //   yga.TJb(ea.N_) → vFb 入队 → s7e 异步加载 → ZDc/JOf 完成
  //
  // 自定义字体需拦截三处：
  //   1. jR.quc/EE — 避免 FontPicker 把「仿宋_GB2312」模糊成 Arial
  //   2. hyb Proxy — 画布 font(t) 直接 hyb[t]，不经过 quc
  //   3. inject kec — 1001 不在默认包内，须 sync XHR 注入二进制
  //
  // 勿在 prepare 内调用 queueDocumentFontsSlide：vFb hook 会 ensure→prepare 递归爆栈。
  // 勿调用 ra.yd.WN() 做全量缩略图刷新：会导致左侧幻灯片栏消失。
  // ---------------------------------------------------------------------------

  // Slide SDK 赋值 uji 后直接调用 h()，不会走 AscFonts.uji()；需在赋值后 setTimeout 补跑 patch。
  function hookSlideFontCatalogInit() {
    var asc = (window.AscFonts = window.AscFonts || {});
    if (asc.__CUSTOM_FONTS_UJI_WATCH__) {
      return false;
    }

    asc.__CUSTOM_FONTS_UJI_WATCH__ = true;
    var wrapped = asc.uji;
    Object.defineProperty(asc, "uji", {
      configurable: true,
      get: function () {
        return wrapped;
      },
      set: function (fn) {
        if (!fn || fn.__CUSTOM_FONTS_UJI_WRAPPED__) {
          wrapped = fn;
          return;
        }
        var inner = fn;
        wrapped = inner;
        // Slide SDK 赋值后直接 h()，不会走 setter 返回的 wrapper；与 Cell acj 一样用 setTimeout 补跑。
        window.setTimeout(runSlideFontPipelineAfterCatalogInit, 0);
      },
    });
    return true;
  }

  // uji 赋值 = Slide catalog 初始化（h()）；setTimeout 补跑 pipeline（赋值后 SDK 直接调 h()）。
  function runSlideFontPipelineAfterCatalogInit() {
    if (
      !window.AscFonts ||
      !window.AscFonts.jec ||
      !window.AscFonts.hyb ||
      !window.AscFonts.jR
    ) {
      return false;
    }

    syncSlideFontBaseUrl();
    hookSlideDocumentFontLoading();
    hookSlideFontDelivery();
    hookSlideFontPickerGl();
    ensureSlideFontCatalogOnce();
    PATCHED_SLIDE = true;
    scheduleToolbarFontListReload();
    scheduleSlideFontReloadOnceWhenReady();
    return true;
  }

  // AllFonts.js 先于 slide sdk 执行；监听 lU 赋值以便在首次 yga.TJb 前 patch vFb/TJb。
  function hookSlideFontManagerInit() {
    var common = (window.AscCommon = window.AscCommon || {});
    if (common.__CUSTOM_FONTS_LU_WATCH__) {
      return false;
    }

    common.__CUSTOM_FONTS_LU_WATCH__ = true;
    var current = common.lU;
    Object.defineProperty(common, "lU", {
      configurable: true,
      enumerable: true,
      get: function () {
        return current;
      },
      set: function (value) {
        current = value;
        hookSlideDocumentFontLoading();
      },
    });
    hookSlideDocumentFontLoading();
    hookSlideFontDelivery();
    return true;
  }

  // SDK 重建 FontPicker 索引（Gl）后补注册 hyb 别名并清 sRd 缓存。
  function hookSlideFontPickerGl() {
    var jR = window.AscFonts && window.AscFonts.jR;
    if (!jR || jR.__CUSTOM_FONTS_GL_PATCHED__ || !jR.Gl) {
      return false;
    }

    var origGl = jR.Gl.bind(jR);
    jR.Gl = function () {
      origGl();
      var registry = loadCustomFontRegistry();
      fixFontPickerThumbnails(registry);
      if (SLIDE_CATALOG_READY) {
        registerAliasFamiliesSlide(registry);
        clearSlideFontPickerCache(jR);
      }
    };
    jR.__CUSTOM_FONTS_GL_PATCHED__ = true;
    return true;
  }

  // 字体异步加载完成回调；reload 后的 JOf 里只做一次 Me+Rp 重排主画布。
  function hookSlideEditorJOfOnce(editor) {
    if (!editor || editor.__CUSTOM_FONTS_JOF_ONCE__ || !isPresentationEditor(editor)) {
      return false;
    }

    if (!editor.JOf) {
      return false;
    }

    var origJOf = editor.JOf.bind(editor);
    editor.JOf = function (reason) {
      var result = origJOf(reason);
      // 仅在自定义 reload 完成后重排一次主画布；不调 yd.WN()，避免缩略图栏消失。
      if (!SLIDE_FONTS_RELOADED || SLIDE_LAYOUT_REFRESHED) {
        return result;
      }
      window.setTimeout(function () {
        if (!isPresentationEditor(editor) || SLIDE_LAYOUT_REFRESHED) {
          return;
        }
        SLIDE_LAYOUT_REFRESHED = true;
        var ra = editor.ra;
        if (ra && ra.Ea && ra.Ea.Me) {
          ra.Ea.Me();
        }
        if (ra && ra.Rp) {
          ra.Rp();
        }
      }, 0);
      return result;
    };
    editor.__CUSTOM_FONTS_JOF_ONCE__ = true;
    return true;
  }

  // ea.N_ 由 INd() 写入，可能是数组或类数组；打开文档前为空。
  function slideDocumentFontListReady(ea) {
    if (!ea || !ea.N_) {
      return false;
    }
    if (typeof ea.N_.length === "number" && ea.N_.length > 0) {
      return true;
    }
    for (var key in ea.N_) {
      if (Object.prototype.hasOwnProperty.call(ea.N_, key)) {
        return true;
      }
    }
    return false;
  }

  // 演示稿 INd 完成后 ea.N_ 才可用；轮询最多 10s。
  function waitForSlideDocumentFontsReady(callback) {
    if (!callback) {
      return;
    }
    var tries = 0;
    function attempt() {
      var editor = window.Asc && window.Asc.editor;
      var ea = editor && editor.ra && editor.ra.Ea;
      if (isPresentationEditor(editor) && slideDocumentFontListReady(ea)) {
        callback();
        return;
      }
      if (++tries < 200) {
        window.setTimeout(attempt, 50);
      }
    }
    attempt();
  }

  // 类比 Excel reloadCellDocumentFontsFromWorkbook：patch 就绪后强制重跑 TJb(ea.N_)。
  // SLIDE_FONTS_RELOADED 保证只 reload 一次，避免 TJb→JOf→reload 死循环。
  function scheduleSlideFontReloadOnceWhenReady() {
    if (SLIDE_FONTS_RELOADED || SLIDE_RELOAD_SCHEDULED) {
      return;
    }
    SLIDE_RELOAD_SCHEDULED = true;

    waitForSlideDocumentFontsReady(function () {
      var tries = 0;
      function attempt() {
        if (SLIDE_FONTS_RELOADED) {
          return;
        }
        if (!SLIDE_BINARIES_INJECTED || !SLIDE_CATALOG_READY) {
          ensureSlideFontCatalogOnce();
          if (++tries < 120) {
            window.setTimeout(attempt, 50);
          }
          return;
        }
        if (!canReloadSlideDocumentFonts()) {
          if (++tries < 120) {
            window.setTimeout(attempt, 50);
          }
          return;
        }

        SLIDE_FONTS_RELOADED = true;
        reloadSlideDocumentFontsFromPresentation();
      }

      attempt();
    });
  }

  // 保留旧名供调试；内部已改为 always-once reload。
  function scheduleSlideFontReloadOnceIfNeeded() {
    scheduleSlideFontReloadOnceWhenReady();
  }

  // 用演示稿实际字体列表重载：yga.TJb(ea.N_) → vFb → s7e → JOf。
  function reloadSlideDocumentFontsFromPresentation() {
    if (!canReloadSlideDocumentFonts()) {
      return false;
    }
    var editor = window.Asc && window.Asc.editor;
    var ea = editor && editor.ra && editor.ra.Ea;
    var yga = editor && editor.yga;
    if (!ea || !yga || !yga.TJb) {
      return false;
    }
    ensureSlideFontCatalogOnce();
    yga.TJb(ea.N_);
    return true;
  }

  function ensureSlideFontCatalogOnce() {
    if (SLIDE_CATALOG_READY) {
      syncSlideFontBaseUrl();
      return true;
    }
    if (!prepareSlideFontCatalog()) {
      return false;
    }
    SLIDE_CATALOG_READY = true;
    return true;
  }

  // 别名写入须走 hyb 底层对象；Proxy 只负责读路径（画布 font(t)）。
  function getSlideHybTarget() {
    var hyb = window.AscFonts && window.AscFonts.hyb;
    if (!hyb) {
      return null;
    }
    return hyb.__CUSTOM_FONTS_HYB_TARGET__ || hyb;
  }

  // 把 registry 里所有别名指向同一 hyb 下标（如 演示佛系体 → 仿宋_GB2312 的 idx）。
  function registerAliasFamiliesSlide(registry) {
    var hyb = getSlideHybTarget();
    if (!hyb) {
      return false;
    }

    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      var entry = registry[id];
      var list = listFontNames(entry);
      var idx;
      var anchor = null;
      for (var j = 0; j < list.length; j++) {
        var candidate = list[j];
        if (candidate && hyb[candidate] !== undefined) {
          anchor = candidate;
          idx = hyb[candidate];
          break;
        }
      }
      if (anchor === null || idx === undefined) {
        continue;
      }
      for (var i = 0; i < list.length; i++) {
        var alias = list[i];
        if (alias) {
          hyb[alias] = idx;
        }
      }
    }

    return true;
  }

  // Slide 画布 font(t) 直接读 hyb[t]，不经过 jR.quc；用 Proxy 兜底别名/后缀归一化。
  function installSlideHybLookupProxy() {
    var asc = window.AscFonts;
    if (!asc || !asc.hyb || asc.hyb.__CUSTOM_FONTS_HYB_PROXY__) {
      return false;
    }

    var target = asc.hyb;
    asc.hyb = new Proxy(target, {
      get: function (obj, prop) {
        if (typeof prop !== "string") {
          return obj[prop];
        }
        if (Object.prototype.hasOwnProperty.call(obj, prop)) {
          return obj[prop];
        }
        var resolved = resolveDocumentFontName(plainFontName(prop));
        if (
          resolved &&
          resolved !== prop &&
          Object.prototype.hasOwnProperty.call(obj, resolved)
        ) {
          return obj[resolved];
        }
        return obj[prop];
      },
      set: function (obj, prop, value) {
        obj[prop] = value;
        return true;
      },
    });
    asc.hyb.__CUSTOM_FONTS_HYB_PROXY__ = true;
    asc.hyb.__CUSTOM_FONTS_HYB_TARGET__ = target;
    return true;
  }

  function clearSlideFontPickerCache(jR) {
    if (jR && jR.sRd) {
      jR.sRd = {};
    }
  }

  // lU.q_f / editor.yga.q_f 默认 ../../../../fonts/ 在 iframe 内 404，须与 getFontsBaseUrl 对齐。
  function syncSlideFontBaseUrl() {
    var base = getFontsBaseUrl();
    var lU = window.AscCommon && window.AscCommon.lU;
    if (lU) {
      lU.q_f = base;
    }
    var editor = window.Asc && window.Asc.editor;
    if (editor && editor.yga) {
      editor.yga.q_f = base;
    }
    return base;
  }

  function resolveSlideCatalogEntry(name) {
    var resolved = resolveDocumentFontName(plainFontName(name));
    var hyb = window.AscFonts && window.AscFonts.hyb;
    var cSb = window.AscFonts && window.AscFonts.cSb;
    if (
      !resolved ||
      !hyb ||
      !cSb ||
      !Object.prototype.hasOwnProperty.call(hyb, resolved)
    ) {
      return null;
    }
    var idx = hyb[resolved];
    var entry = cSb[idx];
    if (!entry) {
      return null;
    }
    return { name: resolved, idx: idx, entry: entry };
  }

  // vFb 内部调用 jR.EE；patch EE 后 TJb 队列里的字体才会命中自定义 catalog。
  function installSlideCatalogFontResolution() {
    var jR = window.AscFonts && window.AscFonts.jR;
    var hyb = window.AscFonts && window.AscFonts.hyb;
    var cSb = window.AscFonts && window.AscFonts.cSb;
    if (!jR || !jR.quc || !hyb || !cSb || jR.__CATALOG_FONT_PATCHED__) {
      return false;
    }

    var origQuc = jR.quc.bind(jR);
    jR.quc = function (name) {
      var family = resolveDocumentFontName(plainFontName(name));
      if (Object.prototype.hasOwnProperty.call(hyb, family)) {
        var picked = { II: family, N6: family };
        if (jR.sRd) {
          jR.sRd[name] = picked;
          if (family !== name) {
            jR.sRd[family] = picked;
          }
        }
        return picked;
      }
      return origQuc(name);
    };

    if (jR.EE) {
      var origEE = jR.EE.bind(jR);
      jR.EE = function (name, style, out) {
        var family = resolveDocumentFontName(plainFontName(name));
        if (Object.prototype.hasOwnProperty.call(hyb, family)) {
          if (out) {
            out.ua = family;
            if (jR.G7e) {
              out.cO = jR.G7e(name, out);
            }
          }
          return cSb[hyb[family]];
        }
        return origEE(name, style, out);
      };
    }

    clearSlideFontPickerCache(jR);
    jR.__CATALOG_FONT_PATCHED__ = true;
    return true;
  }

  // 同步 XHR 读 public/fonts/{id}，写入 kec 并 HMb 绑定 jec；避免走异步 q_f 404。
  function injectCustomFontBinariesSlide(ids) {
    var asc = window.AscFonts;
    if (!asc || !asc.jec || !asc.kec || !asc.$2b) {
      return false;
    }
    if (SLIDE_BINARIES_INJECTED) {
      return true;
    }

    var injected = false;
    for (var i = 0; i < ids.length; i++) {
      var fileId = ids[i];
      var fileIndex = findJecIndex(asc.jec, fileId);
      if (fileIndex < 0) {
        continue;
      }

      var entry = asc.jec[fileIndex];
      if (!entry || (entry.ODa && entry.ODa())) {
        continue;
      }

      var data = loadCatalogFontSync(fileId);
      if (!data) {
        continue;
      }

      var streamIndex = asc.kec.length;
      asc.kec.push(new asc.$2b(data, data.length));
      if (entry.HMb) {
        entry.HMb(streamIndex);
      }
      entry.mz = 0;
      if (asc.cPd) {
        asc.cPd(streamIndex);
      }
      injected = true;
    }

    if (injected) {
      SLIDE_BINARIES_INJECTED = true;
    }
    return injected;
  }

  // Slide catalog 一次性安装：别名 → hyb Proxy → quc/EE → q_f → 二进制。
  // 禁止在此调用 queueDocumentFontsSlide（会与 vFb hook 互相递归）。
  function prepareSlideFontCatalog() {
    if (SLIDE_CATALOG_PREPARING) {
      return !!(
        window.AscFonts &&
        window.AscFonts.jR &&
        window.AscFonts.jR.__CATALOG_FONT_PATCHED__
      );
    }
    if (
      !window.AscFonts ||
      !window.AscFonts.jec ||
      !window.AscFonts.hyb ||
      !window.AscFonts.jR
    ) {
      return false;
    }

    SLIDE_CATALOG_PREPARING = true;
    try {
      var registry = loadCustomFontRegistry();
      fixFontPickerThumbnails(registry);
      hookFontListInit();
      hookSlideFontDelivery();
      var expanded = expandRegistry(registry);
      registerAliasFamiliesSlide(registry);
      installSlideHybLookupProxy();
      installSlideCatalogFontResolution();
      syncSlideFontBaseUrl();
      clearSlideFontPickerCache(window.AscFonts.jR);
      injectCustomFontBinariesSlide(expanded.ids);
      scheduleToolbarFontListReload();
      return true;
    } finally {
      SLIDE_CATALOG_PREPARING = false;
    }
  }

  function canReloadSlideDocumentFonts() {
    var editor = window.Asc && window.Asc.editor;
    return !!(
      isPresentationEditor(editor) &&
      editor.ra &&
      editor.ra.Ea &&
      slideDocumentFontListReady(editor.ra.Ea) &&
      editor.yga &&
      editor.yga.TJb
    );
  }

  // 独立 queue 须用 lU.__CUSTOM_FONTS_ORIG_VFB__，不可在 prepare 阶段调用。
  function queueDocumentFontsSlide(names) {
    var editor = window.Asc && window.Asc.editor;
    var yga = (editor && editor.yga) || (window.AscCommon && window.AscCommon.lU);
    if (!yga || !yga.vFb || !names.length) {
      return;
    }
    for (var i = 0; i < names.length; i++) {
      yga.vFb(names[i]);
    }
    if (yga.s7e) {
      yga.s7e();
    }
  }

  // 拦截文档字体入队与批量 TJb：每次 vFb/TJb 前确保 catalog 已 patch。
  function hookSlideDocumentFontLoading() {
    var lU = window.AscCommon && window.AscCommon.lU;
    if (!lU || lU.__CUSTOM_FONTS_DOC_FONTS_PATCHED__ || !lU.vFb) {
      return false;
    }

    var origVFb = lU.vFb.bind(lU);
    lU.__CUSTOM_FONTS_ORIG_VFB__ = origVFb;
    lU.vFb = function (name, priority) {
      syncSlideFontBaseUrl();
      // prepare 进行中不能再 ensure，否则 prepare→queue→vFb→ensure 栈溢出。
      if (!SLIDE_CATALOG_PREPARING) {
        ensureSlideFontCatalogOnce();
      }
      var resolved = resolveDocumentFontName(plainFontName(name));
      return origVFb(resolved || name, priority);
    };

    if (lU.TJb) {
      var origTJb = lU.TJb.bind(lU);
      // SDK 打开文档：INd() 后立即 TJb(ea.N_)；须在本函数返回前完成 patch。
      lU.TJb = function (fonts) {
        syncSlideFontBaseUrl();
        var wasIncomplete = slidePatchIncomplete();
        ensureSlideFontCatalogOnce();
        if (wasIncomplete && slidePatchIncomplete()) {
          SLIDE_UNPATCHED_TJB = true;
        }
        return origTJb(fonts);
      };
    }

    lU.__CUSTOM_FONTS_DOC_FONTS_PATCHED__ = true;
    return true;
  }

  // --- Excel FontPicker / 二进制 ---

  function getCellCanonicalName(family, kpb, jpb) {
    if (!family || !kpb || !Object.prototype.hasOwnProperty.call(kpb, family)) {
      return family;
    }
    var idx = kpb[family];
    var entry = jpb && jpb[idx];
    return entry && entry.Fa ? entry.Fa : family;
  }

  function clearCellFontPickerCache(dW) {
    // dW.oge 缓存 FontPicker 结果；catalog/二进制变更后必须清空，否则仍命中 Arial。
    if (dW && dW.oge) {
      dW.oge = {};
    }
  }

  function resolveCellCatalogFamily(family) {
    family = resolveDocumentFontName(family);
    var asc = window.AscFonts;
    return getCellCanonicalName(
      family,
      asc && asc.KPb,
      asc && asc.JPb
    );
  }

  function resolveCellCatalogEntry(name) {
    var resolved = resolveCellCatalogFamily(plainFontName(name));
    var kpb = window.AscFonts && window.AscFonts.KPb;
    var jpb = window.AscFonts && window.AscFonts.JPb;
    if (
      !resolved ||
      !kpb ||
      !jpb ||
      !Object.prototype.hasOwnProperty.call(kpb, resolved)
    ) {
      return null;
    }
    var idx = kpb[resolved];
    var entry = jpb[idx];
    if (!entry) {
      return null;
    }
    return { name: resolved, idx: idx, entry: entry };
  }

  /**
   * Excel 核心修复：KPb 已注册的自定义字体必须直接查 catalog，
   * 不能走 FontPicker（dW.isd/oK），否则会被模糊匹配成 Arial。
   */
  function installCellCatalogFontResolution() {
    var dW = window.AscFonts && window.AscFonts.dW;
    if (!dW || !window.AscFonts.KPb || !window.AscFonts.JPb || dW.__CATALOG_FONT_PATCHED__) {
      return false;
    }

    if (dW.isd) {
      var origIsd = dW.isd.bind(dW);
      dW.isd = function (name) {
        var plain = plainFontName(name);
        var hit = resolveCellCatalogEntry(plain);
        if (hit) {
          if (dW.oge[plain] && dW.oge[plain].tha === hit.name) {
            return dW.oge[plain];
          }
          var pick = { jM: plain, tha: hit.name };
          dW.oge[plain] = pick;
          if (plain !== hit.name) {
            dW.oge[hit.name] = pick;
          }
          return pick;
        }
        return origIsd(resolveCellCatalogFamily(plain));
      };
    }

    if (dW.oK) {
      var origOk = dW.oK.bind(dW);
      dW.oK = function (name, style, out) {
        var hit = resolveCellCatalogEntry(name);
        if (hit) {
          if (out) {
            out.Fa = hit.name;
            if (dW.Eyf) {
              out.bU = dW.Eyf(name, out);
            }
          }
          return hit.entry;
        }
        return origOk(resolveCellCatalogFamily(plainFontName(name)), style, out);
      };
    }

    if (dW.ksd) {
      var origKsd = dW.ksd.bind(dW);
      dW.ksd = function (name) {
        var hit = resolveCellCatalogEntry(name);
        if (hit) {
          return hit.name;
        }
        return origKsd(resolveCellCatalogFamily(plainFontName(name)));
      };
    }

    clearCellFontPickerCache(dW);
    dW.__CATALOG_FONT_PATCHED__ = true;
    return true;
  }

  function injectCustomFontBinariesWord(ids) {
    var asc = window.AscFonts;
    if (!asc || !asc.dpc || !asc.epc || !asc.Tcc) {
      return false;
    }

    var injected = false;
    for (var i = 0; i < ids.length; i++) {
      var fileId = ids[i];
      var fileIndex = findDpcIndex(asc.dpc, fileId);
      if (fileIndex < 0) {
        continue;
      }

      var entry = asc.dpc[fileIndex];
      if (!entry || (entry.hKa && entry.hKa())) {
        continue;
      }

      var data = loadCatalogFontSync(fileId);
      if (!data) {
        continue;
      }

      var streamIndex = asc.epc.length;
      asc.epc.push(new asc.Tcc(data, data.length));
      if (entry.wXb) {
        entry.wXb(streamIndex);
      }
      entry.EB = 0;
      if (asc.$0d) {
        asc.$0d(streamIndex);
      }
      injected = true;
    }

    return injected;
  }

  function pushCellFontStream(asc, entry, data) {
    var StreamCtor = asc && (asc.vlc || asc.JAf);
    if (!asc || !asc.zyc || !StreamCtor || !entry || !data) {
      return false;
    }

    var streamIndex = asc.zyc.length;
    asc.zyc.push(new StreamCtor(data, data.length));
    if (entry.b4b) {
      entry.b4b(streamIndex);
    }
    entry.AE = 0;
    if (asc.Zde) {
      asc.Zde(streamIndex);
    }
    return true;
  }

  function loadCellFontBinaryAsync(entry, callback) {
    var V_ = window.AscCommon && window.AscCommon.V_;
    if (!entry || !entry.c6a || !V_) {
      if (callback) {
        callback(false);
      }
      return;
    }
    if (entry.$Pa && entry.$Pa()) {
      if (callback) {
        callback(true);
      }
      return;
    }

    var base = V_.uzg || getFontsBaseUrl();
    entry.c6a(base, function () {
      if (callback) {
        callback(!!(entry.$Pa && entry.$Pa()));
      }
    });
  }

  // 把 registry 字体二进制同步写入 yyc/zyc；1001 等不在静态 __fonts_files 里，靠运行时 catalog。
  function injectCustomFontBinariesCell(ids) {
    var asc = window.AscFonts;
    if (!asc || !asc.yyc || !asc.zyc) {
      return false;
    }

    var injected = false;
    for (var i = 0; i < ids.length; i++) {
      var fileId = ids[i];
      var fileIndex = findYycIndex(asc.yyc, fileId);
      if (fileIndex < 0) {
        continue;
      }

      var entry = asc.yyc[fileIndex];
      if (!entry || (entry.$Pa && entry.$Pa())) {
        injected = true;
        continue;
      }

      var data = loadCatalogFontSync(fileId);
      if (data && pushCellFontStream(asc, entry, data)) {
        clearCellFontPickerCache(window.AscFonts && window.AscFonts.dW);
        injected = true;
        continue;
      }

      loadCellFontBinaryAsync(entry, function (ok) {
        if (!ok) {
          return;
        }
        reloadCellDocumentFontsFromWorkbook();
        scheduleSpreadsheetLayoutRefresh();
      });
      injected = true;
    }

    return injected;
  }

  // Word 文档字体队列：mJ.GPb + Lif 触发 SDK 异步加载（Word 排版刷新主要靠此，非 asc_calculate）。
  function wordCatalogReady() {
    var asc = window.AscFonts;
    return !!(
      asc &&
      asc.QQ &&
      asc.L1b &&
      asc.dpc &&
      asc.epc &&
      asc.Tcc
    );
  }

  function syncWordEngineFontRefs() {
    var mJ = window.AscCommon && window.AscCommon.mJ;
    if (!mJ) {
      return false;
    }
    var base = getFontsBaseUrl();
    mJ.zbg = base;
    window.__ONLYOFFICE_FONTS_BASE__ = base;
    return true;
  }

  function collectWordBinaryFileIds(registry) {
    var ids = [];
    var seen = {};
    for (var id in registry) {
      if (!Object.prototype.hasOwnProperty.call(registry, id)) {
        continue;
      }
      if (!seen[id]) {
        seen[id] = true;
        ids.push(id);
      }
    }
    return ids;
  }

  function areWordRegistryBinariesReady(ids) {
    var asc = window.AscFonts;
    if (!asc || !asc.dpc) {
      return false;
    }
    for (var i = 0; i < ids.length; i++) {
      var fileIndex = findDpcIndex(asc.dpc, ids[i]);
      if (fileIndex < 0) {
        return false;
      }
      var entry = asc.dpc[fileIndex];
      if (!entry || !(entry.hKa && entry.hKa())) {
        return false;
      }
    }
    return ids.length > 0;
  }

  function prepareWordFontCatalog() {
    if (!wordCatalogReady()) {
      return false;
    }

    var registry = loadCustomFontRegistry();
    registerAliasFamiliesWord(registry);
    if (!PATCHED_WORD) {
      fixFontPickerThumbnails(registry);
      hookEditorForFontList();
      hookFontListInit();
      installWordCatalogFontResolution();
      scheduleToolbarFontListReload();
      PATCHED_WORD = true;
    }
    return true;
  }

  function ensureWordFontBinaries() {
    if (WORD_BINARIES_INJECTED) {
      return true;
    }
    if (!prepareWordFontCatalog()) {
      return false;
    }

    syncWordEngineFontRefs();
    var registry = loadCustomFontRegistry();
    var ids = collectWordBinaryFileIds(registry);
    injectCustomFontBinariesWord(ids);
    if (areWordRegistryBinariesReady(ids)) {
      WORD_BINARIES_INJECTED = true;
      return true;
    }
    return false;
  }

  function reloadWordDocumentFontsFromEditor() {
    if (!ensureWordFontBinaries()) {
      return false;
    }
    queueDocumentFontsWord(expandRegistry(loadCustomFontRegistry()).names);
    scheduleLayoutRefresh();
    return true;
  }

  function hookWordDocumentFontLoading() {
    var mJ = window.AscCommon && window.AscCommon.mJ;
    if (!mJ || mJ.__CUSTOM_FONTS_WORD_PATCHED__) {
      return false;
    }
    if (!mJ.GPb && !mJ.Lif) {
      return false;
    }

    function beforeDocumentFonts() {
      ensureWordFontBinaries();
    }

    if (mJ.GPb) {
      var origGPb = mJ.GPb;
      mJ.GPb = function () {
        beforeDocumentFonts();
        return origGPb.apply(this, arguments);
      };
    }

    if (mJ.Lif) {
      var origLif = mJ.Lif;
      mJ.Lif = function () {
        beforeDocumentFonts();
        return origLif.apply(this, arguments);
      };
    }

    if (mJ.GE) {
      var origGE = mJ.GE;
      mJ.GE = function () {
        beforeDocumentFonts();
        return origGE.apply(this, arguments);
      };
    }

    mJ.__CUSTOM_FONTS_WORD_PATCHED__ = true;
    return true;
  }

  function hookWordEngineInit() {
    var common = (window.AscCommon = window.AscCommon || {});
    if (common.__CUSTOM_FONTS_MJ_WATCH__) {
      return false;
    }

    common.__CUSTOM_FONTS_MJ_WATCH__ = true;
    var engine = common.mJ;
    Object.defineProperty(common, "mJ", {
      configurable: true,
      get: function () {
        return engine;
      },
      set: function (value) {
        engine = value;
        if (value) {
          syncWordEngineFontRefs();
          hookWordDocumentFontLoading();
        }
      },
    });

    if (engine) {
      syncWordEngineFontRefs();
      hookWordDocumentFontLoading();
    }
    return true;
  }

  function queueDocumentFontsWord(names) {
    var mJ = window.AscCommon && window.AscCommon.mJ;
    if (!mJ || !mJ.GPb || !names.length) {
      return;
    }
    for (var i = 0; i < names.length; i++) {
      mJ.GPb(names[i]);
    }
    if (mJ.Lif) {
      mJ.Lif();
    }
  }

  // Excel 文档字体队列：V_.tXb + ryf（pipeline 内使用，与 Word 的 mJ 不同）。
  function queueDocumentFontsCell(names) {
    var V_ = window.AscCommon && window.AscCommon.V_;
    var fi = window.AscFonts && window.AscFonts.fi;
    if (!V_ || !V_.tXb || !names.length) {
      return;
    }

    if (fi && fi.PO) {
      for (var i = 0; i < names.length; i++) {
        fi.PO(names[i]);
      }
    }

    for (var j = 0; j < names.length; j++) {
      V_.tXb(names[j], 15);
    }
    if (V_.ryf) {
      V_.ryf();
    }
  }

  // --- Excel / Word / Slide 排版刷新（编辑器类型必须分开判断）---

  // 识别 Spreadsheet：看 workbook API；wa 可能尚未创建，不能等 wa 才认定是 Excel。
  function isSpreadsheetEditor(editor) {
    if (!editor || !editor.wd) {
      return false;
    }
    return (
      typeof editor.wd.$b === "function" &&
      editor.wd.yg !== undefined
    );
  }

  // Spreadsheet 视图就绪：OEd 之后才有 wa，此时 asc_calculate 才安全。
  function isSpreadsheetViewReady(editor) {
    return !!(
      isSpreadsheetEditor(editor) &&
      editor.wa &&
      typeof editor.wa.wi === "function"
    );
  }

  // 识别 Presentation：Slide 把 AscCommon.lU 挂在 editor.yga；Word 不用 lU。
  function isPresentationEditor(editor) {
    if (!editor || isSpreadsheetEditor(editor)) {
      return false;
    }
    var lU = window.AscCommon && window.AscCommon.lU;
    return !!(lU && editor.yga === lU && editor.ra);
  }

  function canRefreshSpreadsheetLayout() {
    return isSpreadsheetViewReady(window.Asc && window.Asc.editor);
  }

  function canRefreshWordLayout() {
    var editor = window.Asc && window.Asc.editor;
    if (!editor || isSpreadsheetEditor(editor) || isPresentationEditor(editor)) {
      return false;
    }
    return !!(editor.wd && editor.asc_Recalculate);
  }

  function canRefreshSlideLayout() {
    // 禁止 scheduleLayoutRefresh 对 Slide 主动 Me/Rp/WN；重排仅由 JOf hook 在 reload 后触发一次。
    return false;
  }

  function canRefreshDocumentLayout() {
    return (
      canRefreshSpreadsheetLayout() ||
      canRefreshWordLayout() ||
      canRefreshSlideLayout()
    );
  }

  function refreshSpreadsheetLayout() {
    var editor = window.Asc && window.Asc.editor;
    if (!isSpreadsheetViewReady(editor)) {
      return;
    }
    // Excel：asc_calculate → wa.wi → 重算并 ed() 重绘当前 sheet
    if (editor.asc_calculate) {
      editor.asc_calculate();
    }
  }

  function refreshWordLayout() {
    var editor = window.Asc && window.Asc.editor;
    if (!editor || isSpreadsheetEditor(editor) || isPresentationEditor(editor)) {
      return;
    }
    if (editor.asc_Recalculate) {
      editor.asc_Recalculate();
    }
  }

  function refreshSlideLayout() {
    // 仅 Me+Rp 更新主画布逻辑文档；不调用 ra.yd.WN()（全量缩略图刷新会导致侧栏消失）。
    var editor = window.Asc && window.Asc.editor;
    if (!isPresentationEditor(editor)) {
      return;
    }
    var ra = editor.ra;
    if (ra && ra.Ea && ra.Ea.Me) {
      ra.Ea.Me();
    }
    if (ra && ra.Rp) {
      ra.Rp();
    }
  }

  function refreshDocumentLayout() {
    var editor = window.Asc && window.Asc.editor;
    if (isSpreadsheetEditor(editor)) {
      if (isSpreadsheetViewReady(editor)) {
        refreshSpreadsheetLayout();
      }
      return;
    }
    if (isPresentationEditor(editor)) {
      refreshSlideLayout();
      return;
    }
    if (canRefreshWordLayout()) {
      refreshWordLayout();
    }
  }

  function scheduleLayoutRefresh() {
    if (!canRefreshDocumentLayout()) {
      return;
    }
    window.setTimeout(refreshDocumentLayout, 0);
    window.setTimeout(refreshDocumentLayout, 300);
  }

  // Excel 专用：等 wa 就绪再 refresh，避免 wa 为 null 时调用 asc_calculate。
  function scheduleSpreadsheetLayoutRefresh() {
    waitForSpreadsheetReady(scheduleLayoutRefresh);
  }

  // 仅 Excel：OEd 创建 wa 是异步的；DGc 回调里需轮询到 wa 再 reload + refresh。
  function waitForSpreadsheetReady(callback) {
    if (!callback) {
      return;
    }
    if (isSpreadsheetViewReady(window.Asc && window.Asc.editor)) {
      callback();
      return;
    }
    var tries = 0;
    var timer = window.setInterval(function () {
      if (isSpreadsheetViewReady(window.Asc && window.Asc.editor) || ++tries > 200) {
        window.clearInterval(timer);
        if (isSpreadsheetViewReady(window.Asc && window.Asc.editor)) {
          callback();
        }
      }
    }, 50);
  }

  // --- Excel pipeline ---

  // 汇总 Excel 侧补丁：catalog、别名、FontPicker 绕过、二进制注入。
  function prepareCellFontCatalog() {
    ensureCellFontCatalogFromSnapshot();
    if (
      !window.AscFonts ||
      !window.AscFonts.yyc ||
      !window.AscFonts.KPb ||
      !window.AscFonts.dW ||
      !cellFontCatalogHasFamilies(window.AscFonts.KPb)
    ) {
      return false;
    }
    syncCellEngineFontRefs();

    var registry = loadCustomFontRegistry();
    fixFontPickerThumbnails(registry);
    hookFontListInit();
    var expanded = expandRegistry(registry);
    registerAliasFamiliesCell(registry);
    registerDocumentFontAliasesCell(registry);
    hookCellDocumentFontRegistration();
    hookCellExcelFontBe();
    installCellCatalogFontResolution();
    clearCellFontPickerCache(window.AscFonts.dW);
    injectCustomFontBinariesCell(expanded.ids);
    scheduleToolbarFontListReload();
    return expanded;
  }

  /**
   * Excel 打开后首次正确渲染的关键：用 workbook 字体列表重新走 jyb + koc，
   * 在二进制与 FontPicker 补丁生效后再触发 SDK 异步加载与重绘。
   */
  function reloadCellDocumentFontsFromWorkbook() {
    var editor = window.Asc && window.Asc.editor;
    var wd = editor && editor.wd;
    var V_ = window.AscCommon && window.AscCommon.V_;
    if (!wd || !wd.X5b || !V_) {
      return false;
    }

    prepareCellFontCatalog();
    var expanded = expandRegistry(loadCustomFontRegistry());
    if (expanded) {
      injectCustomFontBinariesCell(expanded.ids);
    }

    clearCellFontPickerCache(window.AscFonts && window.AscFonts.dW);

    var list = buildWorkbookFontList(wd);
    if (window.AscFonts && window.AscFonts.fi && window.AscFonts.fi.jyb) {
      window.AscFonts.fi.jyb(list);
    }

    var Excel = window.AscCommonExcel;
    if (Excel && Excel.zsg && wd.RW) {
      Excel.zsg(wd.RW);
    }
    if (V_.koc) {
      V_.koc(list);
    } else if (V_.ryf) {
      V_.ryf();
    }
    return true;
  }

  // prepare → queue → reload → scheduleLayoutRefresh 的 Excel 字体安装入口。
  function runCellFontPipeline(queueAll) {
    var expanded = prepareCellFontCatalog();
    if (!expanded) {
      return false;
    }
    if (queueAll) {
      queueDocumentFontsCell(expanded.names);
      reloadCellDocumentFontsFromWorkbook();
      scheduleSpreadsheetLayoutRefresh();
    }
    return true;
  }

  // --- Excel SDK hook（acj / V_ / koc / DGc）---

  // 拦截 AscFonts.acj 赋值：Cell SDK 注册 catalog 后立即跑 pipeline。
  function hookCellFontCatalogInit() {
    var asc = (window.AscFonts = window.AscFonts || {});
    if (asc.__CUSTOM_FONTS_ACJ_WATCH__) {
      return false;
    }

    asc.__CUSTOM_FONTS_ACJ_WATCH__ = true;
    var wrapped = asc.acj;
    Object.defineProperty(asc, "acj", {
      configurable: true,
      get: function () {
        return wrapped;
      },
      set: function (fn) {
        if (!fn || fn.__CUSTOM_FONTS_ACJ_WRAPPED__) {
          wrapped = fn;
          return;
        }
        var inner = fn;
        asc.__CUSTOM_FONT_ACJ_INNER__ = inner;
        wrapped = function () {
          if (
            !window.AscFonts ||
            !window.AscFonts.yyc ||
            !cellFontCatalogHasFamilies(window.AscFonts.KPb)
          ) {
            buildCellFontCatalogFromSnapshot();
          } else {
            syncCellEngineFontRefs();
          }
          var result = inner.apply(this, arguments);
          syncCellEngineFontRefs();
          runCellFontPipeline(true);
          hookCellDocumentFontLoading();
          return result;
        };
        wrapped.__CUSTOM_FONTS_ACJ_WRAPPED__ = true;
        // Cell SDK 赋值后会同步调用 h() 而非 asc.acj()，因此在下一轮再补跑 pipeline。
        window.setTimeout(function () {
          ensureCellFontCatalogFromSnapshot();
          syncCellEngineFontRefs();
          runCellFontPipeline(true);
          hookCellDocumentFontLoading();
        }, 0);
      },
    });

    if (typeof wrapped === "function") {
      asc.acj = wrapped;
    }

    return true;
  }

  function hookCellEngineInit() {
    var common = (window.AscCommon = window.AscCommon || {});
    if (common.__CUSTOM_FONTS_V_WATCH__) {
      return false;
    }

    common.__CUSTOM_FONTS_V_WATCH__ = true;
    var engine = common.V_;
    Object.defineProperty(common, "V_", {
      configurable: true,
      get: function () {
        return engine;
      },
      set: function (value) {
        engine = value;
        if (value) {
          hookCellDocumentFontLoading();
        }
      },
    });

    if (engine) {
      hookCellDocumentFontLoading();
    }

    return true;
  }

  // 文档打开时 V_.E0b / V_.koc 加载字体前先 prepareCellFontCatalog。
  function hookCellDocumentFontLoading() {
    var V_ = window.AscCommon && window.AscCommon.V_;
    if (!V_ || V_.__CUSTOM_FONTS_DOC_FONTS_PATCHED__) {
      return false;
    }
    if (!V_.E0b && !V_.koc) {
      return false;
    }

    function beforeDocumentFonts() {
      prepareCellFontCatalog();
    }

    if (V_.E0b) {
      var origE0b = V_.E0b;
      V_.E0b = function (fonts, reason, callback) {
        beforeDocumentFonts();
        return origE0b.call(this, fonts, reason, callback);
      };
    }

    if (V_.koc) {
      var origKoc = V_.koc;
      V_.koc = function (fonts) {
        beforeDocumentFonts();
        return origKoc.call(this, fonts);
      };
    }

    V_.__CUSTOM_FONTS_DOC_FONTS_PATCHED__ = true;
    return true;
  }

  // V_.DGc：Spreadsheet 编辑器绑定后，等 wa 就绪再 reload + refresh（无需用户编辑单元格）。
  function hookCellEditorInit() {
    var V_ = window.AscCommon && window.AscCommon.V_;
    if (!V_ || V_.__CUSTOM_FONTS_DGC_PATCHED__ || !V_.DGc) {
      return false;
    }

    var origDgc = V_.DGc;
    V_.DGc = function (editor) {
      origDgc.call(this, editor);
      runCellFontPipeline(true);
      waitForSpreadsheetReady(function () {
        reloadCellDocumentFontsFromWorkbook();
        scheduleSpreadsheetLayoutRefresh();
      });
    };
    V_.__CUSTOM_FONTS_DGC_PATCHED__ = true;
    return true;
  }

  function hookInitNativeEditors() {
    if (window.__CUSTOM_FONTS_INIT_NATIVE_PATCHED__) {
      return false;
    }
    var orig = window.InitNativeEditors;
    if (typeof orig !== "function") {
      return false;
    }

    window.__CUSTOM_FONTS_INIT_NATIVE_PATCHED__ = true;
    window.InitNativeEditors = function () {
      orig.apply(this, arguments);
      ensureCellFontCatalogFromSnapshot();
      syncCellEngineFontRefs();
      runCellFontPipeline(true);
      reloadCellDocumentFontsFromWorkbook();
    };
    return true;
  }

  // ---------------------------------------------------------------------------
  // 启动：注册 hook，分别安装 Word / Excel 补丁
  // ---------------------------------------------------------------------------

  function tryInstallWord() {
    hookWordEngineInit();
    if (!wordCatalogReady()) {
      return PATCHED_WORD && WORD_BINARIES_INJECTED;
    }
    if (ensureWordFontBinaries()) {
      queueDocumentFontsWord(expandRegistry(loadCustomFontRegistry()).names);
      scheduleLayoutRefresh();
    }
    return PATCHED_WORD && WORD_BINARIES_INJECTED;
  }

  function tryInstallCell() {
    if (PATCHED_CELL) {
      return PATCHED_CELL;
    }
    // Cell 走 runCellFontPipeline，内部会 reload workbook 字体并 scheduleLayoutRefresh。
    if (!runCellFontPipeline(true)) {
      return false;
    }
    PATCHED_CELL = true;
    return true;
  }

  function tryInstallSlide() {
    // 依赖 uji/h() 已创建 jec、hyb、jR；PPT 页不加载 Word 的 QQ/L1b。
    if (
      PATCHED_SLIDE ||
      !window.AscFonts ||
      !window.AscFonts.jec ||
      !window.AscFonts.hyb ||
      !window.AscFonts.jR
    ) {
      return PATCHED_SLIDE;
    }

    return runSlideFontPipelineAfterCatalogInit();
  }

  function isFontPatchingComplete() {
    var asc = window.AscFonts;
    var wordNeeded = !!(asc && asc.QQ && asc.L1b);
    var wordDone =
      !wordNeeded || (PATCHED_WORD && WORD_BINARIES_INJECTED);
    var cellDone = PATCHED_CELL || !(asc && asc.yyc && asc.KPb);
    var slideDone =
      PATCHED_SLIDE || !(asc && asc.jec && asc.hyb && asc.jR);
    return wordDone && cellDone && slideDone;
  }

  function tryInstallAll() {
    tryInstallWord();
    tryInstallCell();
    tryInstallSlide();
    return isFontPatchingComplete();
  }

  // web-apps 通过 RequireJS 晚于 SDK 初始化。主安装轮询可能已经完成，
  // 此时 ComboBoxFonts 才定义；单独等待它，保证搜索缓存补丁一定能挂上。
  function waitForWebAppsComboFontHook() {
    var tries = 0;
    var timer = window.setInterval(function () {
      var patched = hookComboBoxFontsWebApps();
      var proto =
        window.Common &&
        window.Common.UI &&
        window.Common.UI.ComboBoxFonts &&
        window.Common.UI.ComboBoxFonts.prototype;
      if (
        patched ||
        (proto && proto.__CUSTOM_FONTS_COMBO_PATCHED__) ||
        ++tries > 1200
      ) {
        window.clearInterval(timer);
      }
    }, 50);
  }

  // Cell / Slide 会在 AllFonts 之后替换 AscFonts；主安装轮询可能已结束，
  // 因此单独等待最终运行时的 pickFont 再挂接自定义文件 id 映射。
  function waitForRuntimeCustomFontPicker() {
    var tries = 0;
    var timer = window.setInterval(function () {
      if (hookRuntimeCustomFontPicker() || ++tries > 1200) {
        window.clearInterval(timer);
      }
    }, 50);
  }

  // hookSlideFontManagerInit 必须最先注册（AllFonts 早于 slide sdk 加载）。
  function pollUntilReady() {
    hookSlideFontManagerInit();
    hookSlideFontCatalogInit();
    hookSlideDocumentFontLoading();
    hookSlideFontPickerGl();
    if (tryInstallSlide()) {
      return;
    }
    var timer = window.setInterval(function () {
      hookSlideFontManagerInit();
      hookSlideFontCatalogInit();
      hookSlideDocumentFontLoading();
      hookSlideFontPickerGl();
      if (tryInstallSlide()) {
        window.clearInterval(timer);
      }
    }, 50);
  }

  window.__reloadEditorFontListForToolbar = function () {
    return reloadEditorFontListForToolbar();
  };

  window.__forceCustomFontBinaries = function () {
    var ok = runCellFontPipeline(true);
    reloadCellDocumentFontsFromWorkbook();
    scheduleSpreadsheetLayoutRefresh();
    return ok;
  };

  window.__remapAndReloadExcelFonts = function () {
    reloadCellDocumentFontsFromWorkbook();
    scheduleSpreadsheetLayoutRefresh();
    return true;
  };

  // PPT iframe 控制台：__debugSlideFonts() — 检查 patch / 二进制 / reload 状态。
  window.__debugSlideFonts = function () {
    var asc = window.AscFonts;
    var editor = window.Asc && window.Asc.editor;
    var lU = window.AscCommon && window.AscCommon.lU;
    var primary = listFontNames(loadCustomFontRegistry()["1001"] || [])[0];
    var hit = primary ? resolveSlideCatalogEntry(primary) : null;
    var jecEntry =
      asc && asc.jec && primary
        ? asc.jec[findJecIndex(asc.jec, "1001")]
        : null;
    console.log("__debugSlideFonts", {
      primary: primary,
      hybPrimary: primary && asc && asc.hyb ? asc.hyb[primary] : undefined,
      patchedSlide: PATCHED_SLIDE,
      catalogReady: SLIDE_CATALOG_READY,
      binariesInjected: SLIDE_BINARIES_INJECTED,
      unpatchTjb: SLIDE_UNPATCHED_TJB,
      fontsReloaded: SLIDE_FONTS_RELOADED,
      layoutRefreshed: SLIDE_LAYOUT_REFRESHED,
      jRPatched: !!(asc && asc.jR && asc.jR.__CATALOG_FONT_PATCHED__),
      hybProxy: !!(asc && asc.hyb && asc.hyb.__CUSTOM_FONTS_HYB_PROXY__),
      fontsBase: syncSlideFontBaseUrl(),
      isPresentation: isPresentationEditor(editor),
      editorYgaIsLU: !!(editor && lU && editor.yga === lU),
      jecLoaded: !!(jecEntry && jecEntry.ODa && jecEntry.ODa()),
      hit: hit,
    });
    return {
      primary: primary,
      patchedSlide: PATCHED_SLIDE,
      hit: hit,
    };
  };

  window.__debugWordFonts = function () {
    var asc = window.AscFonts;
    var mJ = window.AscCommon && window.AscCommon.mJ;
    var l1b = asc && asc.L1b;
    var registry = loadCustomFontRegistry();
    var ids = collectWordBinaryFileIds(registry);
    var docNames = [
      "楷体_GB2312",
      "仿宋_GB2312",
      "黑体",
      "宋体",
      "方正小标宋简体",
      "方正仿宋_GBK",
    ];
    var aliases = {};
    var binaries = {};
    var i;
    for (i = 0; i < docNames.length; i++) {
      aliases[docNames[i]] = l1b ? l1b[docNames[i]] : undefined;
    }
    for (i = 0; i < ids.length; i++) {
      var fileId = ids[i];
      var idx = asc && asc.dpc ? findDpcIndex(asc.dpc, fileId) : -1;
      var entry = idx >= 0 && asc.dpc ? asc.dpc[idx] : null;
      binaries[fileId] = {
        idx: idx,
        EB: entry && entry.EB,
        ok: !!(entry && entry.hKa && entry.hKa()),
      };
    }
    console.log("__debugWordFonts", {
      patchedWord: PATCHED_WORD,
      binariesInjected: WORD_BINARIES_INJECTED,
      fontsBase: getFontsBaseUrl(),
      mJzbg: mJ && mJ.zbg,
      aliases: aliases,
      binaries: binaries,
    });
    return {
      patchedWord: PATCHED_WORD,
      binariesInjected: WORD_BINARIES_INJECTED,
      aliases: aliases,
      binaries: binaries,
    };
  };

  window.__debugCellFonts = function () {
    var asc = window.AscFonts;
    var V_ = window.AscCommon && window.AscCommon.V_;
    var editor = window.Asc && window.Asc.editor;
    var wd = editor && editor.wd;
    var primary = listFontNames(loadCustomFontRegistry()["1001"] || [])[0];
    var yycEntry =
      asc && asc.yyc && primary
        ? asc.yyc[asc.KPb && asc.KPb[primary] !== undefined ? asc.JPb[asc.KPb[primary]].qea : -1]
        : null;
    var docFonts = wd && wd.X5b ? wd.X5b() : {};
    console.log("__debugCellFonts", {
      uzg: V_ && V_.uzg,
      primary: primary,
      kpbPrimary: primary && asc && asc.KPb ? asc.KPb[primary] : undefined,
      yycVa: yycEntry && yycEntry.Va,
      yycAE: yycEntry && yycEntry.AE,
      yycStream: yycEntry && yycEntry.Ahc,
      docFonts: docFonts,
    });
    return {
      uzg: V_ && V_.uzg,
      primary: primary,
      docFonts: docFonts,
    };
  };

  pollUntilReady();
})(window);
