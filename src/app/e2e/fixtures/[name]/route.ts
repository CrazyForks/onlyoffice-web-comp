import fs from "node:fs/promises";
import path from "node:path";

const fixtureDir = path.join(
  process.cwd(),
  "tests",
  "e2e",
  "files",
);

const mimeTypes: Record<string, string> = {
  ".docx":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx":
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx":
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".json": "application/json; charset=utf-8",
};

export async function GET(
  _request: Request,
  context: { params: Promise<{ name: string }> },
) {
  const { name } = await context.params;

  if (!/^[a-z0-9][a-z0-9._-]*$/i.test(name)) {
    return new Response("Bad Request", { status: 400 });
  }

  const filePath = path.join(fixtureDir, name);
  if (!filePath.startsWith(fixtureDir + path.sep)) {
    return new Response("Forbidden", { status: 403 });
  }

  try {
    const data = await fs.readFile(filePath);
    return new Response(new Uint8Array(data), {
      headers: {
        "Cache-Control": "no-cache",
        "Content-Type":
          mimeTypes[path.extname(name).toLowerCase()] ||
          "application/octet-stream",
      },
    });
  } catch {
    return new Response("Not Found", { status: 404 });
  }
}
