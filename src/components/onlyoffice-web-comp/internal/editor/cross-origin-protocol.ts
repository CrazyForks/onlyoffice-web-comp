export const CROSS_ORIGIN_BRIDGE_MESSAGE = {
  EDITOR_COMMAND: "editor:command",
  EDITOR_RESPONSE: "editor:response",
  EDITOR_EVENT: "editor:event",
  EDITOR_SET_READONLY: "editor:set-readonly",
} as const;

export const CROSS_ORIGIN_EDITOR_COMMAND = {
  EDITOR_SUBSCRIBE: "editor:subscribe",
  COMMENT_ADD: "comment:add",
  COMMENT_UPDATE: "comment:update",
  COMMENT_REMOVE: "comment:remove",
  COMMENT_GO_TO: "comment:go-to",
  COMMENT_LIST: "comment:list",
  COMMENT_SUBSCRIBE: "comment:subscribe",
  REVISION_ADD_DEMO: "revision:add-demo",
  REVISION_LIST: "revision:list",
  REVISION_SET_TRACK: "revision:set-track",
  REVISION_IS_TRACK: "revision:is-track",
  REVISION_HAVE_CHANGES: "revision:have-changes",
  REVISION_PREPARE_REVIEW: "revision:prepare-review",
  REVISION_NEXT: "revision:next",
  REVISION_PREV: "revision:prev",
  REVISION_GO_TO: "revision:go-to",
  REVISION_ACCEPT: "revision:accept",
  REVISION_REJECT: "revision:reject",
  REVISION_ACCEPT_ALL: "revision:accept-all",
  REVISION_REJECT_ALL: "revision:reject-all",
  REVISION_ACCEPT_SELECTION: "revision:accept-selection",
  REVISION_REJECT_SELECTION: "revision:reject-selection",
  REVISION_SUBSCRIBE: "revision:subscribe",
} as const;

export const CROSS_ORIGIN_EDITOR_EVENT = {
  ADD_COMMENT: "asc_onAddComment",
  CHANGE_COMMENT: "asc_onChangeCommentData",
  REMOVE_COMMENT: "asc_onRemoveComment",
  SHOW_REVISIONS_CHANGE: "asc_onShowRevisionsChange",
  TRACK_REVISIONS_CHANGE: "asc_onOnTrackRevisionsChange",
  DOCUMENT_MODIFIED_CHANGED: "asc_onDocumentModifiedChanged",
} as const;
