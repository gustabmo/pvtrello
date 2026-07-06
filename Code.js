/**
 * PV (Procès-Verbal) Generator
 * ----------------------------
 * Reads config from the "PVConfig" sheet, pulls all lists and cards (including
 * archived ones) from a Trello board, and writes every activity that happened
 * within a given date/time window into the target tab, in chronological order.
 *
 * Handles large boards safely: Trello calls are batched (up to 10 cards' worth
 * of activity per HTTP request), and if a run still can't finish inside Apps
 * Script's 6-minute execution limit, the script saves its progress and
 * reschedules itself to continue a few seconds later - fully automatically.
 *
 * SETUP
 * 1. Extensions > Apps Script, paste this file in.
 * 2. Fill in PRIVATE_Trello_APIkey and PRIVATE_Trello_APItoken below.
 *    (Get a key at https://trello.com/app-key, then generate a token from that page.)
 * 3. Create a sheet tab called "PVConfig" with a header row containing:
 *      IdBoard | From | To | PVTab
 *    and, in the row right below, the values (From/To must be real date-time cells).
 * 4. Run generatePV (or use the "PV Tools" menu that appears when you open the file).
 *    If it needs more than one run to finish, you'll see it keep growing every
 *    ~10 seconds until a toast confirms it's done - no need to re-run it yourself.
 */


const CONFIG_SHEET_NAME = "PVConfig";
const TRELLO_API_BASE = "https://api.trello.com/1";
const BATCH_CHUNK_SIZE = 10;          // Trello's /batch endpoint caps at 10 sub-requests
const TIME_BUDGET_MS = 4.5 * 60 * 1000; // stop and reschedule before hitting the 6-min hard cap
const CONTINUATION_DELAY_SEC = 10;
const PROP_STATE = "PV_STATE";
const PROP_TRIGGER_ID = "PV_TRIGGER_ID";

// ---------------------------------------------------------------------------
// Menu (optional convenience)
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("PV Tools")
    .addItem("Generate PV", "generatePV")
    .addItem("Cancel PV generation in progress", "cancelPV")
    .addToUi();
}

// ---------------------------------------------------------------------------
// Main entry points
// ---------------------------------------------------------------------------
function generatePV() {
  cancelPV(); // clear any stuck previous run/trigger before starting fresh

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_(ss);
  getOrCreatePvSheet_(ss, config.pvTabName); // creates/clears + writes header row

  const lists = getAllLists_(config.idBoard);

  const state = {
    idBoard: config.idBoard,
    from: config.from.toISOString(),
    to: config.to.toISOString(),
    pvTabName: config.pvTabName,
    lists: lists.map(function (l) { return { id: l.id, name: l.name }; }),
    listIndex: 0,
    cardOffset: 0
  };

  runChunk_(state);
}

// Called by the auto-reschedule trigger; also safe to run manually to resume.
function continuePV_() {
  const state = loadState_();
  if (!state) return; // nothing in progress
  runChunk_(state);
}

// Stops an in-progress run and removes its trigger/state (partial data stays in the sheet).
function cancelPV() {
  const triggerId = PropertiesService.getScriptProperties().getProperty(PROP_TRIGGER_ID);
  if (triggerId) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
    });
  }
  PropertiesService.getScriptProperties().deleteProperty(PROP_TRIGGER_ID);
  PropertiesService.getScriptProperties().deleteProperty(PROP_STATE);
}

// ---------------------------------------------------------------------------
// Core chunked processing loop
// ---------------------------------------------------------------------------
function runChunk_(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(state.pvTabName);
  const from = new Date(state.from);
  const to = new Date(state.to);
  const startTime = Date.now();

  while (state.listIndex < state.lists.length) {
    const list = state.lists[state.listIndex];
    const cards = getAllCards_(list.id);

    for (let i = state.cardOffset; i < cards.length; i += BATCH_CHUNK_SIZE) {
      if (Date.now() - startTime > TIME_BUDGET_MS) {
        state.cardOffset = i;
        saveState_(state);
        scheduleContinuation_();
        ss.toast("PV generation paused at list \"" + list.name + "\" - it will resume automatically in a few seconds.");
        return;
      }

      const chunk = cards.slice(i, i + BATCH_CHUNK_SIZE);
      const actionsByCard = getActionsBatch_(chunk.map(function (c) { return c.id; }), from, to);

      const rows = [];
      chunk.forEach(function (card, idx) {
        const activities = (actionsByCard[idx] || [])
          .map(describeAction_)
          .filter(function (a) { return a !== null; })
          .sort(function (a, b) { return a.date - b.date; }); // chronological within the card

        first = true;
        activities.forEach(function (activity) {
          rows.push([
            first ? list.name : "", 
            first ? card.name : "", 
            first ? (card.shortUrl || card.url) : "", 
            activity.date, 
            activity.text
          ]);
          first = false;
        });
      });

      if (rows.length > 0) appendRows_(sheet, rows);
    }

    state.listIndex++;
    state.cardOffset = 0;
    saveState_(state); // checkpoint after each completed list
  }

  finalizeSheet_(sheet);
  cancelPV(); // clear trigger/state now that we're done
  ss.toast("PV generation complete: see \"" + state.pvTabName + "\".");
}

function saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(PROP_STATE, JSON.stringify(state));
}

function loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_STATE);
  return raw ? JSON.parse(raw) : null;
}

function scheduleContinuation_() {
  // Remove any prior continuation trigger first, then schedule a fresh one.
  const existingId = PropertiesService.getScriptProperties().getProperty(PROP_TRIGGER_ID);
  if (existingId) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getUniqueId() === existingId) ScriptApp.deleteTrigger(t);
    });
  }
  const trigger = ScriptApp.newTrigger("continuePV_")
    .timeBased()
    .after(CONTINUATION_DELAY_SEC * 1000)
    .create();
  PropertiesService.getScriptProperties().setProperty(PROP_TRIGGER_ID, trigger.getUniqueId());
}

// ---------------------------------------------------------------------------
// Config reading
// ---------------------------------------------------------------------------
function getConfig_(ss) {
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + CONFIG_SHEET_NAME + '" not found.');
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valueRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  const colIndex = {};
  headerRow.forEach(function (header, i) {
    colIndex[String(header).trim().toLowerCase()] = i;
  });

  function readValue(name) {
    const idx = colIndex[name.toLowerCase()];
    if (idx === undefined) {
      throw new Error('Column "' + name + '" not found in "' + CONFIG_SHEET_NAME + '" header row.');
    }
    return valueRow[idx];
  }

  const idBoard = String(readValue("IdBoard")).trim();
  const from = readValue("From");
  const to = readValue("To");
  const pvTabName = String(readValue("PVTab")).trim();

  if (!idBoard) throw new Error("IdBoard is empty in PVConfig.");
  if (!(from instanceof Date)) throw new Error("From must be a date/time value in PVConfig.");
  if (!(to instanceof Date)) throw new Error("To must be a date/time value in PVConfig.");
  if (!pvTabName) throw new Error("PVTab is empty in PVConfig.");

  return { idBoard: idBoard, from: from, to: to, pvTabName: pvTabName };
}

// ---------------------------------------------------------------------------
// Target sheet handling
// ---------------------------------------------------------------------------
function getOrCreatePvSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
    sheet.getBandings().forEach(function (b) { b.remove(); });
  } else {
    sheet = ss.insertSheet(name);
  }

  const headers = ["List Name", "Card Name", "Date and Time", "Activity"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  sheet.setFrozenRows(1);
  return sheet;
}

// Appends a batch of rows starting right after the sheet's current last row.
function appendRows_(sheet, rows) {
  const startRow = sheet.getLastRow() + 1;

  const plainValues = rows.map(function (r) { return [r[0], "", r[3], r[4]]; });
  sheet.getRange(startRow, 1, rows.length, 4).setValues(plainValues);

  const richTextValues = rows.map(function (r) {
    return [
      SpreadsheetApp.newRichTextValue()
        .setText(r[1])
        .setLinkUrl(r[2])
        .build()
    ];
  });
  sheet.getRange(startRow, 2, rows.length, 1).setRichTextValues(richTextValues);

  sheet.getRange(startRow, 3, rows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm");
  sheet.getRange(startRow, 4, rows.length, 1).setWrap(true);
  sheet.getRange(startRow, 1, rows.length, 4).setVerticalAlignment("middle");
  sheet.setRowHeights(startRow, rows.length, 24);
}

// Final cosmetic pass once every row is in place.
function finalizeSheet_(sheet) {
  sheet.setColumnWidth(1, 150); // List Name
  sheet.setColumnWidth(2, 220); // Card Name
  sheet.setColumnWidth(3, 130); // Date and Time
  sheet.setColumnWidth(4, 520); // Activity
  sheet.setFrozenRows(1);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(1, 1, lastRow, 4).applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false
    );
  }
}

// ---------------------------------------------------------------------------
// Trello API helpers
// ---------------------------------------------------------------------------
function trelloFetch_(path, params) {
  params = params || {};
  params.key = PRIVATE_Trello_APIkey;
  params.token = PRIVATE_Trello_APItoken;

  const query = Object.keys(params)
    .map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]); })
    .join("&");
  const url = TRELLO_API_BASE + path + "?" + query;

  const MAX_RETRIES = 5;
  let attempt = 0;
  while (true) {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = response.getResponseCode();

    if (code === 200) {
      return JSON.parse(response.getContentText());
    }

    if (code === 429 && attempt < MAX_RETRIES) {
      const retryAfterHeader = response.getHeaders()["Retry-After"] || response.getHeaders()["retry-after"];
      const waitMs = retryAfterHeader ? Number(retryAfterHeader) * 1000 : Math.pow(2, attempt) * 1000;
      Utilities.sleep(waitMs);
      attempt++;
      continue;
    }

    throw new Error("Trello API error (" + code + ") on " + path + ": " + response.getContentText());
  }
}

function getAllLists_(idBoard) {
  // filter=all includes open + archived (closed) lists, returned in board order.
  return trelloFetch_("/boards/" + idBoard + "/lists", { filter: "all", fields: "name,pos" });
}

function getAllCards_(idList) {
  // filter=all includes open + archived cards, returned in board order.
  return trelloFetch_("/lists/" + idList + "/cards", {
    filter: "all",
    fields: "name,shortUrl,url,pos",
    limit: 1000
  });
}

// Fetches actions for up to BATCH_CHUNK_SIZE cards in a single HTTP round trip
// via Trello's /batch endpoint. Returns an array parallel to cardIds, each
// entry already filtered to the [from, to] window.
function getActionsBatch_(cardIds, from, to) {
  const sinceIso = from.toISOString();
  const beforeIso = new Date(to.getTime() + 1000).toISOString(); // +1s to include the "to" instant

  const subpaths = cardIds.map(function (id) {
    return "/cards/" + id + "/actions?filter=all&since=" + encodeURIComponent(sinceIso) +
      "&before=" + encodeURIComponent(beforeIso) + "&limit=1000";
  });

  const batchResult = trelloFetch_("/batch", { urls: subpaths.join(",") });

  return cardIds.map(function (id, i) {
    const item = batchResult[i];
    const statusKey = item ? Object.keys(item)[0] : null;

    if (statusKey === "200") {
      return filterByDate_(item["200"], from, to);
    }
    // Rare per-item failure inside the batch (huge action history, transient error, etc.)
    // Fall back to a direct call for just this one card.
    return filterByDate_(getActionsDirect_(id, from, to), from, to);
  });
}

function getActionsDirect_(idCard, from, to) {
  const sinceIso = from.toISOString();
  const beforeIso = new Date(to.getTime() + 1000).toISOString();
  let actions = [];
  let before = beforeIso;
  let page = 0;
  const MAX_PAGES = 10; // safety cap (10,000 actions per card)

  while (page < MAX_PAGES) {
    const batch = trelloFetch_("/cards/" + idCard + "/actions", {
      filter: "all",
      since: sinceIso,
      before: before,
      limit: 1000
    });
    if (batch.length === 0) break;
    actions = actions.concat(batch);
    if (batch.length < 1000) break;
    before = batch[batch.length - 1].date; // paginate further back in time
    page++;
  }
  return actions;
}

function filterByDate_(actions, from, to) {
  return actions.filter(function (a) {
    const d = new Date(a.date);
    return d >= from && d <= to;
  });
}

// ---------------------------------------------------------------------------
// Turning a raw Trello action into a human-readable line
// ---------------------------------------------------------------------------
function describeAction_(action) {
  const date = new Date(action.date);
  const data = action.data || {};
  let text = null;

  switch (action.type) {
    case "commentCard":
      text = data.text || "(empty comment)";
      break;
    case "updateComment":
      text = "Edited comment: " + (data.action && data.action.text ? data.action.text : "");
      break;
    case "deleteComment":
      text = "Deleted a comment";
      break;
    case "createCard":
      text = "Card created";
      break;
    case "copyCard":
      text = "Card copied" + (data.cardSource ? " from \"" + data.cardSource.name + "\"" : "");
      break;
    case "convertToCardFromCheckItem":
      text = "Converted from a checklist item";
      break;
    case "emailCard":
      text = "Card created/updated by email";
      break;
    case "addAttachmentToCard":
      text = "Added attachment \"" + (data.attachment ? data.attachment.name : "") + "\"";
      break;
    case "deleteAttachmentFromCard":
      text = "Removed an attachment";
      break;
    case "addChecklistToCard":
      text = "Added checklist \"" + (data.checklist ? data.checklist.name : "") + "\"";
      break;
    case "removeChecklistFromCard":
      text = "Removed checklist \"" + (data.checklist ? data.checklist.name : "") + "\"";
      break;
    case "updateCheckItemStateOnCard":
      text = "Checklist item \"" + (data.checkItem ? data.checkItem.name : "") + "\" marked " +
        (data.checkItem && data.checkItem.state === "complete" ? "complete" : "incomplete");
      break;
    case "addMemberToCard":
      text = "Added member " + (data.member ? data.member.name || data.member.username : "");
      break;
    case "removeMemberFromCard":
      text = "Removed member " + (data.member ? data.member.name || data.member.username : "");
      break;
    case "addLabelToCard":
      text = "Added label \"" + (data.label ? (data.label.name || data.label.color) : "") + "\"";
      break;
    case "removeLabelFromCard":
      text = "Removed label \"" + (data.label ? (data.label.name || data.label.color) : "") + "\"";
      break;
    case "updateCard":
      text = describeUpdateCard_(data);
      break;
    default:
      // Unrecognized/rare action type: fall back to a readable version of the raw type.
      text = action.type.replace(/([A-Z])/g, " $1").trim();
      text = text.charAt(0).toUpperCase() + text.slice(1);
  }

  if (text === null) return null;
  return { date: date, text: text };
}

function describeUpdateCard_(data) {
  const old = data.old || {};

  if (data.listBefore && data.listAfter) {
    return "Moved from \"" + data.listBefore.name + "\" to \"" + data.listAfter.name + "\"";
  }
  if (Object.prototype.hasOwnProperty.call(old, "closed")) {
    return data.card && data.card.closed ? "Card archived" : "Card sent back to board (unarchived)";
  }
  if (Object.prototype.hasOwnProperty.call(old, "name")) {
    return "Renamed from \"" + old.name + "\" to \"" + (data.card ? data.card.name : "") + "\"";
  }
  if (Object.prototype.hasOwnProperty.call(old, "desc")) {
    return "Description updated";
  }
  if (Object.prototype.hasOwnProperty.call(old, "due")) {
    if (data.card && data.card.due) {
      return "Due date set to " + Utilities.formatDate(new Date(data.card.due), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    }
    return "Due date removed";
  }
  if (Object.prototype.hasOwnProperty.call(old, "dueComplete")) {
    return "Due date marked " + (data.card && data.card.dueComplete ? "complete" : "incomplete");
  }
  if (Object.prototype.hasOwnProperty.call(old, "pos")) {
    return "Reordered within the list";
  }
  if (Object.prototype.hasOwnProperty.call(old, "idList")) {
    return "Moved to another list";
  }
  return "Card updated";
}