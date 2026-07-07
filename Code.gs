// ============================================================
// ПЕРЕВОЗЧИК — Google Apps Script Backend
// Вставь этот код в Apps Script и задеплой как Web App
// Доступ: Anyone (анонимный)
// ============================================================

const SHEET_NAME = "GameState";
const PLAYERS_SHEET = "Players";

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  let result;

  if (action === "players") {
    result = getPlayers();
  } else if (action === "gameState") {
    result = getGameState();
  } else if (action === "resetGame") {
    result = resetGame();
  } else {
    result = { error: "Unknown action" };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Bad JSON" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = body.action;
  let result;

  if (action === "join") {
    result = joinPlayer(body.name, body.car);
  } else if (action === "submitAnswer") {
    result = submitAnswer(body.id, body.questionId, body.answer, body.fuel);
  } else if (action === "startPassenger") {
    result = setGameState({ state: "passenger", clientId: body.clientId, clientName: body.clientName });
  } else if (action === "startQuestion") {
    result = setGameState({ state: "question", questionId: body.questionId, clientId: body.clientId, trapTitle: body.trapTitle });
  } else if (action === "showResults") {
    result = setGameState({ state: "results", questionId: body.questionId });
  } else if (action === "randomEvent") {
    result = setGameState({ state: "event", eventId: body.eventId, eventTitle: body.eventTitle, eventText: body.eventText, eventEffect: body.eventEffect });
  } else if (action === "applyEvent") {
    result = applyRandomEvent(body.effect);
  } else if (action === "waiting") {
    result = setGameState({ state: "waiting" });
  } else if (action === "showRace") {
    result = setGameState({ state: "race" });
  } else if (action === "showFinalResults") {
    result = setGameState({ state: "finalResults" });
  } else if (action === "resetGame") {
    result = resetGame();
  } else {
    result = { error: "Unknown action" };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Игроки ──────────────────────────────────────────────────

function joinPlayer(name, car) {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["id", "name", "car", "fuel", "answer", "shield", "timestamp"]);
  }

  // Проверяем нет ли уже такого игрока
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      return { id: data[i][0], name: data[i][1], car: data[i][2], fuel: data[i][3] };
    }
  }

  const id = Utilities.getUuid();
  sheet.appendRow([id, name, car, 100, "", false, new Date().toISOString()]);
  return { id, name, car, fuel: 100 };
}

function getPlayers() {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);
  if (sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const players = [];
  for (let i = 1; i < data.length; i++) {
    players.push({
      id: data[i][0],
      name: data[i][1],
      car: data[i][2],
      fuel: data[i][3],
      answer: data[i][4],
      shield: data[i][5]
    });
  }
  return players;
}

function submitAnswer(playerId, questionId, answer, fuelChange) {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === playerId) {
      const hasShield = data[i][5];
      let newFuel = data[i][3];
      let shieldUsed = false;

      if (hasShield && fuelChange < 0) {
        // Щит защищает от одного штрафа — гасим щит именно здесь,
        // сервер — единственный источник истины по щиту.
        sheet.getRange(i + 1, 6).setValue(false);
        shieldUsed = true;
      } else {
        newFuel = Math.max(0, Math.min(100, newFuel + fuelChange));
        sheet.getRange(i + 1, 4).setValue(newFuel);
      }

      sheet.getRange(i + 1, 5).setValue(answer);
      return { success: true, fuel: newFuel, shieldUsed };
    }
  }
  return { error: "Player not found" };
}

function applyRandomEvent(effect) {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);
  if (sheet.getLastRow() < 2) return { success: true };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const newFuel = Math.max(0, Math.min(100, data[i][3] + effect));
    sheet.getRange(i + 1, 4).setValue(newFuel);
  }

  // Турбо-событие (effect === 0) — выдать щиты всем
  if (effect === 0) {
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, 6).setValue(true);
    }
  }

  return { success: true };
}

// ── Состояние игры ───────────────────────────────────────────

function getGameState() {
  const sheet = getOrCreateSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    return { state: "waiting", questionId: null, clientId: null };
  }
  const data = sheet.getRange(1, 1, 1, 10).getValues()[0];
  return {
    state: data[0] || "waiting",
    questionId: data[1] || null,
    clientId: data[2] || null,
    clientName: data[3] || null,
    trapTitle: data[4] || null,
    eventId: data[5] || null,
    eventTitle: data[6] || null,
    eventText: data[7] || null,
    eventEffect: data[8] || null
  };
}

function setGameState(params) {
  const sheet = getOrCreateSheet(SHEET_NAME);

  // Сброс ответов при новом вопросе
  if (params.state === "question") {
    clearAnswers();
  }

  sheet.getRange(1, 1, 1, 10).setValues([[
    params.state || "waiting",
    params.questionId || "",
    params.clientId || "",
    params.clientName || "",
    params.trapTitle || "",
    params.eventId || "",
    params.eventTitle || "",
    params.eventText || "",
    params.eventEffect || "",
    new Date().toISOString()
  ]]);

  return { success: true };
}

function clearAnswers() {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);
  if (sheet.getLastRow() < 2) return;
  const lastRow = sheet.getLastRow();
  sheet.getRange(2, 5, lastRow - 1, 1).setValue("");
}

function resetGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const playersSheet = ss.getSheetByName(PLAYERS_SHEET);
  if (playersSheet) playersSheet.clearContents();

  const stateSheet = ss.getSheetByName(SHEET_NAME);
  if (stateSheet) stateSheet.clearContents();

  return { success: true };
}
