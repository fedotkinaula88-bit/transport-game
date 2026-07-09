// ============================================================
// ПЕРЕВОЗЧИК — Google Apps Script Backend
// Вставь этот код в Apps Script и задеплой как Web App
// Доступ: Anyone (анонимный)
// ============================================================

const SHEET_NAME = "GameState";
const PLAYERS_SHEET = "Players";
const POLL_SHEET = "PollVotes";

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
  } else if (action === "pollVotes") {
    result = getPollVotes();
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
    result = submitAnswer(body.id, body.questionId, body.answer, body.fuel, body.type);
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
  } else if (action === "startPoll") {
    result = startPoll(body.pollQuestion, body.pollOptions);
  } else if (action === "submitPollVote") {
    result = submitPollVote(body.id, body.name, body.choice);
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
// Колонки Players: A id, B name, C car, D fuel, E answer, F shield,
// G timestamp, H fastCount, I trafficCount, J crashCount

function joinPlayer(name, car) {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["id", "name", "car", "fuel", "answer", "shield", "timestamp", "fastCount", "trafficCount", "crashCount"]);
  }

  // Проверяем нет ли уже такого игрока
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      return { id: data[i][0], name: data[i][1], car: data[i][2], fuel: data[i][3] };
    }
  }

  const id = Utilities.getUuid();
  sheet.appendRow([id, name, car, 100, "", false, new Date().toISOString(), 0, 0, 0]);
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
      shield: data[i][5],
      fastCount: data[i][7] || 0,
      trafficCount: data[i][8] || 0,
      crashCount: data[i][9] || 0
    });
  }
  return players;
}

function submitAnswer(playerId, questionId, answer, fuelChange, answerType) {
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

      // Считаем к какому стилю ответов тянется игрок (для итогового отчёта)
      if (answerType === "fast") {
        sheet.getRange(i + 1, 8).setValue((data[i][7] || 0) + 1);
      } else if (answerType === "traffic") {
        sheet.getRange(i + 1, 9).setValue((data[i][8] || 0) + 1);
      } else if (answerType === "crash") {
        sheet.getRange(i + 1, 10).setValue((data[i][9] || 0) + 1);
      }

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
// Колонки GameState: A state, B questionId, C clientId, D clientName,
// E trapTitle, F eventId, G eventTitle, H eventText, I eventEffect,
// J timestamp, K pollQuestion, L pollOptionsJSON

function getGameState() {
  const sheet = getOrCreateSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    return { state: "waiting", questionId: null, clientId: null };
  }
  const data = sheet.getRange(1, 1, 1, 12).getValues()[0];
  let pollOptions = [];
  try { pollOptions = data[11] ? JSON.parse(data[11]) : []; } catch (e) { pollOptions = []; }

  return {
    state: data[0] || "waiting",
    questionId: data[1] || null,
    clientId: data[2] || null,
    clientName: data[3] || null,
    trapTitle: data[4] || null,
    eventId: data[5] || null,
    eventTitle: data[6] || null,
    eventText: data[7] || null,
    eventEffect: data[8] || null,
    pollQuestion: data[10] || null,
    pollOptions: pollOptions
  };
}

function setGameState(params) {
  const sheet = getOrCreateSheet(SHEET_NAME);

  // Сброс ответов при новом вопросе
  if (params.state === "question") {
    clearAnswers();
  }

  sheet.getRange(1, 1, 1, 12).setValues([[
    params.state || "waiting",
    params.questionId || "",
    params.clientId || "",
    params.clientName || "",
    params.trapTitle || "",
    params.eventId || "",
    params.eventTitle || "",
    params.eventText || "",
    params.eventEffect || "",
    new Date().toISOString(),
    params.pollQuestion || "",
    params.pollOptions ? JSON.stringify(params.pollOptions) : ""
  ]]);

  return { success: true };
}

function clearAnswers() {
  const sheet = getOrCreateSheet(PLAYERS_SHEET);
  if (sheet.getLastRow() < 2) return;
  const lastRow = sheet.getLastRow();
  sheet.getRange(2, 5, lastRow - 1, 1).setValue("");
}

// ── Опрос с обратной связью ───────────────────────────────────

function startPoll(question, options) {
  const sheet = getOrCreateSheet(POLL_SHEET);
  sheet.clearContents();
  sheet.appendRow(["playerId", "playerName", "choice", "timestamp"]);
  return setGameState({ state: "poll", pollQuestion: question, pollOptions: options });
}

function submitPollVote(playerId, playerName, choice) {
  const sheet = getOrCreateSheet(POLL_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["playerId", "playerName", "choice", "timestamp"]);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === playerId) {
      sheet.getRange(i + 1, 3).setValue(choice);
      sheet.getRange(i + 1, 4).setValue(new Date().toISOString());
      return { success: true };
    }
  }

  sheet.appendRow([playerId, playerName, choice, new Date().toISOString()]);
  return { success: true };
}

function getPollVotes() {
  const sheet = getOrCreateSheet(POLL_SHEET);
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const votes = [];
  for (let i = 1; i < data.length; i++) {
    votes.push({ playerId: data[i][0], playerName: data[i][1], choice: data[i][2] });
  }
  return votes;
}

// ── Сброс ──────────────────────────────────────────────────────

function resetGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const playersSheet = ss.getSheetByName(PLAYERS_SHEET);
  if (playersSheet) playersSheet.clearContents();

  const stateSheet = ss.getSheetByName(SHEET_NAME);
  if (stateSheet) stateSheet.clearContents();

  const pollSheet = ss.getSheetByName(POLL_SHEET);
  if (pollSheet) pollSheet.clearContents();

  return { success: true };
}
