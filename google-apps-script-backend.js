// ============================================================
// UNDERTOW — New Music Backend
// Google Apps Script
// ============================================================

const SPREADSHEET_ID = '1kTlFPVCw4xmo_luHUksCko_WgWzov5VVY53-bE3Qc1g';

function getSheet(name) {
}

function doGet(e) {
  const action = e.parameter.action || 'songs';
  if (action === 'songs' || action === 'results') {
    return jsonResponse(getSongsWithVotes());
  }
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'vote') return jsonResponse(recordVotes(data));
    if (data.action === 'suggest') return jsonResponse(recordSuggestion(data));
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function getSongsWithVotes() {
  const songsSheet = getSheet('Songs');
  const votesSheet = getSheet('Votes');

  const songData = songsSheet.getDataRange().getValues();
  const songs = [];

  for (let i = 1; i < songData.length; i++) {
    const row = songData[i];
    if (!row[1]) continue;
    songs.push({
      id: row[0] || i,
      title: row[1],
      artist: row[2] || '',
      sheet: row[3] || '',
      champion: row[4] || '',
      style: row[5] || '',
      notes: row[6] || '',
      link: row[7] || '',
      votes: { yes: 0, open: 0, no: 0 }
    });
  }

  if (votesSheet.getLastRow() > 1) {
    const voteData = votesSheet.getDataRange().getValues();
    const latestVotes = {};
    for (let i = 1; i < voteData.length; i++) {
      const voter = voteData[i][1];
      const songId = voteData[i][2];
      const vote = voteData[i][3];
      latestVotes[voter + '__' + songId] = vote;
    }
    for (var key in latestVotes) {
      var songId = key.split('__')[1];
      var vote = latestVotes[key];
      var song = songs.find(function(s) { return String(s.id) === String(songId); });
      if (song && song.votes[vote] !== undefined) {
        song.votes[vote]++;
      }
    }
  }

  return { success: true, songs: songs };
}

function recordVotes(data) {
  var sheet = getSheet('Votes');
  var now = new Date();
  var songIds = Object.keys(data.votes);
  for (var i = 0; i < songIds.length; i++) {
    sheet.appendRow([now, data.voter, songIds[i], data.votes[songIds[i]]]);
  }
  return { success: true };
}

function recordSuggestion(data) {
  var sheet = getSheet('Suggestions');
  sheet.appendRow([
    new Date(),
    data.suggestor,
    data.song,
    data.artist,
    data.sheet,
    data.champion,
    data.link,
    data.notes
  ]);
  return { success: true };
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
