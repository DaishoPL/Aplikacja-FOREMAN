const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function report(rows, includeBreaks) {
  const sheet = {
    getLastRow: () => rows.length + 1,
    getRange: (start, column) => column === 1 ? {
      getDisplayValues: () => rows.map(row => row.slice(0, 8)),
      getValues: () => rows.map(row => [row[0]])
    } : null
  };
  const context = { Date, SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'Spis wykonanych prac' ? sheet : null }) } };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync('Backend.gs', 'utf8'), context);
  return context.getTimeByWeekReport(includeBreaks);
}

function row(date, start, end, deck, task, id = '') {
  return [new Date(date + 'T12:00:00'), start, end, deck, '', task, id, 'Technik'];
}

test('raport korzysta z wykonanych prac, sumuje tygodnie i wydziela prace dodatkowe', () => {
  const result = report([
    row('2024-01-02', '22:00', '02:00', 'A', 'Wiring'),
    row('2025-01-01', '08:00', '10:00', 'B', 'Wiring'),
    row('2025-01-01', '08:00', '09:00', 'A', 'Repair - EXTRA JOB'),
    row('2025-01-01', '08:00', '09:00', 'TRAVEL', 'Travel'),
    row('2025-01-01', '08:00', '09:00', 'A', 'ANULOWANE Repair'),
    row('2025-01-01', '08:00', '08:00', 'A', 'Zero')
  ], false);
  assert.deepEqual(Array.from(result.allTasks), ['Travel', 'Wiring']);
  assert.equal(result.decks.A.tasks.Wiring.weeklyHours[1], 4);
  assert.equal(result.decks.B.tasks.Wiring.weeklyHours[1], 2);
  assert.equal(result.extraDeck.tasks['A, Repair - EXTRA JOB'].total, 1);
  assert.equal(result.decks.TRAVEL.tasks.Travel.total, 1);
});

test('opcja przerw uwzględnia nazwy zadań i pokładów, również w pracach dodatkowych', () => {
  const rows = [
    row('2025-01-01', '08:00', '09:00', 'BREAK', 'Rest'),
    row('2025-01-01', '08:00', '09:00', 'A', 'Brake - EXTRA JOB'),
    row('2025-01-01', '08:00', '09:00', 'A', 'przerwa')
  ];
  assert.deepEqual(Array.from(report(rows, false).allWeeks), []);
  const included = report(rows, true);
  assert.equal(included.decks.BREAK.tasks.Rest.total, 1);
  assert.equal(included.extraDeck.tasks['A, Brake - EXTRA JOB'].total, 1);
  assert.equal(included.decks.A.tasks.przerwa.total, 1);
});