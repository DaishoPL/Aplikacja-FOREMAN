const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function client() {
  const events = [];
  const content = { innerHTML: '' };
  const container = { classList: { remove() {} } };
  const runner = {
    withSuccessHandler(handler) { this.success = handler; return this; },
    withFailureHandler(handler) { this.failure = handler; return this; },
    saveReportPdf(data, name, type) { events.push(['upload', data, name, type]); }
  };
  const context = {
    document: { addEventListener() {}, getElementById: id => id === 'reportContent' ? content : container },
    google: { script: { run: runner } },
    alert: message => events.push(['alert', message])
  };
  const source = fs.readFileSync('Component_Raporty', 'utf8');
  const script = source.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
  vm.createContext(context);
  vm.runInContext(script, context);
  const pdf = { output: () => 'data:application/pdf;base64,abc', save: name => events.push(['download', name]) };
  return { context, runner, events, pdf, content };
}

test('pobieranie PDF zaczyna się dopiero po potwierdzeniu zapisu', () => {
  const { context, runner, events, pdf } = client();
  context.savePdfToDrive(pdf, 'karty.pdf', 'cards', () => events.push(['finished']));
  assert.deepEqual(events, [['upload', 'data:application/pdf;base64,abc', 'karty.pdf', 'cards']]);
  runner.success({ success: true, fileId: 'id' });
  assert.deepEqual(events.slice(1), [['download', 'karty.pdf'], ['finished']]);
});

test('błąd lub brak potwierdzenia nie rozpoczyna pobierania PDF', () => {
  for (const response of ['failure', 'unconfirmed']) {
    const { context, runner, events, pdf } = client();
    context.savePdfToDrive(pdf, 'raport.pdf', 'report', () => events.push(['finished']));
    if (response === 'failure') runner.failure({ message: 'brak dostępu' });
    else runner.success({ success: true });
    assert.equal(events.some(event => event[0] === 'download'), false);
    assert.equal(events.some(event => event[0] === 'alert'), true);
    assert.equal(events.at(-1)[0], 'finished');
  }
});

test('raport tygodniowy pokazuje same prace dodatkowe', () => {
  const { context, content } = client();
  context.renderTimeByWeekReport({
    decks: {}, allWeeks: [1], allTasks: [], settings: { includeBreaks: false },
    extraTasks: ['A, Repair - EXTRA JOB'],
    extraDeck: { tasks: { 'A, Repair - EXTRA JOB': { total: 1, weeklyHours: { 1: 1 } } } }
  });
  assert.match(content.innerHTML, /Repair - EXTRA JOB/);
  assert.doesNotMatch(content.innerHTML, /Brak zarejestrowanych prac/);
});

test('raport tygodniowy nie dodaje niewykonanego zadania Supervision', () => {
  const { context, content } = client();
  context.renderTimeByWeekReport({
    decks: { A: { tasks: { Wiring: { total: 1, weeklyHours: { 1: 1 } } } } },
    allWeeks: [1], allTasks: ['Wiring'], settings: { includeBreaks: false }, extraTasks: []
  });
  assert.match(content.innerHTML, /Wiring/);
  assert.doesNotMatch(content.innerHTML, /Supervision/);
});