const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function iterator(items) {
  let index = 0;
  return { hasNext: () => index < items.length, next: () => items[index++] };
}

function folder(name, children = [], parents = []) {
  return { getName: () => name, getUrl: () => 'url:' + name, getFolders: () => iterator(children), getParents: () => iterator(parents) };
}

function setup(names = ['Tabela', 'Raporty', 'PZO', 'ZLEC']) {
  const project = folder('Dowolny projekt');
  const children = names.map(name => folder(name, [], [project]));
  project.getFolders = () => iterator(children);
  const table = children[0];
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getId: () => 'source' }) },
    DriveApp: { getFileById: () => ({ getParents: () => iterator([table]) }) }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync('Backend.gs', 'utf8'), context);
  return context;
}

test('foldery projektu są odnajdywane względem arkusza, bez zależności od wielkości liter', () => {
  const context = setup([' tabela ', ' Raporty ', 'PZO']);
  assert.equal(context.getProjectFolder_('raporty').getName(), ' Raporty ');
  assert.equal(context.getProjectFolder_('pzo').getName(), 'PZO');
});

test('brakujące i zduplikowane foldery powodują błąd', () => {
  assert.throws(() => setup().getProjectFolder_('Rachunki'), /znaleziono 0/);
  assert.throws(() => setup(['Tabela', 'Raporty', ' raporty ']).getProjectFolder_('Raporty'), /znaleziono 2/);
  assert.throws(() => setup(['Tabela', ' tabela ', 'Raporty']).getProjectFolder_('Raporty'), /folderu Tabela/);
  assert.throws(() => setup(['inne', 'Raporty']).getProjectFolder_('Raporty'), /bezpośrednio/);
});

test('zlecenia, protokoły i rachunki korzystają z właściwych folderów projektu', () => {
  const context = setup(['Tabela', 'ZLEC', 'PZO', 'Rachunki']);
  const folders = context.getApplicationReportFolders_();
  assert.equal(folders.zlec.getName(), 'ZLEC');
  assert.equal(folders.pzo.getName(), 'PZO');
  assert.equal(context.getCostFolderUrl(), 'url:Rachunki');
});

test('PDF jest zapisywany w folderze odpowiadającym rodzajowi raportu', () => {
  const context = setup(['Tabela', 'protokoły', 'Raporty']);
  const calls = [];
  context.Utilities = {
    base64Decode: value => value,
    newBlob: (data, type, name) => ({ data, type, name })
  };
  const original = context.getProjectFolder_;
  context.getProjectFolder_ = name => {
    const target = original(name);
    target.createFile = blob => { calls.push({ folder: target.getName(), blob }); return { getId: () => 'saved' }; };
    return target;
  };
  assert.equal(context.saveReportPdf('data:application/pdf;base64,abc', 'karty.pdf', 'cards').fileId, 'saved');
  context.saveReportPdf('data:application/pdf;base64,def', 'raport.pdf', 'report');
  assert.deepEqual(calls.map(call => call.folder), ['protokoły', 'Raporty']);
  assert.throws(() => context.saveReportPdf('invalid', 'x.pdf', 'report'), /Nieprawidłowe/);
  assert.throws(() => context.saveReportPdf('data:application/pdf;base64,abc', 'x.pdf', 'other'), /Nieznany/);
});