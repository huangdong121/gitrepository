const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');
const { webcrypto } = require('node:crypto');

const html = readFileSync(join(__dirname, 'all_in_one_prototype.html'), 'utf8');
const script = new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1]);

function sharedBrowser() {
  const data = new Map();
  const queues = new Map();
  return {
    failWrites: false,
    storage: {
      getItem: key => data.get(key) ?? null,
      setItem(key, value) { data.set(key, value); },
    },
    locks: {
      request(key, action) {
        const result = (queues.get(key) || Promise.resolve()).then(action);
        queues.set(key, result.catch(() => {}));
        return result;
      },
    },
  };
}

function app(shared = sharedBrowser()) {
  const elements = new Map();
  const listeners = new Map();
  function element(id) {
    if (!elements.has(id)) {
      const classes = new Set();
      elements.set(id, {
        value: id === 'fltSt' ? '全部' : '',
        innerHTML: '', textContent: '', disabled: false,
        classList: {
          add: value => classes.add(value),
          remove: value => classes.delete(value),
          contains: value => classes.has(value),
          toggle(value, enabled) { if (enabled) classes.add(value); else classes.delete(value); },
        },
        addEventListener() {}, click() {}, remove() {},
      });
    }
    return elements.get(id);
  }
  const writeControl = element('testWriteControl');
  const context = vm.createContext({
    document: {
      getElementById: element,
      querySelector: () => null,
      querySelectorAll: selector => selector.includes('[data-biz-action]') ? [writeControl] : [],
      addEventListener: (type, action) => listeners.set(`document:${type}`, action),
      createElement: tag => element(tag),
      body: { appendChild() {} },
    },
    window: { addEventListener: (type, action) => listeners.set(type, action) },
    localStorage: {
      getItem: key => shared.storage.getItem(key),
      setItem(key, value) {
        if (shared.failWrites) throw new Error('Storage full');
        shared.storage.setItem(key, value);
      },
    },
    sessionStorage: { getItem: () => null, setItem() {} },
    navigator: { locks: shared.locks },
    crypto: webcrypto, structuredClone, TextDecoder, Blob, URL,
    setTimeout: () => 0, clearTimeout() {}, confirm: () => true, console,
  });
  script.runInContext(context);
  return {
    context, element, listeners,
    run: source => vm.runInContext(source, context),
    data: source => JSON.parse(JSON.stringify(vm.runInContext(source, context))),
  };
}

function csvFile(text) {
  const bytes = new TextEncoder().encode(text);
  return { name: '调整.csv', size: bytes.length, arrayBuffer: async () => bytes.buffer };
}

test('CSV round-trip preserves all nine businesses, codes, cents and pending status', () => {
  const a = app();
  for (const code of a.data('Object.keys(BIZ)')) {
    assert.deepEqual(a.data(`parseVoucherCSV(voucherCSV(BIZ.${code}),BIZ.${code})`), a.data(`BIZ.${code}.v1`));
  }
});

test('CSV handles quoted text and protects spreadsheet formulas on export', () => {
  const a = app();
  a.run(`BIZ.settle02.v1[0].txt='=1+2, "摘要"\\n第二行'`);
  const csv = a.run('voucherCSV(BIZ.settle02)');
  assert.match(csv, /'=1\+2/);
  assert.equal(a.run('parseVoucherCSV(voucherCSV(BIZ.settle02),BIZ.settle02)[0].txt'), '=1+2, "摘要"\n第二行');
  for (const invalid of ['"unterminated', 'abc"def', '"abc"x,def']) {
    a.context.invalid = invalid;
    assert.throws(() => a.run('parseCSV(invalid)'), /引号/);
  }
});

test('import rejects invalid schema, scope, amounts, dates, IDs and per-voucher imbalance', () => {
  const a = app();
  const mutations = [
    "cells[0][0]='错误表头'",
    "cells[1][0]='000010008'",
    "cells[1][1]='202607'",
    "cells[1][13]='NaN'",
    "cells[1][13]='-1'",
    "cells[1][13]='1.5'",
    "cells[1][13]='999999999999999999'",
    "cells[1][15]='20260832'",
    "cells[1][11]='X'",
    "cells.push(cells[1])",
    "cells[1][13]='1'",
    "cells[1][13]='1';cells[4][13]='1'",
  ];
  for (const mutation of mutations) {
    assert.throws(() => a.run(`{
      const cells=parseCSV(voucherCSV(BIZ.settle02));
      ${mutation};
      parseVoucherCSV(cells.map(r=>r.map(v=>'"'+v.replace(/"/g,'""')+'"').join(',')).join('\\r\\n'),BIZ.settle02);
    }`), undefined, mutation);
  }
  assert.throws(() => a.run("parseVoucherCSV(VOUCHER_HEADER.join(','),BIZ.settle02)"), /1 至/);
});

test('voucher audit only affects pending voucher rows in the selected filter', async () => {
  const a = app();
  await a.run('claimWorkflow()');
  a.element('voucherBatch').value = 'A026';
  a.element('voucherSide').value = 'H';
  a.run('openVoucherAudit()');
  await a.run('batchConfirm(1)');
  assert.deepEqual(a.data('BIZ.settle02.v1.map(r=>r.st)'), ['已通过', '待审核', '待审核', '待审核']);
  assert(a.run("BIZ.settle02.t1.every(r=>r.st==='待审核')"));
  a.element('voucherBatch').value = '';
  a.element('voucherSide').value = '';
  a.run('openVoucherAudit()');
  await a.run('batchConfirm(0)');
  assert.deepEqual(a.data('BIZ.settle02.v1.map(r=>r.st)'), ['已通过', '已驳回', '已驳回', '已驳回']);
  a.run('openVoucherAudit()');
  assert.equal(a.run('batchTarget'), null);
});

test('group Excel batch audit does not change source or voucher rows', async () => {
  const a = app();
  a.run("renderBiz('settle06')");
  await a.run('claimWorkflow()');
  a.run('openBatchExcelAudit()');
  await a.run('batchConfirm(1)');
  assert(a.run("BIZ.settle06.excel_data.every(r=>r.st==='已通过')"));
  assert(a.run("BIZ.settle06.t1.every(r=>r.st==='待审核')"));
  assert(a.run("BIZ.settle06.v1.every(r=>r.st==='待审核')"));
});

test('upload validates first, replaces data, resets approval and persists through refresh', async () => {
  const shared = sharedBrowser();
  const a = app(shared);
  await a.run('claimWorkflow()');
  a.run('openVoucherAudit()');
  await a.run('batchConfirm(1)');
  a.context.file = csvFile(a.run(`voucherCSV({...BIZ.settle02,v1:BIZ.settle02.v1.map(r=>({...r,amt:r.amt+10,txt:'财务调整 <img src=x onerror=alert(1)>'}))})`));
  a.run('openVoucherUpload()');
  await a.run('previewVoucherFile(file)');
  assert.equal(a.element('confirmVoucherUpload').disabled, false);
  await a.run('confirmVoucherUpload()');
  assert(a.run("BIZ.settle02.v1.every(r=>r.st==='待审核')"));
  assert.equal(a.run('BIZ.settle02.v1[0].amt'), 9807551);
  assert.equal(a.run('BIZ.settle02.voucherUpload.name'), '调整.csv');
  assert.doesNotMatch(a.element('t2Body').innerHTML, /<img/);
  assert.match(a.element('t2Body').innerHTML, /&lt;img/);
  const refreshed = app(shared);
  assert.equal(refreshed.run('BIZ.settle02.v1[0].amt'), 9807551);
  assert.equal(refreshed.run('ownsBusiness()'), true);
});

test('invalid uploads leave original vouchers intact and cannot be confirmed', async () => {
  const a = app();
  await a.run('claimWorkflow()');
  const before = a.data('BIZ.settle02.v1');
  a.context.file = csvFile('bad,header\n1,2');
  a.run('openVoucherUpload()');
  await a.run('previewVoucherFile(file)');
  assert.equal(a.element('confirmVoucherUpload').disabled, true);
  assert.match(a.element('voucherUploadError').textContent, /表头/);
  await a.run('confirmVoucherUpload()');
  assert.deepEqual(a.data('BIZ.settle02.v1'), before);
});

test('a pending file read is discarded after changing accounts', async () => {
  const a = app();
  await a.run('claimWorkflow()');
  let finish;
  a.context.file = { name: '调整.csv', size: 100, arrayBuffer: () => new Promise(resolve => { finish = resolve; }) };
  const content = new TextEncoder().encode(a.run('voucherCSV(BIZ.settle02)')).buffer;
  a.run('openVoucherUpload()');
  const preview = a.run('previewVoucherFile(file)');
  a.run("switchAccount('W8406769@XZ')");
  finish(content);
  await preview;
  assert.equal(a.run('pendingImport'), null);
  assert.equal(a.element('confirmVoucherUpload').disabled, true);
});

test('concurrent claims have one owner; other accounts cannot mutate or release it', async () => {
  const shared = sharedBrowser();
  const a = app(shared), b = app(shared);
  b.run("switchAccount('W8406769@XZ')");
  await Promise.all([a.run('claimWorkflow()'), b.run('claimWorkflow()')]);
  assert.equal(a.run('readWorkflow().owner'), 'W8406768@XZ');
  assert.equal(b.run('ownsBusiness()'), false);
  assert.equal(b.element('testWriteControl').disabled, true);
  const before = shared.storage.getItem(a.run('workflowKey()'));
  for (const call of ['openVoucherAudit()', 'batchConfirm(1)', 'doCompose()', 'addEda()', 'releaseWorkflow()', 'confirmVoucherUpload()']) {
    await b.run(call);
  }
  assert.equal(shared.storage.getItem(a.run('workflowKey()')), before);
  a.run("renderBiz('settle06')");
  await a.run('claimWorkflow()');
  b.run("renderBiz('settle06')");
  const count = b.run('BIZ.settle06.excel_data.length');
  await b.run('simulateUpload()');
  assert.equal(b.run('BIZ.settle06.excel_data.length'), count);
});

test('all nine businesses have independent locks and the owner can transfer access', async () => {
  const shared = sharedBrowser();
  const a = app(shared), b = app(shared);
  b.run("switchAccount('W8406769@XZ')");
  for (const code of a.data('Object.keys(BIZ)')) {
    a.run(`renderBiz('${code}')`);
    assert.equal(a.run('ownsBusiness()'), false);
    await a.run('claimWorkflow()');
    assert.equal(a.run('ownsBusiness()'), true);
    b.run(`renderBiz('${code}')`);
    assert.equal(b.run('ownsBusiness()'), false);
    await a.run('releaseWorkflow()');
    await b.run('claimWorkflow()');
    assert.equal(b.run('ownsBusiness()'), true);
  }
  a.run("renderBiz('settle02');BIZ.settle02.period='202609'");
  assert.equal(a.run('readWorkflow().owner'), null);
});

test('queued actions and stale audit snapshots cannot modify changed data', async () => {
  const shared = sharedBrowser();
  const a = app(shared), b = app(shared);
  await a.run('claimWorkflow()');
  a.run('openVoucherAudit()');
  const queued = a.run('batchConfirm(1)');
  a.run("switchAccount('W8406769@XZ')");
  await queued;
  assert(a.run("BIZ.settle02.v1.every(r=>r.st==='待审核')"));
  a.run("switchAccount('W8406768@XZ');openVoucherAudit()");
  await b.run("writeBusiness(()=>{BIZ.settle02.v1[0].txt='其他标签页调整'})");
  await a.run('batchConfirm(1)');
  assert(a.run("BIZ.settle02.v1.every(r=>r.st==='待审核')"));
  assert.match(a.element('toast').textContent, /数据已发生变化/);
});

test('storage notifications synchronize ownership and storage failures block writes', async () => {
  const shared = sharedBrowser();
  const a = app(shared), b = app(shared);
  b.run("switchAccount('W8406769@XZ')");
  await a.run('claimWorkflow()');
  b.listeners.get('storage')({ key: a.run('workflowKey()') });
  assert.match(b.element('workflowBar').innerHTML, /W8406768@XZ 正在出账/);
  const before = a.data('BIZ.settle02.v1');
  shared.failWrites = true;
  await a.run("writeBusiness(()=>{BIZ.settle02.v1[0].txt='不得保存'})");
  assert.deepEqual(a.data('BIZ.settle02.v1'), before);
  assert.equal(a.run('ownsBusiness()'), false);
  assert.equal(a.element('testWriteControl').disabled, true);
});

test('EDA records and logs are business-scoped and require approved vouchers for generation', async () => {
  const a = app();
  await a.run('claimWorkflow()');
  const before = a.run('BIZ.settle02.eda.length');
  await a.run('addEda()');
  assert.equal(a.run('BIZ.settle02.eda.length'), before);
  a.run('openVoucherAudit()');
  await a.run('batchConfirm(1)');
  await a.run('addEda()');
  assert.equal(a.run('BIZ.settle02.eda.length'), before + 1);
  assert.equal(a.run('BIZ.settle01.eda.length'), before);
  assert.equal(a.run('BIZ.settle02.eda[0].v2'), '待校验');
  a.element('fltSt').value = '通知失败';
  a.run('renderT3();showPosting(2)');
  assert.match(a.element('dlgPostBody').innerHTML, /20260512155533;77358;01/);
  assert.doesNotMatch(a.element('dlgPostBody').innerHTML, /20260317094844;84679;01/);
  assert.match(a.element('root').innerHTML, /过账效验/);
  assert.match(a.element('t3Body').innerHTML, /详细日志/);
});
