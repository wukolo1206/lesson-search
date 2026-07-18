const test = require('node:test');
const assert = require('node:assert/strict');
const LookupCore = require('../lookup-core.js');

test('queryTerm accepts a valid encoded Chinese term only', () => {
  assert.equal(LookupCore.queryTerm('?q=%E8%98%8B%E6%9E%9C'), '蘋果');
  assert.equal(LookupCore.queryTerm('?q=蘋果'), '蘋果');
  assert.equal(LookupCore.queryTerm(''), '');
  assert.equal(LookupCore.queryTerm('?q=蘋果蘋果蘋果蘋果蘋果蘋果'), '');
  assert.equal(LookupCore.queryTerm('?q=apple'), '');
  assert.equal(LookupCore.queryTerm('?q=%E8%98'), '');
});
