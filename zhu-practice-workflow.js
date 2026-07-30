// zhu-practice-workflow.js — 練習卷五步流程狀態
var ZhuPracticeWorkflow = (function () {
  'use strict';

  var SCHEMA_VERSION = 1;
  var STEPS = [
    { id: 'select', title: '選題' },
    { id: 'prompt', title: '準備短文' },
    { id: 'check', title: '檢查短文' },
    { id: 'draft', title: '編輯草稿' },
    { id: 'download', title: '預覽下載' }
  ];

  function severityOf(problem) {
    if (!problem) { return ''; }
    if (problem.severity === 'blocking' ||
        problem.severity === 'confirm' ||
        problem.severity === 'overridable') {
      return problem.severity;
    }
    return '';
  }

  function hasBlocking(passage) {
    passage = passage || {};
    return (passage.problems || []).some(function (x) {
      return severityOf(x) === 'blocking';
    });
  }

  function passageReady(passage) {
    passage = passage || {};
    if (!passage.doc || !passage.checked || hasBlocking(passage) || !passage.confirmed) {
      return false;
    }
    var soft = (passage.problems || []).some(function (x) {
      var severity = severityOf(x);
      return severity === 'confirm' || severity === 'overridable';
    });
    return !soft || !!passage.overridden;
  }

  function readiness(state) {
    state = state || {};
    var p = state.passage || {};
    var selectDone = !!p.ok;
    var promptDone = selectDone && !!p.promptText;
    var checkDone = promptDone && passageReady(p);
    var draftDone = checkDone && !!state.draft && state.draftValid === true;

    return {
      select: {
        canEnter: true,
        complete: selectDone,
        reason: selectDone ? '' : '至少取得 4 組可用對比字'
      },
      prompt: {
        canEnter: selectDone,
        complete: promptDone,
        reason: selectDone ? '' : '請先選題並取得至少 4 組可用對比字'
      },
      check: {
        canEnter: promptDone,
        complete: checkDone,
        reason: promptDone ? '' : '請先產生短文提示詞'
      },
      draft: {
        canEnter: checkDone,
        complete: draftDone,
        reason: checkDone ? '' : '請先完成短文檢查與教師確認'
      },
      download: {
        canEnter: draftDone,
        complete: draftDone,
        reason: draftDone ? '' : '請先修正草稿格式'
      }
    };
  }

  function steps() {
    return STEPS.map(function (x) {
      return { id: x.id, title: x.title };
    });
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function isPlainObject(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
  }

  function isValidState(state) {
    if (!isPlainObject(state) || !isPlainObject(state.passage)) {
      return false;
    }
    if (state.passage.problems !== undefined &&
        !Array.isArray(state.passage.problems)) {
      return false;
    }
    if (state.currentStep !== undefined &&
        !STEPS.some(function (step) { return step.id === state.currentStep; })) {
      return false;
    }
    return true;
  }

  function resetPassageAfterSourceChange(passage) {
    passage.promptText = '';
    passage.doc = null;
    passage.problems = [];
    passage.checked = false;
    passage.confirmed = false;
    passage.overridden = false;
  }

  function invalidate(source, changeType) {
    var next = clone(isPlainObject(source) ? source : {});
    if (!isPlainObject(next.passage)) {
      next.passage = {};
    }
    next.notice = '';

    if (changeType === 'selection' || changeType === 'grade' || changeType === 'override') {
      resetPassageAfterSourceChange(next.passage);
      next.draft = null;
      next.draftValid = false;
      next.currentStep = 'select';
      next.notice = '內容已變更，請重新產生提示詞並檢查；原始短文已保留。';
    } else if (changeType === 'raw') {
      next.passage.doc = null;
      next.passage.problems = [];
      next.passage.checked = false;
      next.passage.confirmed = false;
      next.passage.overridden = false;
      next.draft = null;
      next.draftValid = false;
      next.currentStep = 'check';
    } else if (changeType === 'draft') {
      next.draftValid = false;
      next.currentStep = 'draft';
    }

    return next;
  }

  function serialize(state) {
    return JSON.stringify({
      schemaVersion: SCHEMA_VERSION,
      state: clone(state)
    });
  }

  function restore(raw) {
    var parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (error) {
      return { ok: false, reason: '上次草稿無法還原' };
    }
    if (!isPlainObject(parsed) || parsed.schemaVersion !== SCHEMA_VERSION ||
        !isValidState(parsed.state)) {
      return { ok: false, reason: '上次草稿版本不相容' };
    }
    return { ok: true, state: parsed.state };
  }

  var api = {
    steps: steps,
    readiness: readiness,
    severityOf: severityOf,
    passageReady: passageReady,
    invalidate: invalidate,
    serialize: serialize,
    restore: restore
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
