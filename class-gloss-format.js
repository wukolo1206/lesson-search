(function (global) {
  'use strict';

  const BOPOMOFO_TOKEN = /[\u3105-\u312f\u02ca\u02c7\u02cb\u02d9]+/g;

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatText(value) {
    return escapeHtml(value).replace(
      BOPOMOFO_TOKEN,
      '<span class="bopomofo-token">$&</span>'
    );
  }

  global.ClassGlossFormat = Object.freeze({ formatText });
})(typeof window !== 'undefined' ? window : globalThis);
