// Stand-in for google.script.run so the pages can be hosted outside Apps Script (GitHub Pages).
// Each call is a plain POST to the web app's doPost with no cookies, so the visitor's Google
// sign-in (school account, several accounts at once) never gets involved.
// Usage is unchanged: google.script.run.withSuccessHandler(f).withFailureHandler(g).someFn(args)
var RSG_API = 'https://script.google.com/macros/s/AKfycbwUOM6M2g4yHjqqD-0UlDAzYouubVgm-5HOVdqGNFkdl9d3L6gu6k7Rdang28vo3JEhEQ/exec';

(function () {
  function call(fn, args, ok, fail) {
    // A string body is sent as text/plain, which skips the CORS preflight Apps Script can't answer.
    fetch(RSG_API, { method: 'POST', body: JSON.stringify({ fn: fn, args: args }), credentials: 'omit' })
      .then(function (r) {
        if (!r.ok) throw new Error('Server error ' + r.status + '.');
        return r.json().catch(function () { throw new Error('Unexpected response from the scoring server.'); });
      })
      .then(function (d) {
        // Handlers run outside the promise chain so a bug in one isn't reported as a network error
        setTimeout(function () {
          if ('error' in d) { if (fail) fail(new Error(d.error)); }
          else if (ok) ok(d.result === undefined ? null : d.result);
        });
      }, function (e) {
        setTimeout(function () { if (fail) fail(e); else console.error(e); });
      });
  }

  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () { call(name, [].slice.call(arguments), ok, fail); };
      }
    });
  }

  window.google = { script: { run: runner(null, null) } };
})();
