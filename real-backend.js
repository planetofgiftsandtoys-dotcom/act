/*******************************************************
 * REAL BACKEND CLIENT SHIM
 * Drop-in replacement for google.script.run that talks
 * to our Express server (/api/call/:fn) instead of
 * running inside Apps Script. Index.html is untouched —
 * it just needs this script tag instead of mock-backend.js.
 *******************************************************/
(function () {
  function makeRunner() {
    let successFn = null, failureFn = null;
    let proxy;
    const runner = {
      withSuccessHandler(fn) { successFn = fn; return proxy; },
      withFailureHandler(fn) { failureFn = fn; return proxy; }
    };
    proxy = new Proxy(runner, {
      get(target, prop) {
        if (prop in target) return target[prop];
        return function (...args) {
          fetch("/api/call/" + prop, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(args)
          })
            .then(r => r.json())
            .then(data => {
              if (data && data.__error) throw new Error(data.__error);
              if (successFn) successFn(data);
            })
            .catch(err => { if (failureFn) failureFn(err); else console.error(err); });
          return makeRunner();
        };
      }
    });
    return proxy;
  }
  window.google = { script: { run: makeRunner() } };
})();
