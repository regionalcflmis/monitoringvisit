/**
 * API.JS — google.script.run compatibility shim
 * ------------------------------------------------------------------
 * The form was originally an Apps Script HTML page, where server calls
 * look like:
 *     google.script.run.withSuccessHandler(ok).withFailureHandler(err).someFn(a, b)
 *
 * As a standalone TWA/PWA the page is served from your own hosting, so it
 * must call Apps Script over HTTPS instead. This shim recreates the exact
 * same `google.script.run` surface on top of fetch() — which means the
 * form's existing JavaScript needs NO changes at all.
 *
 * IMPORTANT — why POST with text/plain:
 * Apps Script does not answer CORS preflight (OPTIONS) requests. Sending
 * Content-Type: text/plain keeps this a "simple request", so the browser
 * skips preflight entirely and the call succeeds.
 */

// ⬇️ PASTE YOUR APPS SCRIPT WEB APP /exec URL HERE (see README step 2)
var API_URL = 'https://script.google.com/macros/s/AKfycbxDveVEj52wG7Owk5icdeMzEiPWrjJtAOjYF8jt84d-HucFfqwxGOZYO-3919jmgU47Lw/exec';

(function () {
  function callServer(fnName, args) {
    return fetch(API_URL, {
      method: 'POST',
      // text/plain avoids the CORS preflight Apps Script can't answer
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: fnName, args: args })
    })
      .then(function (res) { return res.json(); })
      .then(function (payload) {
        if (payload && payload.error) throw new Error(payload.error);
        return payload ? payload.result : null;
      });
  }

  function Runner(successCb, failureCb) {
    this._ok = successCb || function () {};
    this._err = failureCb || function (e) { console.error('Server call failed:', e); };
  }

  Runner.prototype.withSuccessHandler = function (cb) {
    return new Runner(cb, this._err);
  };
  Runner.prototype.withFailureHandler = function (cb) {
    return new Runner(this._ok, cb);
  };

  // Every server function the form calls gets proxied here.
  var SERVER_FUNCTIONS = [
    'getBootstrapData',
    'getFollowUpData',
    'getDashboardData',
    'uploadPhoto',
    'submitReport',
    'checkOfficerPin',
    'checkLoginEmail',
    'checkLoginExempt',
    'getAdminEmployeeList',
    'adminUpdateEmployee',
    'getEmployeeAssignedCFLs',
    'getMySubmissions',
    'requestEditPermission',
    'getEditRequests',
    'respondEditRequest',
    'getApprovedEditForUser',
    'getSubmissionForEdit',
    'adminAddEmployeeRow',
    'adminDeleteEmployeeRow',
    'getLoggedInUser'
  ];

  SERVER_FUNCTIONS.forEach(function (name) {
    Runner.prototype[name] = function () {
      var args = Array.prototype.slice.call(arguments);
      var self = this;
      callServer(name, args)
        .then(function (result) { self._ok(result); })
        .catch(function (err) { self._err(err); });
    };
  });

  // Expose the same global the original Apps Script page used.
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  Object.defineProperty(window.google.script, 'run', {
    get: function () { return new Runner(); }
  });
})();
