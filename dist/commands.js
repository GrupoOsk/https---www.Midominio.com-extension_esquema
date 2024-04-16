/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./node_modules/process/browser.js":
/*!*****************************************!*\
  !*** ./node_modules/process/browser.js ***!
  \*****************************************/
/***/ (function(module) {

// shim for using process in browser
var process = module.exports = {};

// cached from whatever global is present so that test runners that stub it
// don't break things.  But we need to wrap it in a try catch in case it is
// wrapped in strict mode code which doesn't define any globals.  It's inside a
// function because try/catches deoptimize in certain engines.

var cachedSetTimeout;
var cachedClearTimeout;

function defaultSetTimout() {
    throw new Error('setTimeout has not been defined');
}
function defaultClearTimeout () {
    throw new Error('clearTimeout has not been defined');
}
(function () {
    try {
        if (typeof setTimeout === 'function') {
            cachedSetTimeout = setTimeout;
        } else {
            cachedSetTimeout = defaultSetTimout;
        }
    } catch (e) {
        cachedSetTimeout = defaultSetTimout;
    }
    try {
        if (typeof clearTimeout === 'function') {
            cachedClearTimeout = clearTimeout;
        } else {
            cachedClearTimeout = defaultClearTimeout;
        }
    } catch (e) {
        cachedClearTimeout = defaultClearTimeout;
    }
} ())
function runTimeout(fun) {
    if (cachedSetTimeout === setTimeout) {
        //normal enviroments in sane situations
        return setTimeout(fun, 0);
    }
    // if setTimeout wasn't available but was latter defined
    if ((cachedSetTimeout === defaultSetTimout || !cachedSetTimeout) && setTimeout) {
        cachedSetTimeout = setTimeout;
        return setTimeout(fun, 0);
    }
    try {
        // when when somebody has screwed with setTimeout but no I.E. maddness
        return cachedSetTimeout(fun, 0);
    } catch(e){
        try {
            // When we are in I.E. but the script has been evaled so I.E. doesn't trust the global object when called normally
            return cachedSetTimeout.call(null, fun, 0);
        } catch(e){
            // same as above but when it's a version of I.E. that must have the global object for 'this', hopfully our context correct otherwise it will throw a global error
            return cachedSetTimeout.call(this, fun, 0);
        }
    }


}
function runClearTimeout(marker) {
    if (cachedClearTimeout === clearTimeout) {
        //normal enviroments in sane situations
        return clearTimeout(marker);
    }
    // if clearTimeout wasn't available but was latter defined
    if ((cachedClearTimeout === defaultClearTimeout || !cachedClearTimeout) && clearTimeout) {
        cachedClearTimeout = clearTimeout;
        return clearTimeout(marker);
    }
    try {
        // when when somebody has screwed with setTimeout but no I.E. maddness
        return cachedClearTimeout(marker);
    } catch (e){
        try {
            // When we are in I.E. but the script has been evaled so I.E. doesn't  trust the global object when called normally
            return cachedClearTimeout.call(null, marker);
        } catch (e){
            // same as above but when it's a version of I.E. that must have the global object for 'this', hopfully our context correct otherwise it will throw a global error.
            // Some versions of I.E. have different rules for clearTimeout vs setTimeout
            return cachedClearTimeout.call(this, marker);
        }
    }



}
var queue = [];
var draining = false;
var currentQueue;
var queueIndex = -1;

function cleanUpNextTick() {
    if (!draining || !currentQueue) {
        return;
    }
    draining = false;
    if (currentQueue.length) {
        queue = currentQueue.concat(queue);
    } else {
        queueIndex = -1;
    }
    if (queue.length) {
        drainQueue();
    }
}

function drainQueue() {
    if (draining) {
        return;
    }
    var timeout = runTimeout(cleanUpNextTick);
    draining = true;

    var len = queue.length;
    while(len) {
        currentQueue = queue;
        queue = [];
        while (++queueIndex < len) {
            if (currentQueue) {
                currentQueue[queueIndex].run();
            }
        }
        queueIndex = -1;
        len = queue.length;
    }
    currentQueue = null;
    draining = false;
    runClearTimeout(timeout);
}

process.nextTick = function (fun) {
    var args = new Array(arguments.length - 1);
    if (arguments.length > 1) {
        for (var i = 1; i < arguments.length; i++) {
            args[i - 1] = arguments[i];
        }
    }
    queue.push(new Item(fun, args));
    if (queue.length === 1 && !draining) {
        runTimeout(drainQueue);
    }
};

// v8 likes predictible objects
function Item(fun, array) {
    this.fun = fun;
    this.array = array;
}
Item.prototype.run = function () {
    this.fun.apply(null, this.array);
};
process.title = 'browser';
process.browser = true;
process.env = {};
process.argv = [];
process.version = ''; // empty string to avoid regexp issues
process.versions = {};

function noop() {}

process.on = noop;
process.addListener = noop;
process.once = noop;
process.off = noop;
process.removeListener = noop;
process.removeAllListeners = noop;
process.emit = noop;
process.prependListener = noop;
process.prependOnceListener = noop;

process.listeners = function (name) { return [] }

process.binding = function (name) {
    throw new Error('process.binding is not supported');
};

process.cwd = function () { return '/' };
process.chdir = function (dir) {
    throw new Error('process.chdir is not supported');
};
process.umask = function() { return 0; };


/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
!function() {
/*!**********************************!*\
  !*** ./src/commands/commands.js ***!
  \**********************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return e; }; var t, e = {}, r = Object.prototype, n = r.hasOwnProperty, o = Object.defineProperty || function (t, e, r) { t[e] = r.value; }, i = "function" == typeof Symbol ? Symbol : {}, a = i.iterator || "@@iterator", c = i.asyncIterator || "@@asyncIterator", u = i.toStringTag || "@@toStringTag"; function define(t, e, r) { return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e]; } try { define({}, ""); } catch (t) { define = function define(t, e, r) { return t[e] = r; }; } function wrap(t, e, r, n) { var i = e && e.prototype instanceof Generator ? e : Generator, a = Object.create(i.prototype), c = new Context(n || []); return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a; } function tryCatch(t, e, r) { try { return { type: "normal", arg: t.call(e, r) }; } catch (t) { return { type: "throw", arg: t }; } } e.wrap = wrap; var h = "suspendedStart", l = "suspendedYield", f = "executing", s = "completed", y = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var p = {}; define(p, a, function () { return this; }); var d = Object.getPrototypeOf, v = d && d(d(values([]))); v && v !== r && n.call(v, a) && (p = v); var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p); function defineIteratorMethods(t) { ["next", "throw", "return"].forEach(function (e) { define(t, e, function (t) { return this._invoke(e, t); }); }); } function AsyncIterator(t, e) { function invoke(r, o, i, a) { var c = tryCatch(t[r], t, o); if ("throw" !== c.type) { var u = c.arg, h = u.value; return h && "object" == _typeof(h) && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) { invoke("next", t, i, a); }, function (t) { invoke("throw", t, i, a); }) : e.resolve(h).then(function (t) { u.value = t, i(u); }, function (t) { return invoke("throw", t, i, a); }); } a(c.arg); } var r; o(this, "_invoke", { value: function value(t, n) { function callInvokeWithMethodAndArg() { return new e(function (e, r) { invoke(t, n, e, r); }); } return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(e, r, n) { var o = h; return function (i, a) { if (o === f) throw new Error("Generator is already running"); if (o === s) { if ("throw" === i) throw a; return { value: t, done: !0 }; } for (n.method = i, n.arg = a;;) { var c = n.delegate; if (c) { var u = maybeInvokeDelegate(c, n); if (u) { if (u === y) continue; return u; } } if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) { if (o === h) throw o = s, n.arg; n.dispatchException(n.arg); } else "return" === n.method && n.abrupt("return", n.arg); o = f; var p = tryCatch(e, r, n); if ("normal" === p.type) { if (o = n.done ? s : l, p.arg === y) continue; return { value: p.arg, done: n.done }; } "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg); } }; } function maybeInvokeDelegate(e, r) { var n = r.method, o = e.iterator[n]; if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y; var i = tryCatch(o, e.iterator, r.arg); if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y; var a = i.arg; return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y); } function pushTryEntry(t) { var e = { tryLoc: t[0] }; 1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e); } function resetTryEntry(t) { var e = t.completion || {}; e.type = "normal", delete e.arg, t.completion = e; } function Context(t) { this.tryEntries = [{ tryLoc: "root" }], t.forEach(pushTryEntry, this), this.reset(!0); } function values(e) { if (e || "" === e) { var r = e[a]; if (r) return r.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) { var o = -1, i = function next() { for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next; return next.value = t, next.done = !0, next; }; return i.next = i; } } throw new TypeError(_typeof(e) + " is not iterable"); } return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), o(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) { var e = "function" == typeof t && t.constructor; return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name)); }, e.mark = function (t) { return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t; }, e.awrap = function (t) { return { __await: t }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () { return this; }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) { void 0 === i && (i = Promise); var a = new AsyncIterator(wrap(t, r, n, o), i); return e.isGeneratorFunction(r) ? a : a.next().then(function (t) { return t.done ? t.value : a.next(); }); }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () { return this; }), define(g, "toString", function () { return "[object Generator]"; }), e.keys = function (t) { var e = Object(t), r = []; for (var n in e) r.push(n); return r.reverse(), function next() { for (; r.length;) { var t = r.pop(); if (t in e) return next.value = t, next.done = !1, next; } return next.done = !0, next; }; }, e.values = values, Context.prototype = { constructor: Context, reset: function reset(e) { if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t); }, stop: function stop() { this.done = !0; var t = this.tryEntries[0].completion; if ("throw" === t.type) throw t.arg; return this.rval; }, dispatchException: function dispatchException(e) { if (this.done) throw e; var r = this; function handle(n, o) { return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o; } for (var o = this.tryEntries.length - 1; o >= 0; --o) { var i = this.tryEntries[o], a = i.completion; if ("root" === i.tryLoc) return handle("end"); if (i.tryLoc <= this.prev) { var c = n.call(i, "catchLoc"), u = n.call(i, "finallyLoc"); if (c && u) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } else if (c) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); } else { if (!u) throw new Error("try statement without catch or finally"); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } } } }, abrupt: function abrupt(t, e) { for (var r = this.tryEntries.length - 1; r >= 0; --r) { var o = this.tryEntries[r]; if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) { var i = o; break; } } i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null); var a = i ? i.completion : {}; return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a); }, complete: function complete(t, e) { if ("throw" === t.type) throw t.arg; return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y; }, finish: function finish(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y; } }, catch: function _catch(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.tryLoc === t) { var n = r.completion; if ("throw" === n.type) { var o = n.arg; resetTryEntry(r); } return o; } } throw new Error("illegal catch attempt"); }, delegateYield: function delegateYield(e, r, n) { return this.delegate = { iterator: values(e), resultName: r, nextLoc: n }, "next" === this.method && (this.arg = t), y; } }, e; }
function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, _toPropertyKey(descriptor.key), descriptor); } }
function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : String(i); }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }
function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }
/* global global, Office, self, window */

/// <reference path="../jquery.min.js" />
var _require = __webpack_require__(/*! process */ "./node_modules/process/browser.js"),
  send = _require.send;
var intervalId = "";
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    Entorno = new EntornoClase();
  }
  ;
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  var message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Mi Dominio",
    icon: "Icon.32x32",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
;
function Opcion1(_x) {
  return _Opcion.apply(this, arguments);
}
function _Opcion() {
  _Opcion = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee9(event) {
    return _regeneratorRuntime().wrap(function _callee9$(_context9) {
      while (1) switch (_context9.prev = _context9.next) {
        case 0:
          Entorno.OutlookCorreoActual = Office.context.mailbox.item;
          _context9.next = 3;
          return Entorno.OpcionA1();
        case 3:
          event.completed();
        case 4:
        case "end":
          return _context9.stop();
      }
    }, _callee9);
  }));
  return _Opcion.apply(this, arguments);
}
;
function prependHeaderOnSend(_x2) {
  return _prependHeaderOnSend.apply(this, arguments);
}
function _prependHeaderOnSend() {
  _prependHeaderOnSend = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee10(event) {
    return _regeneratorRuntime().wrap(function _callee10$(_context10) {
      while (1) switch (_context10.prev = _context10.next) {
        case 0:
          Entorno.OutlookCorreoActual = Office.context.mailbox.item;
          _context10.next = 3;
          return Entorno.Correo.SetDatosGenerarTexto();
        case 3:
          if (TokenCliente) {
            Office.context.mailbox.item.body.getTypeAsync({
              asyncContext: event
            }, function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                if (asyncResult.error) {
                  Entorno.MostrarMensaje("No se pudo pegar el texto", "generartexto");
                } else {
                  Entorno.MostrarMensaje("No se pudo pegar el texto", "generartexto");
                }

                // Si la operación falla, establecer un intervalo para volver a intentarlo cada minuto
                intervalId = setInterval(function () {
                  prependHeaderOnSend(event);
                }, 60000);
                return;
              }

              // Si la operación tiene éxito, pegar el texto y detener el intervalo
              clearInterval(intervalId);
              Entorno.PegarTexto(asyncResult);
            });
          } else {

            //Identificarse
          }
          ;
        case 5:
        case "end":
          return _context10.stop();
      }
    }, _callee10);
  }));
  return _prependHeaderOnSend.apply(this, arguments);
}
Office.actions.associate("prependHeaderOnSend", prependHeaderOnSend);
var MiDominioClase = /*#__PURE__*/function () {
  function MiDominioClase() {
    _classCallCheck(this, MiDominioClase);
  }
  _createClass(MiDominioClase, [{
    key: "GetTextoExtension",
    value: function () {
      var _GetTextoExtension = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee() {
        var Resultado;
        return _regeneratorRuntime().wrap(function _callee$(_context) {
          while (1) switch (_context.prev = _context.next) {
            case 0:
              Entorno.OutlookCorreoActual = Office.context.mailbox.item;
              _context.next = 3;
              return Entorno.Correo.SetDatosGenerarTexto();
            case 3:
              Resultado = Entorno.Correo.Asunto;
              return _context.abrupt("return", Resultado);
            case 5:
            case "end":
              return _context.stop();
          }
        }, _callee);
      }));
      function GetTextoExtension() {
        return _GetTextoExtension.apply(this, arguments);
      }
      return GetTextoExtension;
    }()
  }]);
  return MiDominioClase;
}();
;
var CorreoClase = /*#__PURE__*/function () {
  function CorreoClase() {
    _classCallCheck(this, CorreoClase);
    this.Asunto = "";
    this.CuerpoHTML = "";
    this.Texto = "";
  }
  _createClass(CorreoClase, [{
    key: "SetDatosGenerarTexto",
    value: function () {
      var _SetDatosGenerarTexto = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2() {
        return _regeneratorRuntime().wrap(function _callee2$(_context2) {
          while (1) switch (_context2.prev = _context2.next) {
            case 0:
              _context2.prev = 0;
              Entorno.Correo.Asunto = "";
              Entorno.Correo.CuerpoHTML = "";
              _context2.next = 5;
              return new Promise(function (resolve, reject) {
                Entorno.OutlookCorreoActual.subject.getAsync(function (result) {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                  } else {
                    reject('');
                  }
                });
              });
            case 5:
              Entorno.Correo.Asunto = _context2.sent;
              _context2.next = 11;
              break;
            case 8:
              _context2.prev = 8;
              _context2.t0 = _context2["catch"](0);
              Entorno.MostrarMensaje("no se pudo pegar el texto", "generartexto");
            case 11:
            case "end":
              return _context2.stop();
          }
        }, _callee2, null, [[0, 8]]);
      }));
      function SetDatosGenerarTexto() {
        return _SetDatosGenerarTexto.apply(this, arguments);
      }
      return SetDatosGenerarTexto;
    }()
  }]);
  return CorreoClase;
}();
;
var EntornoClase = /*#__PURE__*/function () {
  function EntornoClase() {
    _classCallCheck(this, EntornoClase);
    this.OutlookCorreoActual = null;
    this.Correo = new CorreoClase();
    this.MiDominio = new MiDominioClase();
  }
  _createClass(EntornoClase, [{
    key: "MostrarMensaje",
    value: function () {
      var _MostrarMensaje = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee3(mensajeMostrar, accionid) {
        var message;
        return _regeneratorRuntime().wrap(function _callee3$(_context3) {
          while (1) switch (_context3.prev = _context3.next) {
            case 0:
              message = {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: mensajeMostrar + "  ",
                icon: "Icon.32x32",
                persistent: true
              };
              Office.context.mailbox.item.notificationMessages.replaceAsync(accionid, message);
            case 2:
            case "end":
              return _context3.stop();
          }
        }, _callee3);
      }));
      function MostrarMensaje(_x3, _x4) {
        return _MostrarMensaje.apply(this, arguments);
      }
      return MostrarMensaje;
    }()
  }, {
    key: "MostrarMensajeIcono",
    value: function () {
      var _MostrarMensajeIcono = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee4(mensajeMostrar, accionid, iconoMostrar) {
        var message;
        return _regeneratorRuntime().wrap(function _callee4$(_context4) {
          while (1) switch (_context4.prev = _context4.next) {
            case 0:
              message = {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: mensajeMostrar + "  ",
                icon: iconoMostrar,
                persistent: true
              };
              Office.context.mailbox.item.notificationMessages.replaceAsync(accionid, message);
            case 2:
            case "end":
              return _context4.stop();
          }
        }, _callee4);
      }));
      function MostrarMensajeIcono(_x5, _x6, _x7) {
        return _MostrarMensajeIcono.apply(this, arguments);
      }
      return MostrarMensajeIcono;
    }()
  }, {
    key: "CerrarMensaje",
    value: function () {
      var _CerrarMensaje = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee5(accionid) {
        return _regeneratorRuntime().wrap(function _callee5$(_context5) {
          while (1) switch (_context5.prev = _context5.next) {
            case 0:
              Office.context.mailbox.item.notificationMessages.removeAsync(accionid);
            case 1:
            case "end":
              return _context5.stop();
          }
        }, _callee5);
      }));
      function CerrarMensaje(_x8) {
        return _CerrarMensaje.apply(this, arguments);
      }
      return CerrarMensaje;
    }()
  }, {
    key: "CorregirTexto",
    value: function () {
      var _CorregirTexto = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee6(Texto) {
        var textArea;
        return _regeneratorRuntime().wrap(function _callee6$(_context6) {
          while (1) switch (_context6.prev = _context6.next) {
            case 0:
              textArea = document.createElement('textarea');
              textArea.innerHTML = Texto;
              return _context6.abrupt("return", textArea.value);
            case 3:
            case "end":
              return _context6.stop();
          }
        }, _callee6);
      }));
      function CorregirTexto(_x9) {
        return _CorregirTexto.apply(this, arguments);
      }
      return CorregirTexto;
    }()
  }, {
    key: "OpcionA1",
    value: function () {
      var _OpcionA = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee7() {
        return _regeneratorRuntime().wrap(function _callee7$(_context7) {
          while (1) switch (_context7.prev = _context7.next) {
            case 0:
              Entorno.MostrarMensajeIcono("Opcion A1", "opciona1", "icon.naranja");
            case 1:
            case "end":
              return _context7.stop();
          }
        }, _callee7);
      }));
      function OpcionA1() {
        return _OpcionA.apply(this, arguments);
      }
      return OpcionA1;
    }()
  }, {
    key: "PegarTexto",
    value: function () {
      var _PegarTexto = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee8(Datos) {
        var bodyFormat, texto, textoFinal;
        return _regeneratorRuntime().wrap(function _callee8$(_context8) {
          while (1) switch (_context8.prev = _context8.next) {
            case 0:
              Entorno.MostrarMensaje("Generando texto", "generandotexto");
              bodyFormat = Datos.value;
              _context8.next = 4;
              return Entorno.MiDominio.GetTextoExtension();
            case 4:
              texto = _context8.sent;
              _context8.next = 7;
              return Entorno.CorregirTexto(texto);
            case 7:
              textoFinal = _context8.sent;
              Entorno.CerrarMensaje("generandotexto");
              if (!(textoFinal == '')) {
                _context8.next = 14;
                break;
              }
              Entorno.MostrarMensaje("No se pudp pegar el texto", "generartextorespuesta");
              return _context8.abrupt("return");
            case 14:
              console.log("Office.context.mailbox.item.body", Office.context.mailbox.item.body);
              Office.context.mailbox.item.body.prependOnSendAsync(textoFinal, {
                asyncContext: Datos.asyncContext,
                coercionType: bodyFormat
              }, function (Datos) {
                console.log("Datos", Datos);
                if (Datos.status === Office.AsyncResultStatus.Failed) {
                  Entorno.MostrarMensaje("No se pudp pegar el texto", "generartextorespuesta");
                  return;
                }
                Datos.asyncContext.completed();
              });
              Entorno.MostrarMensaje("Texto generado", "generartextorespuesta");
            case 17:
              ;
            case 18:
            case "end":
              return _context8.stop();
          }
        }, _callee8);
      }));
      function PegarTexto(_x10) {
        return _PegarTexto.apply(this, arguments);
      }
      return PegarTexto;
    }()
  }]);
  return EntornoClase;
}();
;
var Entorno = null;
function getGlobal() {
  return typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : typeof __webpack_require__.g !== "undefined" ? __webpack_require__.g : undefined;
}
var g = getGlobal();
g.action = action;
g.Opcion1 = Opcion1;
g.Entorno = Entorno;
}();
/******/ })()
;
//# sourceMappingURL=commands.js.map