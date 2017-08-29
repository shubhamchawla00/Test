var AdaptiveCards =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 69);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Utilities
//



function _class(obj) { return Object.prototype.toString.call(obj); }

function isString(obj) { return _class(obj) === '[object String]'; }

var _hasOwnProperty = Object.prototype.hasOwnProperty;

function has(object, key) {
  return _hasOwnProperty.call(object, key);
}

// Merge objects
//
function assign(obj /*from1, from2, from3, ...*/) {
  var sources = Array.prototype.slice.call(arguments, 1);

  sources.forEach(function (source) {
    if (!source) { return; }

    if (typeof source !== 'object') {
      throw new TypeError(source + 'must be object');
    }

    Object.keys(source).forEach(function (key) {
      obj[key] = source[key];
    });
  });

  return obj;
}

// Remove element from array and put another array at those position.
// Useful for some operations with tokens
function arrayReplaceAt(src, pos, newElements) {
  return [].concat(src.slice(0, pos), newElements, src.slice(pos + 1));
}

////////////////////////////////////////////////////////////////////////////////

function isValidEntityCode(c) {
  /*eslint no-bitwise:0*/
  // broken sequence
  if (c >= 0xD800 && c <= 0xDFFF) { return false; }
  // never used
  if (c >= 0xFDD0 && c <= 0xFDEF) { return false; }
  if ((c & 0xFFFF) === 0xFFFF || (c & 0xFFFF) === 0xFFFE) { return false; }
  // control codes
  if (c >= 0x00 && c <= 0x08) { return false; }
  if (c === 0x0B) { return false; }
  if (c >= 0x0E && c <= 0x1F) { return false; }
  if (c >= 0x7F && c <= 0x9F) { return false; }
  // out of range
  if (c > 0x10FFFF) { return false; }
  return true;
}

function fromCodePoint(c) {
  /*eslint no-bitwise:0*/
  if (c > 0xffff) {
    c -= 0x10000;
    var surrogate1 = 0xd800 + (c >> 10),
        surrogate2 = 0xdc00 + (c & 0x3ff);

    return String.fromCharCode(surrogate1, surrogate2);
  }
  return String.fromCharCode(c);
}


var UNESCAPE_MD_RE  = /\\([!"#$%&'()*+,\-.\/:;<=>?@[\\\]^_`{|}~])/g;
var ENTITY_RE       = /&([a-z#][a-z0-9]{1,31});/gi;
var UNESCAPE_ALL_RE = new RegExp(UNESCAPE_MD_RE.source + '|' + ENTITY_RE.source, 'gi');

var DIGITAL_ENTITY_TEST_RE = /^#((?:x[a-f0-9]{1,8}|[0-9]{1,8}))/i;

var entities = __webpack_require__(5);

function replaceEntityPattern(match, name) {
  var code = 0;

  if (has(entities, name)) {
    return entities[name];
  }

  if (name.charCodeAt(0) === 0x23/* # */ && DIGITAL_ENTITY_TEST_RE.test(name)) {
    code = name[1].toLowerCase() === 'x' ?
      parseInt(name.slice(2), 16)
    :
      parseInt(name.slice(1), 10);
    if (isValidEntityCode(code)) {
      return fromCodePoint(code);
    }
  }

  return match;
}

/*function replaceEntities(str) {
  if (str.indexOf('&') < 0) { return str; }

  return str.replace(ENTITY_RE, replaceEntityPattern);
}*/

function unescapeMd(str) {
  if (str.indexOf('\\') < 0) { return str; }
  return str.replace(UNESCAPE_MD_RE, '$1');
}

function unescapeAll(str) {
  if (str.indexOf('\\') < 0 && str.indexOf('&') < 0) { return str; }

  return str.replace(UNESCAPE_ALL_RE, function (match, escaped, entity) {
    if (escaped) { return escaped; }
    return replaceEntityPattern(match, entity);
  });
}

////////////////////////////////////////////////////////////////////////////////

var HTML_ESCAPE_TEST_RE = /[&<>"]/;
var HTML_ESCAPE_REPLACE_RE = /[&<>"]/g;
var HTML_REPLACEMENTS = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;'
};

function replaceUnsafeChar(ch) {
  return HTML_REPLACEMENTS[ch];
}

function escapeHtml(str) {
  if (HTML_ESCAPE_TEST_RE.test(str)) {
    return str.replace(HTML_ESCAPE_REPLACE_RE, replaceUnsafeChar);
  }
  return str;
}

////////////////////////////////////////////////////////////////////////////////

var REGEXP_ESCAPE_RE = /[.?*+^$[\]\\(){}|-]/g;

function escapeRE(str) {
  return str.replace(REGEXP_ESCAPE_RE, '\\$&');
}

////////////////////////////////////////////////////////////////////////////////

function isSpace(code) {
  switch (code) {
    case 0x09:
    case 0x20:
      return true;
  }
  return false;
}

// Zs (unicode class) || [\t\f\v\r\n]
function isWhiteSpace(code) {
  if (code >= 0x2000 && code <= 0x200A) { return true; }
  switch (code) {
    case 0x09: // \t
    case 0x0A: // \n
    case 0x0B: // \v
    case 0x0C: // \f
    case 0x0D: // \r
    case 0x20:
    case 0xA0:
    case 0x1680:
    case 0x202F:
    case 0x205F:
    case 0x3000:
      return true;
  }
  return false;
}

////////////////////////////////////////////////////////////////////////////////

/*eslint-disable max-len*/
var UNICODE_PUNCT_RE = __webpack_require__(3);

// Currently without astral characters support.
function isPunctChar(ch) {
  return UNICODE_PUNCT_RE.test(ch);
}


// Markdown ASCII punctuation characters.
//
// !, ", #, $, %, &, ', (, ), *, +, ,, -, ., /, :, ;, <, =, >, ?, @, [, \, ], ^, _, `, {, |, }, or ~
// http://spec.commonmark.org/0.15/#ascii-punctuation-character
//
// Don't confuse with unicode punctuation !!! It lacks some chars in ascii range.
//
function isMdAsciiPunct(ch) {
  switch (ch) {
    case 0x21/* ! */:
    case 0x22/* " */:
    case 0x23/* # */:
    case 0x24/* $ */:
    case 0x25/* % */:
    case 0x26/* & */:
    case 0x27/* ' */:
    case 0x28/* ( */:
    case 0x29/* ) */:
    case 0x2A/* * */:
    case 0x2B/* + */:
    case 0x2C/* , */:
    case 0x2D/* - */:
    case 0x2E/* . */:
    case 0x2F/* / */:
    case 0x3A/* : */:
    case 0x3B/* ; */:
    case 0x3C/* < */:
    case 0x3D/* = */:
    case 0x3E/* > */:
    case 0x3F/* ? */:
    case 0x40/* @ */:
    case 0x5B/* [ */:
    case 0x5C/* \ */:
    case 0x5D/* ] */:
    case 0x5E/* ^ */:
    case 0x5F/* _ */:
    case 0x60/* ` */:
    case 0x7B/* { */:
    case 0x7C/* | */:
    case 0x7D/* } */:
    case 0x7E/* ~ */:
      return true;
    default:
      return false;
  }
}

// Hepler to unify [reference labels].
//
function normalizeReference(str) {
  // use .toUpperCase() instead of .toLowerCase()
  // here to avoid a conflict with Object.prototype
  // members (most notably, `__proto__`)
  return str.trim().replace(/\s+/g, ' ').toUpperCase();
}

////////////////////////////////////////////////////////////////////////////////

// Re-export libraries commonly used in both markdown-it and its plugins,
// so plugins won't have to depend on them explicitly, which reduces their
// bundled size (e.g. a browser build).
//
exports.lib                 = {};
exports.lib.mdurl           = __webpack_require__(9);
exports.lib.ucmicro         = __webpack_require__(72);

exports.assign              = assign;
exports.isString            = isString;
exports.has                 = has;
exports.unescapeMd          = unescapeMd;
exports.unescapeAll         = unescapeAll;
exports.isValidEntityCode   = isValidEntityCode;
exports.fromCodePoint       = fromCodePoint;
// exports.replaceEntities     = replaceEntities;
exports.escapeHtml          = escapeHtml;
exports.arrayReplaceAt      = arrayReplaceAt;
exports.isSpace             = isSpace;
exports.isWhiteSpace        = isWhiteSpace;
exports.isMdAsciiPunct      = isMdAsciiPunct;
exports.isPunctChar         = isPunctChar;
exports.escapeRE            = escapeRE;
exports.normalizeReference  = normalizeReference;


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/**
 * class Ruler
 *
 * Helper class, used by [[MarkdownIt#core]], [[MarkdownIt#block]] and
 * [[MarkdownIt#inline]] to manage sequences of functions (rules):
 *
 * - keep rules in defined order
 * - assign the name to each rule
 * - enable/disable rules
 * - add/replace rules
 * - allow assign rules to additional named chains (in the same)
 * - cacheing lists of active rules
 *
 * You will not need use this class directly until write plugins. For simple
 * rules control use [[MarkdownIt.disable]], [[MarkdownIt.enable]] and
 * [[MarkdownIt.use]].
 **/



/**
 * new Ruler()
 **/
function Ruler() {
  // List of added rules. Each element is:
  //
  // {
  //   name: XXX,
  //   enabled: Boolean,
  //   fn: Function(),
  //   alt: [ name2, name3 ]
  // }
  //
  this.__rules__ = [];

  // Cached rule chains.
  //
  // First level - chain name, '' for default.
  // Second level - diginal anchor for fast filtering by charcodes.
  //
  this.__cache__ = null;
}

////////////////////////////////////////////////////////////////////////////////
// Helper methods, should not be used directly


// Find rule index by name
//
Ruler.prototype.__find__ = function (name) {
  for (var i = 0; i < this.__rules__.length; i++) {
    if (this.__rules__[i].name === name) {
      return i;
    }
  }
  return -1;
};


// Build rules lookup cache
//
Ruler.prototype.__compile__ = function () {
  var self = this;
  var chains = [ '' ];

  // collect unique names
  self.__rules__.forEach(function (rule) {
    if (!rule.enabled) { return; }

    rule.alt.forEach(function (altName) {
      if (chains.indexOf(altName) < 0) {
        chains.push(altName);
      }
    });
  });

  self.__cache__ = {};

  chains.forEach(function (chain) {
    self.__cache__[chain] = [];
    self.__rules__.forEach(function (rule) {
      if (!rule.enabled) { return; }

      if (chain && rule.alt.indexOf(chain) < 0) { return; }

      self.__cache__[chain].push(rule.fn);
    });
  });
};


/**
 * Ruler.at(name, fn [, options])
 * - name (String): rule name to replace.
 * - fn (Function): new rule function.
 * - options (Object): new rule options (not mandatory).
 *
 * Replace rule by name with new function & options. Throws error if name not
 * found.
 *
 * ##### Options:
 *
 * - __alt__ - array with names of "alternate" chains.
 *
 * ##### Example
 *
 * Replace existing typorgapher replacement rule with new one:
 *
 * ```javascript
 * var md = require('markdown-it')();
 *
 * md.core.ruler.at('replacements', function replace(state) {
 *   //...
 * });
 * ```
 **/
Ruler.prototype.at = function (name, fn, options) {
  var index = this.__find__(name);
  var opt = options || {};

  if (index === -1) { throw new Error('Parser rule not found: ' + name); }

  this.__rules__[index].fn = fn;
  this.__rules__[index].alt = opt.alt || [];
  this.__cache__ = null;
};


/**
 * Ruler.before(beforeName, ruleName, fn [, options])
 * - beforeName (String): new rule will be added before this one.
 * - ruleName (String): name of added rule.
 * - fn (Function): rule function.
 * - options (Object): rule options (not mandatory).
 *
 * Add new rule to chain before one with given name. See also
 * [[Ruler.after]], [[Ruler.push]].
 *
 * ##### Options:
 *
 * - __alt__ - array with names of "alternate" chains.
 *
 * ##### Example
 *
 * ```javascript
 * var md = require('markdown-it')();
 *
 * md.block.ruler.before('paragraph', 'my_rule', function replace(state) {
 *   //...
 * });
 * ```
 **/
Ruler.prototype.before = function (beforeName, ruleName, fn, options) {
  var index = this.__find__(beforeName);
  var opt = options || {};

  if (index === -1) { throw new Error('Parser rule not found: ' + beforeName); }

  this.__rules__.splice(index, 0, {
    name: ruleName,
    enabled: true,
    fn: fn,
    alt: opt.alt || []
  });

  this.__cache__ = null;
};


/**
 * Ruler.after(afterName, ruleName, fn [, options])
 * - afterName (String): new rule will be added after this one.
 * - ruleName (String): name of added rule.
 * - fn (Function): rule function.
 * - options (Object): rule options (not mandatory).
 *
 * Add new rule to chain after one with given name. See also
 * [[Ruler.before]], [[Ruler.push]].
 *
 * ##### Options:
 *
 * - __alt__ - array with names of "alternate" chains.
 *
 * ##### Example
 *
 * ```javascript
 * var md = require('markdown-it')();
 *
 * md.inline.ruler.after('text', 'my_rule', function replace(state) {
 *   //...
 * });
 * ```
 **/
Ruler.prototype.after = function (afterName, ruleName, fn, options) {
  var index = this.__find__(afterName);
  var opt = options || {};

  if (index === -1) { throw new Error('Parser rule not found: ' + afterName); }

  this.__rules__.splice(index + 1, 0, {
    name: ruleName,
    enabled: true,
    fn: fn,
    alt: opt.alt || []
  });

  this.__cache__ = null;
};

/**
 * Ruler.push(ruleName, fn [, options])
 * - ruleName (String): name of added rule.
 * - fn (Function): rule function.
 * - options (Object): rule options (not mandatory).
 *
 * Push new rule to the end of chain. See also
 * [[Ruler.before]], [[Ruler.after]].
 *
 * ##### Options:
 *
 * - __alt__ - array with names of "alternate" chains.
 *
 * ##### Example
 *
 * ```javascript
 * var md = require('markdown-it')();
 *
 * md.core.ruler.push('my_rule', function replace(state) {
 *   //...
 * });
 * ```
 **/
Ruler.prototype.push = function (ruleName, fn, options) {
  var opt = options || {};

  this.__rules__.push({
    name: ruleName,
    enabled: true,
    fn: fn,
    alt: opt.alt || []
  });

  this.__cache__ = null;
};


/**
 * Ruler.enable(list [, ignoreInvalid]) -> Array
 * - list (String|Array): list of rule names to enable.
 * - ignoreInvalid (Boolean): set `true` to ignore errors when rule not found.
 *
 * Enable rules with given names. If any rule name not found - throw Error.
 * Errors can be disabled by second param.
 *
 * Returns list of found rule names (if no exception happened).
 *
 * See also [[Ruler.disable]], [[Ruler.enableOnly]].
 **/
Ruler.prototype.enable = function (list, ignoreInvalid) {
  if (!Array.isArray(list)) { list = [ list ]; }

  var result = [];

  // Search by name and enable
  list.forEach(function (name) {
    var idx = this.__find__(name);

    if (idx < 0) {
      if (ignoreInvalid) { return; }
      throw new Error('Rules manager: invalid rule name ' + name);
    }
    this.__rules__[idx].enabled = true;
    result.push(name);
  }, this);

  this.__cache__ = null;
  return result;
};


/**
 * Ruler.enableOnly(list [, ignoreInvalid])
 * - list (String|Array): list of rule names to enable (whitelist).
 * - ignoreInvalid (Boolean): set `true` to ignore errors when rule not found.
 *
 * Enable rules with given names, and disable everything else. If any rule name
 * not found - throw Error. Errors can be disabled by second param.
 *
 * See also [[Ruler.disable]], [[Ruler.enable]].
 **/
Ruler.prototype.enableOnly = function (list, ignoreInvalid) {
  if (!Array.isArray(list)) { list = [ list ]; }

  this.__rules__.forEach(function (rule) { rule.enabled = false; });

  this.enable(list, ignoreInvalid);
};


/**
 * Ruler.disable(list [, ignoreInvalid]) -> Array
 * - list (String|Array): list of rule names to disable.
 * - ignoreInvalid (Boolean): set `true` to ignore errors when rule not found.
 *
 * Disable rules with given names. If any rule name not found - throw Error.
 * Errors can be disabled by second param.
 *
 * Returns list of found rule names (if no exception happened).
 *
 * See also [[Ruler.enable]], [[Ruler.enableOnly]].
 **/
Ruler.prototype.disable = function (list, ignoreInvalid) {
  if (!Array.isArray(list)) { list = [ list ]; }

  var result = [];

  // Search by name and disable
  list.forEach(function (name) {
    var idx = this.__find__(name);

    if (idx < 0) {
      if (ignoreInvalid) { return; }
      throw new Error('Rules manager: invalid rule name ' + name);
    }
    this.__rules__[idx].enabled = false;
    result.push(name);
  }, this);

  this.__cache__ = null;
  return result;
};


/**
 * Ruler.getRules(chainName) -> Array
 *
 * Return array of active functions (rules) for given chain name. It analyzes
 * rules configuration, compiles caches if not exists and returns result.
 *
 * Default chain name is `''` (empty string). It can't be skipped. That's
 * done intentionally, to keep signature monomorphic for high speed.
 **/
Ruler.prototype.getRules = function (chainName) {
  if (this.__cache__ === null) {
    this.__compile__();
  }

  // Chain can be empty, if rules disabled. But we still have to return Array.
  return this.__cache__[chainName] || [];
};

module.exports = Ruler;


/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Token class




/**
 * class Token
 **/

/**
 * new Token(type, tag, nesting)
 *
 * Create new token and fill passed properties.
 **/
function Token(type, tag, nesting) {
  /**
   * Token#type -> String
   *
   * Type of the token (string, e.g. "paragraph_open")
   **/
  this.type     = type;

  /**
   * Token#tag -> String
   *
   * html tag name, e.g. "p"
   **/
  this.tag      = tag;

  /**
   * Token#attrs -> Array
   *
   * Html attributes. Format: `[ [ name1, value1 ], [ name2, value2 ] ]`
   **/
  this.attrs    = null;

  /**
   * Token#map -> Array
   *
   * Source map info. Format: `[ line_begin, line_end ]`
   **/
  this.map      = null;

  /**
   * Token#nesting -> Number
   *
   * Level change (number in {-1, 0, 1} set), where:
   *
   * -  `1` means the tag is opening
   * -  `0` means the tag is self-closing
   * - `-1` means the tag is closing
   **/
  this.nesting  = nesting;

  /**
   * Token#level -> Number
   *
   * nesting level, the same as `state.level`
   **/
  this.level    = 0;

  /**
   * Token#children -> Array
   *
   * An array of child nodes (inline and img tokens)
   **/
  this.children = null;

  /**
   * Token#content -> String
   *
   * In a case of self-closing tag (code, html, fence, etc.),
   * it has contents of this tag.
   **/
  this.content  = '';

  /**
   * Token#markup -> String
   *
   * '*' or '_' for emphasis, fence string for fence, etc.
   **/
  this.markup   = '';

  /**
   * Token#info -> String
   *
   * fence infostring
   **/
  this.info     = '';

  /**
   * Token#meta -> Object
   *
   * A place for plugins to store an arbitrary data
   **/
  this.meta     = null;

  /**
   * Token#block -> Boolean
   *
   * True for block-level tokens, false for inline tokens.
   * Used in renderer to calculate line breaks
   **/
  this.block    = false;

  /**
   * Token#hidden -> Boolean
   *
   * If it's true, ignore this element when rendering. Used for tight lists
   * to hide paragraphs.
   **/
  this.hidden   = false;
}


/**
 * Token.attrIndex(name) -> Number
 *
 * Search attribute index by name.
 **/
Token.prototype.attrIndex = function attrIndex(name) {
  var attrs, i, len;

  if (!this.attrs) { return -1; }

  attrs = this.attrs;

  for (i = 0, len = attrs.length; i < len; i++) {
    if (attrs[i][0] === name) { return i; }
  }
  return -1;
};


/**
 * Token.attrPush(attrData)
 *
 * Add `[ name, value ]` attribute to list. Init attrs if necessary
 **/
Token.prototype.attrPush = function attrPush(attrData) {
  if (this.attrs) {
    this.attrs.push(attrData);
  } else {
    this.attrs = [ attrData ];
  }
};


/**
 * Token.attrSet(name, value)
 *
 * Set `name` attribute to `value`. Override old value if exists.
 **/
Token.prototype.attrSet = function attrSet(name, value) {
  var idx = this.attrIndex(name),
      attrData = [ name, value ];

  if (idx < 0) {
    this.attrPush(attrData);
  } else {
    this.attrs[idx] = attrData;
  }
};


/**
 * Token.attrGet(name)
 *
 * Get the value of attribute `name`, or null if it does not exist.
 **/
Token.prototype.attrGet = function attrGet(name) {
  var idx = this.attrIndex(name), value = null;
  if (idx >= 0) {
    value = this.attrs[idx][1];
  }
  return value;
};


/**
 * Token.attrJoin(name, value)
 *
 * Join value to existing attribute via space. Or create new attribute if not
 * exists. Useful to operate with token classes.
 **/
Token.prototype.attrJoin = function attrJoin(name, value) {
  var idx = this.attrIndex(name);

  if (idx < 0) {
    this.attrPush([ name, value ]);
  } else {
    this.attrs[idx][1] = this.attrs[idx][1] + ' ' + value;
  }
};


module.exports = Token;


/***/ }),
/* 3 */
/***/ (function(module, exports) {

module.exports=/[!-#%-\*,-/:;\?@\[-\]_\{\}\xA1\xA7\xAB\xB6\xB7\xBB\xBF\u037E\u0387\u055A-\u055F\u0589\u058A\u05BE\u05C0\u05C3\u05C6\u05F3\u05F4\u0609\u060A\u060C\u060D\u061B\u061E\u061F\u066A-\u066D\u06D4\u0700-\u070D\u07F7-\u07F9\u0830-\u083E\u085E\u0964\u0965\u0970\u0AF0\u0DF4\u0E4F\u0E5A\u0E5B\u0F04-\u0F12\u0F14\u0F3A-\u0F3D\u0F85\u0FD0-\u0FD4\u0FD9\u0FDA\u104A-\u104F\u10FB\u1360-\u1368\u1400\u166D\u166E\u169B\u169C\u16EB-\u16ED\u1735\u1736\u17D4-\u17D6\u17D8-\u17DA\u1800-\u180A\u1944\u1945\u1A1E\u1A1F\u1AA0-\u1AA6\u1AA8-\u1AAD\u1B5A-\u1B60\u1BFC-\u1BFF\u1C3B-\u1C3F\u1C7E\u1C7F\u1CC0-\u1CC7\u1CD3\u2010-\u2027\u2030-\u2043\u2045-\u2051\u2053-\u205E\u207D\u207E\u208D\u208E\u2308-\u230B\u2329\u232A\u2768-\u2775\u27C5\u27C6\u27E6-\u27EF\u2983-\u2998\u29D8-\u29DB\u29FC\u29FD\u2CF9-\u2CFC\u2CFE\u2CFF\u2D70\u2E00-\u2E2E\u2E30-\u2E44\u3001-\u3003\u3008-\u3011\u3014-\u301F\u3030\u303D\u30A0\u30FB\uA4FE\uA4FF\uA60D-\uA60F\uA673\uA67E\uA6F2-\uA6F7\uA874-\uA877\uA8CE\uA8CF\uA8F8-\uA8FA\uA8FC\uA92E\uA92F\uA95F\uA9C1-\uA9CD\uA9DE\uA9DF\uAA5C-\uAA5F\uAADE\uAADF\uAAF0\uAAF1\uABEB\uFD3E\uFD3F\uFE10-\uFE19\uFE30-\uFE52\uFE54-\uFE61\uFE63\uFE68\uFE6A\uFE6B\uFF01-\uFF03\uFF05-\uFF0A\uFF0C-\uFF0F\uFF1A\uFF1B\uFF1F\uFF20\uFF3B-\uFF3D\uFF3F\uFF5B\uFF5D\uFF5F-\uFF65]|\uD800[\uDD00-\uDD02\uDF9F\uDFD0]|\uD801\uDD6F|\uD802[\uDC57\uDD1F\uDD3F\uDE50-\uDE58\uDE7F\uDEF0-\uDEF6\uDF39-\uDF3F\uDF99-\uDF9C]|\uD804[\uDC47-\uDC4D\uDCBB\uDCBC\uDCBE-\uDCC1\uDD40-\uDD43\uDD74\uDD75\uDDC5-\uDDC9\uDDCD\uDDDB\uDDDD-\uDDDF\uDE38-\uDE3D\uDEA9]|\uD805[\uDC4B-\uDC4F\uDC5B\uDC5D\uDCC6\uDDC1-\uDDD7\uDE41-\uDE43\uDE60-\uDE6C\uDF3C-\uDF3E]|\uD807[\uDC41-\uDC45\uDC70\uDC71]|\uD809[\uDC70-\uDC74]|\uD81A[\uDE6E\uDE6F\uDEF5\uDF37-\uDF3B\uDF44]|\uD82F\uDC9F|\uD836[\uDE87-\uDE8B]|\uD83A[\uDD5E\uDD5F]/

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var ValidationError;
(function (ValidationError) {
    ValidationError[ValidationError["ActionTypeNotAllowed"] = 0] = "ActionTypeNotAllowed";
    ValidationError[ValidationError["CollectionCantBeEmpty"] = 1] = "CollectionCantBeEmpty";
    ValidationError[ValidationError["ElementTypeNotAllowed"] = 2] = "ElementTypeNotAllowed";
    ValidationError[ValidationError["InteractivityNotAllowed"] = 3] = "InteractivityNotAllowed";
    ValidationError[ValidationError["InvalidPropertyValue"] = 4] = "InvalidPropertyValue";
    ValidationError[ValidationError["MissingCardType"] = 5] = "MissingCardType";
    ValidationError[ValidationError["PropertyCantBeNull"] = 6] = "PropertyCantBeNull";
    ValidationError[ValidationError["TooManyActions"] = 7] = "TooManyActions";
    ValidationError[ValidationError["UnknownActionType"] = 8] = "UnknownActionType";
    ValidationError[ValidationError["UnknownElementType"] = 9] = "UnknownElementType";
    ValidationError[ValidationError["UnsupportedCardVersion"] = 10] = "UnsupportedCardVersion";
})(ValidationError = exports.ValidationError || (exports.ValidationError = {}));


/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// HTML5 entities map: { name -> utf16string }
//


/*eslint quotes:0*/
module.exports = __webpack_require__(16);


/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Regexps to match html elements



var attr_name     = '[a-zA-Z_:][a-zA-Z0-9:._-]*';

var unquoted      = '[^"\'=<>`\\x00-\\x20]+';
var single_quoted = "'[^']*'";
var double_quoted = '"[^"]*"';

var attr_value  = '(?:' + unquoted + '|' + single_quoted + '|' + double_quoted + ')';

var attribute   = '(?:\\s+' + attr_name + '(?:\\s*=\\s*' + attr_value + ')?)';

var open_tag    = '<[A-Za-z][A-Za-z0-9\\-]*' + attribute + '*\\s*\\/?>';

var close_tag   = '<\\/[A-Za-z][A-Za-z0-9\\-]*\\s*>';
var comment     = '<!---->|<!--(?:-?[^>-])(?:-?[^-])*-->';
var processing  = '<[?].*?[?]>';
var declaration = '<![A-Z]+\\s+[^>]*>';
var cdata       = '<!\\[CDATA\\[[\\s\\S]*?\\]\\]>';

var HTML_TAG_RE = new RegExp('^(?:' + open_tag + '|' + close_tag + '|' + comment +
                        '|' + processing + '|' + declaration + '|' + cdata + ')');
var HTML_OPEN_CLOSE_TAG_RE = new RegExp('^(?:' + open_tag + '|' + close_tag + ')');

module.exports.HTML_TAG_RE = HTML_TAG_RE;
module.exports.HTML_OPEN_CLOSE_TAG_RE = HTML_OPEN_CLOSE_TAG_RE;


/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Process *this* and _that_
//



// Insert each marker as a separate text token, and add it to delimiter list
//
module.exports.tokenize = function emphasis(state, silent) {
  var i, scanned, token,
      start = state.pos,
      marker = state.src.charCodeAt(start);

  if (silent) { return false; }

  if (marker !== 0x5F /* _ */ && marker !== 0x2A /* * */) { return false; }

  scanned = state.scanDelims(state.pos, marker === 0x2A);

  for (i = 0; i < scanned.length; i++) {
    token         = state.push('text', '', 0);
    token.content = String.fromCharCode(marker);

    state.delimiters.push({
      // Char code of the starting marker (number).
      //
      marker: marker,

      // Total length of these series of delimiters.
      //
      length: scanned.length,

      // An amount of characters before this one that's equivalent to
      // current one. In plain English: if this delimiter does not open
      // an emphasis, neither do previous `jump` characters.
      //
      // Used to skip sequences like "*****" in one step, for 1st asterisk
      // value will be 0, for 2nd it's 1 and so on.
      //
      jump:   i,

      // A position of the token this delimiter corresponds to.
      //
      token:  state.tokens.length - 1,

      // Token level.
      //
      level:  state.level,

      // If this delimiter is matched as a valid opener, `end` will be
      // equal to its position, otherwise it's `-1`.
      //
      end:    -1,

      // Boolean flags that determine if this delimiter could open or close
      // an emphasis.
      //
      open:   scanned.can_open,
      close:  scanned.can_close
    });
  }

  state.pos += scanned.length;

  return true;
};


// Walk through delimiter list and replace text tokens with tags
//
module.exports.postProcess = function emphasis(state) {
  var i,
      startDelim,
      endDelim,
      token,
      ch,
      isStrong,
      delimiters = state.delimiters,
      max = state.delimiters.length;

  for (i = 0; i < max; i++) {
    startDelim = delimiters[i];

    if (startDelim.marker !== 0x5F/* _ */ && startDelim.marker !== 0x2A/* * */) {
      continue;
    }

    // Process only opening markers
    if (startDelim.end === -1) {
      continue;
    }

    endDelim = delimiters[startDelim.end];

    // If the next delimiter has the same marker and is adjacent to this one,
    // merge those into one strong delimiter.
    //
    // `<em><em>whatever</em></em>` -> `<strong>whatever</strong>`
    //
    isStrong = i + 1 < max &&
               delimiters[i + 1].end === startDelim.end - 1 &&
               delimiters[i + 1].token === startDelim.token + 1 &&
               delimiters[startDelim.end - 1].token === endDelim.token - 1 &&
               delimiters[i + 1].marker === startDelim.marker;

    ch = String.fromCharCode(startDelim.marker);

    token         = state.tokens[startDelim.token];
    token.type    = isStrong ? 'strong_open' : 'em_open';
    token.tag     = isStrong ? 'strong' : 'em';
    token.nesting = 1;
    token.markup  = isStrong ? ch + ch : ch;
    token.content = '';

    token         = state.tokens[endDelim.token];
    token.type    = isStrong ? 'strong_close' : 'em_close';
    token.tag     = isStrong ? 'strong' : 'em';
    token.nesting = -1;
    token.markup  = isStrong ? ch + ch : ch;
    token.content = '';

    if (isStrong) {
      state.tokens[delimiters[i + 1].token].content = '';
      state.tokens[delimiters[startDelim.end - 1].token].content = '';
      i++;
    }
  }
};


/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// ~~strike through~~
//



// Insert each marker as a separate text token, and add it to delimiter list
//
module.exports.tokenize = function strikethrough(state, silent) {
  var i, scanned, token, len, ch,
      start = state.pos,
      marker = state.src.charCodeAt(start);

  if (silent) { return false; }

  if (marker !== 0x7E/* ~ */) { return false; }

  scanned = state.scanDelims(state.pos, true);
  len = scanned.length;
  ch = String.fromCharCode(marker);

  if (len < 2) { return false; }

  if (len % 2) {
    token         = state.push('text', '', 0);
    token.content = ch;
    len--;
  }

  for (i = 0; i < len; i += 2) {
    token         = state.push('text', '', 0);
    token.content = ch + ch;

    state.delimiters.push({
      marker: marker,
      jump:   i,
      token:  state.tokens.length - 1,
      level:  state.level,
      end:    -1,
      open:   scanned.can_open,
      close:  scanned.can_close
    });
  }

  state.pos += scanned.length;

  return true;
};


// Walk through delimiter list and replace text tokens with tags
//
module.exports.postProcess = function strikethrough(state) {
  var i, j,
      startDelim,
      endDelim,
      token,
      loneMarkers = [],
      delimiters = state.delimiters,
      max = state.delimiters.length;

  for (i = 0; i < max; i++) {
    startDelim = delimiters[i];

    if (startDelim.marker !== 0x7E/* ~ */) {
      continue;
    }

    if (startDelim.end === -1) {
      continue;
    }

    endDelim = delimiters[startDelim.end];

    token         = state.tokens[startDelim.token];
    token.type    = 's_open';
    token.tag     = 's';
    token.nesting = 1;
    token.markup  = '~~';
    token.content = '';

    token         = state.tokens[endDelim.token];
    token.type    = 's_close';
    token.tag     = 's';
    token.nesting = -1;
    token.markup  = '~~';
    token.content = '';

    if (state.tokens[endDelim.token - 1].type === 'text' &&
        state.tokens[endDelim.token - 1].content === '~') {

      loneMarkers.push(endDelim.token - 1);
    }
  }

  // If a marker sequence has an odd number of characters, it's splitted
  // like this: `~~~~~` -> `~` + `~~` + `~~`, leaving one marker at the
  // start of the sequence.
  //
  // So, we have to move all those markers after subsequent s_close tags.
  //
  while (loneMarkers.length) {
    i = loneMarkers.pop();
    j = i + 1;

    while (j < state.tokens.length && state.tokens[j].type === 's_close') {
      j++;
    }

    j--;

    if (i !== j) {
      token = state.tokens[j];
      state.tokens[j] = state.tokens[i];
      state.tokens[i] = token;
    }
  }
};


/***/ }),
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";



module.exports.encode = __webpack_require__(65);
module.exports.decode = __webpack_require__(64);
module.exports.format = __webpack_require__(66);
module.exports.parse  = __webpack_require__(67);


/***/ }),
/* 10 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var markdownIt = __webpack_require__(19);
var markdownProcessor = new markdownIt();
function processMarkdown(text) {
    return markdownProcessor.render(text);
}
exports.processMarkdown = processMarkdown;
function getValueOrDefault(obj, defaultValue) {
    return obj ? obj : defaultValue;
}
exports.getValueOrDefault = getValueOrDefault;
function isNullOrEmpty(value) {
    return value === undefined || value === null || value === "";
}
exports.isNullOrEmpty = isNullOrEmpty;
function appendChild(node, child) {
    if (child != null && child != undefined) {
        node.appendChild(child);
    }
}
exports.appendChild = appendChild;
function renderSeparation(separationDefinition, orientation) {
    var separator = document.createElement("div");
    if (orientation == "vertical") {
        if (separationDefinition.lineThickness) {
            separator.style.marginTop = (separationDefinition.spacing / 2) + "px";
            separator.style.paddingTop = (separationDefinition.spacing / 2) + "px";
            separator.style.borderTop = separationDefinition.lineThickness + "px solid " + stringToCssColor(separationDefinition.lineColor);
        }
        else {
            separator.style.height = separationDefinition.spacing + "px";
        }
    }
    else {
        if (separationDefinition.lineThickness) {
            separator.style.marginLeft = (separationDefinition.spacing / 2) + "px";
            separator.style.paddingLeft = (separationDefinition.spacing / 2) + "px";
            separator.style.borderLeft = separationDefinition.lineThickness + "px solid " + stringToCssColor(separationDefinition.lineColor);
        }
        else {
            separator.style.width = separationDefinition.spacing + "px";
        }
    }
    return separator;
}
exports.renderSeparation = renderSeparation;
function stringToCssColor(color) {
    var regEx = /#([0-9A-F]{2})([0-9A-F]{2})([0-9A-F]{2})([0-9A-F]{2})?/gi;
    var matches = regEx.exec(color);
    if (matches && matches[4]) {
        var a = parseInt(matches[1], 16) / 255;
        var r = parseInt(matches[2], 16);
        var g = parseInt(matches[3], 16);
        var b = parseInt(matches[4], 16);
        return "rgba(" + r + "," + g + "," + b + "," + a + ")";
    }
    else {
        return color;
    }
}
exports.stringToCssColor = stringToCssColor;
var StringWithSubstitutions = (function () {
    function StringWithSubstitutions() {
        this._isProcessed = false;
        this._original = null;
        this._processed = null;
    }
    StringWithSubstitutions.prototype.substituteInputValues = function (inputs) {
        this._processed = this._original;
        var regEx = /\{{2}([a-z0-9_$@]+).value\}{2}/gi;
        var matches;
        while ((matches = regEx.exec(this._original)) != null) {
            var matchedInput = null;
            for (var i = 0; i < inputs.length; i++) {
                if (inputs[i].id.toLowerCase() == matches[1].toLowerCase()) {
                    matchedInput = inputs[i];
                    break;
                }
            }
            if (matchedInput) {
                this._processed = this._processed.replace(matches[0], matchedInput.value ? matchedInput.value : "");
            }
        }
        ;
        this._isProcessed = true;
    };
    StringWithSubstitutions.prototype.get = function () {
        if (!this._isProcessed) {
            return this._original;
        }
        else {
            return this._processed;
        }
    };
    StringWithSubstitutions.prototype.set = function (value) {
        this._original = value;
        this._isProcessed = false;
    };
    return StringWithSubstitutions;
}());
exports.StringWithSubstitutions = StringWithSubstitutions;


/***/ }),
/* 11 */
/***/ (function(module, exports) {

module.exports=/[\0-\x1F\x7F-\x9F]/

/***/ }),
/* 12 */
/***/ (function(module, exports) {

module.exports=/[ \xA0\u1680\u2000-\u200A\u202F\u205F\u3000]/

/***/ }),
/* 13 */
/***/ (function(module, exports) {

module.exports=/[\0-\uD7FF\uE000-\uFFFF]|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/

/***/ }),
/* 14 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var Enums = __webpack_require__(4);
var Utils = __webpack_require__(10);
var TextFormatters = __webpack_require__(70);
function invokeSetParent(obj, parent) {
    // This is not super pretty, but it the closest emulation of
    // "internal" in TypeScript.
    obj["setParent"](parent);
}
function isActionAllowed(action, forbiddenActionTypes) {
    if (forbiddenActionTypes) {
        for (var i = 0; i < forbiddenActionTypes.length; i++) {
            if (action.getJsonTypeName() === forbiddenActionTypes[i]) {
                return false;
            }
        }
    }
    return true;
}
function isElementAllowed(element, forbiddenElementTypes) {
    if (!hostConfig.supportsInteractivity && element.isInteractive) {
        return false;
    }
    if (forbiddenElementTypes) {
        for (var i = 0; i < forbiddenElementTypes.length; i++) {
            if (element.getJsonTypeName() === forbiddenElementTypes[i]) {
                return false;
            }
        }
    }
    return true;
}
var CardElement = (function () {
    function CardElement() {
        this._parent = null;
        this.horizontalAlignment = "left";
        this.separation = "default";
    }
    CardElement.prototype.internalGetNonZeroPadding = function (element, padding) {
        if (padding.top == 0) {
            padding.top = element.padding.top;
        }
        if (padding.right == 0) {
            padding.right = element.padding.right;
        }
        if (padding.bottom == 0) {
            padding.bottom = element.padding.bottom;
        }
        if (padding.left == 0) {
            padding.left = element.padding.left;
        }
        if (element.parent) {
            this.internalGetNonZeroPadding(element.parent, padding);
        }
    };
    CardElement.prototype.showBottomSpacer = function (requestingElement) {
        if (this.parent) {
            this.parent.showBottomSpacer(this);
        }
    };
    CardElement.prototype.hideBottomSpacer = function (requestingElement) {
        if (this.parent) {
            this.parent.hideBottomSpacer(this);
        }
    };
    CardElement.prototype.setParent = function (value) {
        this._parent = value;
    };
    Object.defineProperty(CardElement.prototype, "useDefaultSizing", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CardElement.prototype, "padding", {
        get: function () {
            return { top: 0, right: 0, bottom: 0, left: 0 };
        },
        enumerable: true,
        configurable: true
    });
    CardElement.prototype.getNonZeroPadding = function () {
        var padding = { top: 0, right: 0, bottom: 0, left: 0 };
        this.internalGetNonZeroPadding(this, padding);
        return padding;
    };
    CardElement.prototype.getForbiddenElementTypes = function () {
        return null;
    };
    CardElement.prototype.getForbiddenActionTypes = function () {
        return null;
    };
    CardElement.prototype.parse = function (json) {
        this.speak = json["speak"];
        this.horizontalAlignment = Utils.getValueOrDefault(json["horizontalAlignment"], "left");
        this.separation = Utils.getValueOrDefault(json["separation"], "default");
    };
    CardElement.prototype.validate = function () {
        return [];
    };
    CardElement.prototype.render = function () {
        var renderedElement = this.internalRender();
        if (renderedElement != null) {
            renderedElement.style.boxSizing = "border-box";
        }
        return renderedElement;
    };
    CardElement.prototype.isLastItem = function (item) {
        return this.parent ? this.parent.isLastItem(item) : true;
    };
    CardElement.prototype.getRootElement = function () {
        var rootElement = this;
        while (rootElement.parent) {
            rootElement = rootElement.parent;
        }
        return rootElement;
    };
    CardElement.prototype.getAllInputs = function () {
        return [];
    };
    Object.defineProperty(CardElement.prototype, "isInteractive", {
        get: function () {
            return false;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CardElement.prototype, "isStandalone", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CardElement.prototype, "parent", {
        get: function () {
            return this._parent;
        },
        enumerable: true,
        configurable: true
    });
    return CardElement;
}());
exports.CardElement = CardElement;
var TextBlock = (function (_super) {
    __extends(TextBlock, _super);
    function TextBlock() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.size = "normal";
        _this.weight = "normal";
        _this.isSubtle = false;
        _this.wrap = false;
        return _this;
    }
    TextBlock.prototype.internalRender = function () {
        if (!Utils.isNullOrEmpty(this.text)) {
            var element = document.createElement("div");
            if (hostConfig.fontFamily) {
                element.style.fontFamily = hostConfig.fontFamily;
            }
            switch (this.horizontalAlignment) {
                case "center":
                    element.style.textAlign = "center";
                    break;
                case "right":
                    element.style.textAlign = "right";
                    break;
                default:
                    element.style.textAlign = "left";
                    break;
            }
            var cssStyle = "text ";
            var fontSize;
            switch (this.size) {
                case "small":
                    fontSize = hostConfig.fontSizes.small;
                    break;
                case "medium":
                    fontSize = hostConfig.fontSizes.medium;
                    break;
                case "large":
                    fontSize = hostConfig.fontSizes.large;
                    break;
                case "extraLarge":
                    fontSize = hostConfig.fontSizes.extraLarge;
                    break;
                default:
                    fontSize = hostConfig.fontSizes.normal;
                    break;
            }
            // Looks like 1.33 is the magic number to compute line-height
            // from font size.
            var computedLineHeight = fontSize * 1.33;
            element.style.fontSize = fontSize + "px";
            element.style.lineHeight = computedLineHeight + "px";
            var actualTextColor = this.color ? this.color : hostConfig.textBlock.color;
            var colorDefinition;
            switch (actualTextColor) {
                case "dark":
                    colorDefinition = hostConfig.colors.dark;
                    break;
                case "light":
                    colorDefinition = hostConfig.colors.light;
                    break;
                case "accent":
                    colorDefinition = hostConfig.colors.accent;
                    break;
                case "good":
                    colorDefinition = hostConfig.colors.good;
                    break;
                case "warning":
                    colorDefinition = hostConfig.colors.warning;
                    break;
                case "attention":
                    colorDefinition = hostConfig.colors.attention;
                    break;
                default:
                    colorDefinition = hostConfig.colors.dark;
                    break;
            }
            element.style.color = Utils.stringToCssColor(this.isSubtle ? colorDefinition.subtle : colorDefinition.normal);
            var fontWeight;
            switch (this.weight) {
                case "lighter":
                    fontWeight = hostConfig.fontWeights.lighter;
                    break;
                case "bolder":
                    fontWeight = hostConfig.fontWeights.bolder;
                    break;
                default:
                    fontWeight = hostConfig.fontWeights.normal;
                    break;
            }
            element.style.fontWeight = fontWeight.toString();
            var formattedText = TextFormatters.formatText(this.text);
            element.innerHTML = Utils.processMarkdown(formattedText);
            if (element.firstElementChild instanceof HTMLElement) {
                var firstElementChild = element.firstElementChild;
                firstElementChild.style.marginTop = "0px";
                firstElementChild.style.width = "100%";
                if (!this.wrap) {
                    firstElementChild.style.overflow = "hidden";
                    firstElementChild.style.textOverflow = "ellipsis";
                }
            }
            if (element.lastElementChild instanceof HTMLElement) {
                element.lastElementChild.style.marginBottom = "0px";
            }
            var anchors = element.getElementsByTagName("a");
            for (var i = 0; i < anchors.length; i++) {
                anchors[i].target = "_blank";
            }
            if (this.wrap) {
                element.style.wordWrap = "break-word";
                if (this.maxLines > 0) {
                    element.style.maxHeight = (computedLineHeight * this.maxLines) + "px";
                    element.style.overflow = "hidden";
                }
            }
            else {
                element.style.whiteSpace = "nowrap";
            }
            return element;
        }
        else {
            return null;
        }
    };
    TextBlock.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.text = json["text"];
        this.size = Utils.getValueOrDefault(json["size"], "normal");
        this.weight = Utils.getValueOrDefault(json["weight"], "normal");
        this.color = Utils.getValueOrDefault(json["color"], hostConfig.textBlock.color);
        this.isSubtle = json["isSubtle"];
        this.wrap = json["wrap"] === undefined ? false : json["wrap"];
        this.maxLines = json["maxLines"];
    };
    TextBlock.prototype.getJsonTypeName = function () {
        return "TextBlock";
    };
    TextBlock.prototype.getDefaultSeparationDefinition = function () {
        switch (this.size) {
            case "small":
                return hostConfig.textBlock.separations.small;
            case "medium":
                return hostConfig.textBlock.separations.medium;
            case "large":
                return hostConfig.textBlock.separations.large;
            case "extraLarge":
                return hostConfig.textBlock.separations.extraLarge;
            default:
                return hostConfig.textBlock.separations.normal;
        }
    };
    TextBlock.prototype.renderSpeech = function () {
        if (this.speak != null)
            return this.speak + '\n';
        if (this.text)
            return '<s>' + this.text + '</s>\n';
        return null;
    };
    return TextBlock;
}(CardElement));
exports.TextBlock = TextBlock;
var Fact = (function () {
    function Fact() {
    }
    Fact.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak + '\n';
        }
        return '<s>' + this.name + ' ' + this.value + '</s>\n';
    };
    return Fact;
}());
exports.Fact = Fact;
var FactSet = (function (_super) {
    __extends(FactSet, _super);
    function FactSet() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.facts = [];
        return _this;
    }
    Object.defineProperty(FactSet.prototype, "useDefaultSizing", {
        get: function () {
            return false;
        },
        enumerable: true,
        configurable: true
    });
    FactSet.prototype.internalRender = function () {
        var element = null;
        if (this.facts.length > 0) {
            element = document.createElement("table");
            element.style.borderWidth = "0px";
            element.style.borderSpacing = "0px";
            element.style.borderStyle = "none";
            element.style.borderCollapse = "collapse";
            element.style.display = "block";
            element.style.overflow = "hidden";
            for (var i = 0; i < this.facts.length; i++) {
                var trElement = document.createElement("tr");
                if (i > 0) {
                    trElement.style.marginTop = hostConfig.factSet.spacing + "px";
                }
                var tdElement = document.createElement("td");
                tdElement.style.padding = "0";
                if (hostConfig.factSet.title.maxWidth) {
                    tdElement.style.maxWidth = hostConfig.factSet.title.maxWidth + "px";
                }
                tdElement.style.verticalAlign = "top";
                var textBlock = new TextBlock();
                textBlock.text = this.facts[i].name;
                textBlock.size = hostConfig.factSet.title.size;
                textBlock.color = hostConfig.factSet.title.color;
                textBlock.isSubtle = hostConfig.factSet.title.isSubtle;
                textBlock.weight = hostConfig.factSet.title.weight;
                textBlock.wrap = hostConfig.factSet.title.wrap;
                textBlock.separation = "none";
                Utils.appendChild(tdElement, textBlock.render());
                Utils.appendChild(trElement, tdElement);
                tdElement = document.createElement("td");
                tdElement.style.padding = "0px 0px 0px 10px";
                tdElement.style.verticalAlign = "top";
                textBlock = new TextBlock();
                textBlock.text = this.facts[i].value;
                textBlock.size = hostConfig.factSet.value.size;
                textBlock.color = hostConfig.factSet.value.color;
                textBlock.isSubtle = hostConfig.factSet.value.isSubtle;
                textBlock.weight = hostConfig.factSet.value.weight;
                textBlock.wrap = hostConfig.factSet.value.wrap;
                textBlock.separation = "none";
                Utils.appendChild(tdElement, textBlock.render());
                Utils.appendChild(trElement, tdElement);
                Utils.appendChild(element, trElement);
            }
        }
        return element;
    };
    FactSet.prototype.getJsonTypeName = function () {
        return "FactSet";
    };
    FactSet.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.factSet.separation;
    };
    FactSet.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        if (json["facts"] != null) {
            var jsonFacts = json["facts"];
            for (var i = 0; i < jsonFacts.length; i++) {
                var fact = new Fact();
                fact.name = jsonFacts[i]["title"];
                fact.value = jsonFacts[i]["value"];
                fact.speak = jsonFacts[i]["speak"];
                this.facts.push(fact);
            }
        }
    };
    FactSet.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak + '\n';
        }
        // render each fact
        var speak = null;
        if (this.facts.length > 0) {
            speak = '';
            for (var i = 0; i < this.facts.length; i++) {
                var speech = this.facts[i].renderSpeech();
                if (speech) {
                    speak += speech;
                }
            }
        }
        return '<p>' + speak + '\n</p>\n';
    };
    return FactSet;
}(CardElement));
exports.FactSet = FactSet;
var Image = (function (_super) {
    __extends(Image, _super);
    function Image() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.style = "normal";
        _this.size = "medium";
        return _this;
    }
    Object.defineProperty(Image.prototype, "useDefaultSizing", {
        get: function () {
            return false;
        },
        enumerable: true,
        configurable: true
    });
    Image.prototype.internalRender = function () {
        var _this = this;
        var element = null;
        if (!Utils.isNullOrEmpty(this.url)) {
            element = document.createElement("div");
            element.classList.add("ac-image");
            element.style.display = "flex";
            element.style.alignItems = "flex-start";
            if (this.selectAction != null) {
                element.classList.add("ac-selectable");
            }
            element.onclick = function (e) {
                if (_this.selectAction) {
                    _this.selectAction.execute();
                    e.cancelBubble = true;
                }
            };
            switch (this.horizontalAlignment) {
                case "center":
                    element.style.justifyContent = "center";
                    break;
                case "right":
                    element.style.justifyContent = "flex-end";
                    break;
                default:
                    element.style.justifyContent = "flex-start";
                    break;
            }
            var imageElement = document.createElement("img");
            imageElement.style.maxHeight = "100%";
            switch (this.size) {
                case "stretch":
                    imageElement.style.width = "100%";
                    break;
                case "auto":
                    imageElement.style.maxWidth = "100%";
                    imageElement.style.maxHeight = "500px";
                    break;
                case "small":
                    imageElement.style.maxWidth = hostConfig.imageSizes.small + "px";
                    break;
                case "large":
                    imageElement.style.maxWidth = hostConfig.imageSizes.large + "px";
                    break;
                case "medium":
                    imageElement.style.maxWidth = hostConfig.imageSizes.medium + "px";
                    break;
            }
            if (this.style == "person") {
                imageElement.style.borderRadius = "50%";
                imageElement.style.backgroundPosition = "50% 50%";
                imageElement.style.backgroundRepeat = "no-repeat";
            }
            imageElement.src = this.url;
            element.appendChild(imageElement);
        }
        return element;
    };
    Image.prototype.getJsonTypeName = function () {
        return "Image";
    };
    Image.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.image.separation;
    };
    Image.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.url = json["url"];
        this.style = Utils.getValueOrDefault(json["style"], "normal");
        this.size = Utils.getValueOrDefault(json["size"], "medium");
        var selectActionJson = json["selectAction"];
        if (selectActionJson != undefined) {
            this.selectAction = Action.createAction(selectActionJson);
            invokeSetParent(this.selectAction, this);
        }
    };
    Image.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak + '\n';
        }
        return null;
    };
    return Image;
}(CardElement));
exports.Image = Image;
var ImageSet = (function (_super) {
    __extends(ImageSet, _super);
    function ImageSet() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._images = [];
        _this.imageSize = "medium";
        return _this;
    }
    ImageSet.prototype.internalRender = function () {
        var element = null;
        if (this._images.length > 0) {
            element = document.createElement("div");
            element.style.display = "flex";
            element.style.flexWrap = "wrap";
            for (var i = 0; i < this._images.length; i++) {
                var renderedImage = this._images[i].render();
                renderedImage.style.display = "inline-flex";
                renderedImage.style.margin = "0px";
                renderedImage.style.marginRight = "10px";
                renderedImage.style.height = "100px";
                Utils.appendChild(element, renderedImage);
            }
        }
        return element;
    };
    ImageSet.prototype.getJsonTypeName = function () {
        return "ImageSet";
    };
    ImageSet.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.imageSet.separation;
    };
    ImageSet.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.imageSize = Utils.getValueOrDefault(json["imageSize"], "medium");
        if (json["images"] != null) {
            var jsonImages = json["images"];
            for (var i = 0; i < jsonImages.length; i++) {
                var image = new Image();
                image.parse(jsonImages[i]);
                image.size = this.imageSize;
                this.addImage(image);
            }
        }
    };
    ImageSet.prototype.addImage = function (image) {
        if (!image.parent) {
            this._images.push(image);
            invokeSetParent(image, this);
        }
        else {
            throw new Error("This image already belongs to another ImageSet");
        }
    };
    ImageSet.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak;
        }
        var speak = null;
        if (this._images.length > 0) {
            speak = '';
            for (var i = 0; i < this._images.length; i++) {
                speak += this._images[i].renderSpeech();
            }
        }
        return speak;
    };
    return ImageSet;
}(CardElement));
exports.ImageSet = ImageSet;
var Input = (function (_super) {
    __extends(Input, _super);
    function Input() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Input.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.input.separation;
    };
    Input.prototype.validate = function () {
        if (!this.id) {
            return [{ error: Enums.ValidationError.PropertyCantBeNull, message: "All inputs must have a unique Id" }];
        }
        else {
            return [];
        }
    };
    Input.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.id = json["id"];
        this.defaultValue = json["value"];
    };
    Input.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak;
        }
        if (this.title) {
            return '<s>' + this.title + '</s>\n';
        }
        return null;
    };
    Input.prototype.getAllInputs = function () {
        return [this];
    };
    Object.defineProperty(Input.prototype, "isInteractive", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    return Input;
}(CardElement));
exports.Input = Input;
var TextInput = (function (_super) {
    __extends(TextInput, _super);
    function TextInput() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TextInput.prototype.internalRender = function () {
        if (this.isMultiline) {
            this._textareaElement = document.createElement("textarea");
            this._textareaElement.className = "ac-input ac-textInput ac-multiline";
            this._textareaElement.style.width = "100%";
            if (!Utils.isNullOrEmpty(this.placeholder)) {
                this._textareaElement.placeholder = this.placeholder;
            }
            if (!Utils.isNullOrEmpty(this.defaultValue)) {
                this._textareaElement.value = this.defaultValue;
            }
            if (this.maxLength > 0) {
                this._textareaElement.maxLength = this.maxLength;
            }
            return this._textareaElement;
        }
        else {
            this._inputElement = document.createElement("input");
            this._inputElement.type = "text";
            this._inputElement.className = "ac-input ac-textInput";
            this._inputElement.style.width = "100%";
            if (!Utils.isNullOrEmpty(this.placeholder)) {
                this._inputElement.placeholder = this.placeholder;
            }
            if (!Utils.isNullOrEmpty(this.defaultValue)) {
                this._inputElement.value = this.defaultValue;
            }
            if (this.maxLength > 0) {
                this._inputElement.maxLength = this.maxLength;
            }
            return this._inputElement;
        }
    };
    TextInput.prototype.getJsonTypeName = function () {
        return "Input.Text";
    };
    TextInput.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.maxLength = json["maxLength"];
        this.isMultiline = json["isMultiline"];
        this.placeholder = json["placeholder"];
    };
    Object.defineProperty(TextInput.prototype, "value", {
        get: function () {
            if (this.isMultiline) {
                return this._textareaElement ? this._textareaElement.value : null;
            }
            else {
                return this._inputElement ? this._inputElement.value : null;
            }
        },
        enumerable: true,
        configurable: true
    });
    return TextInput;
}(Input));
exports.TextInput = TextInput;
var ToggleInput = (function (_super) {
    __extends(ToggleInput, _super);
    function ToggleInput() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ToggleInput.prototype.internalRender = function () {
        var element = document.createElement("div");
        element.className = "ac-input";
        element.style.width = "100%";
        this._checkboxInputElement = document.createElement("input");
        this._checkboxInputElement.type = "checkbox";
        this._checkboxInputElement.style.display = "inline-block";
        this._checkboxInputElement.style.verticalAlign = "middle";
        this._checkboxInputElement.style.margin = "0";
        if (this.defaultValue == this.valueOn) {
            this._checkboxInputElement.checked = true;
        }
        var label = new TextBlock();
        label.text = this.title;
        var labelElement = label.render();
        labelElement.style.display = "inline-block";
        labelElement.style.marginLeft = "6px";
        labelElement.style.verticalAlign = "middle";
        var compoundInput = document.createElement("div");
        Utils.appendChild(element, this._checkboxInputElement);
        Utils.appendChild(element, labelElement);
        return element;
    };
    ToggleInput.prototype.getJsonTypeName = function () {
        return "Input.Toggle";
    };
    ToggleInput.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.title = json["title"];
        this.valueOn = json["valueOn"];
        this.valueOff = json["valueOff"];
    };
    Object.defineProperty(ToggleInput.prototype, "value", {
        get: function () {
            if (this._checkboxInputElement) {
                return this._checkboxInputElement.checked ? this.valueOn : this.valueOff;
            }
            else {
                return null;
            }
        },
        enumerable: true,
        configurable: true
    });
    return ToggleInput;
}(Input));
exports.ToggleInput = ToggleInput;
var Choice = (function () {
    function Choice() {
    }
    return Choice;
}());
exports.Choice = Choice;
var ChoiceSetInput = (function (_super) {
    __extends(ChoiceSetInput, _super);
    function ChoiceSetInput() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.choices = [];
        return _this;
    }
    ChoiceSetInput.prototype.internalRender = function () {
        if (!this.isMultiSelect) {
            if (this.isCompact) {
                // Render as a combo box
                this._selectElement = document.createElement("select");
                this._selectElement.className = "ac-input ac-multichoiceInput";
                this._selectElement.style.width = "100%";
                var option = document.createElement("option");
                option.selected = true;
                option.disabled = true;
                option.hidden = true;
                if (this.placeholder) {
                    option.text = this.placeholder;
                }
                Utils.appendChild(this._selectElement, option);
                for (var i = 0; i < this.choices.length; i++) {
                    var option = document.createElement("option");
                    option.value = this.choices[i].value;
                    option.text = this.choices[i].title;
                    if (this.choices[i].value == this.defaultValue) {
                        option.selected = true;
                    }
                    Utils.appendChild(this._selectElement, option);
                }
                return this._selectElement;
            }
            else {
                // Render as a series of radio buttons
                var element = document.createElement("div");
                element.className = "ac-input";
                element.style.width = "100%";
                this._toggleInputs = [];
                for (var i = 0; i < this.choices.length; i++) {
                    var radioInput = document.createElement("input");
                    radioInput.type = "radio";
                    radioInput.style.margin = "0";
                    radioInput.style.display = "inline-block";
                    radioInput.style.verticalAlign = "middle";
                    radioInput.name = this.id;
                    radioInput.value = this.choices[i].value;
                    if (this.choices[i].value == this.defaultValue) {
                        radioInput.checked = true;
                    }
                    this._toggleInputs.push(radioInput);
                    var label = new TextBlock();
                    label.text = this.choices[i].title;
                    var labelElement = label.render();
                    labelElement.style.display = "inline-block";
                    labelElement.style.marginLeft = "6px";
                    labelElement.style.verticalAlign = "middle";
                    var compoundInput = document.createElement("div");
                    Utils.appendChild(compoundInput, radioInput);
                    Utils.appendChild(compoundInput, labelElement);
                    Utils.appendChild(element, compoundInput);
                }
                return element;
            }
        }
        else {
            // Render as a list of toggle inputs
            var defaultValues = this.defaultValue ? this.defaultValue.split(",") : null;
            var element = document.createElement("div");
            element.className = "ac-input";
            element.style.width = "100%";
            this._toggleInputs = [];
            for (var i = 0; i < this.choices.length; i++) {
                var checkboxInput = document.createElement("input");
                checkboxInput.type = "checkbox";
                checkboxInput.style.margin = "0";
                checkboxInput.style.display = "inline-block";
                checkboxInput.style.verticalAlign = "middle";
                checkboxInput.value = this.choices[i].value;
                if (defaultValues) {
                    if (defaultValues.indexOf(this.choices[i].value) >= 0) {
                        checkboxInput.checked = true;
                    }
                }
                this._toggleInputs.push(checkboxInput);
                var label = new TextBlock();
                label.text = this.choices[i].title;
                var labelElement = label.render();
                labelElement.style.display = "inline-block";
                labelElement.style.marginLeft = "6px";
                labelElement.style.verticalAlign = "middle";
                var compoundInput = document.createElement("div");
                Utils.appendChild(compoundInput, checkboxInput);
                Utils.appendChild(compoundInput, labelElement);
                Utils.appendChild(element, compoundInput);
            }
            return element;
        }
    };
    ChoiceSetInput.prototype.getJsonTypeName = function () {
        return "Input.ChoiceSet";
    };
    ChoiceSetInput.prototype.validate = function () {
        var result = [];
        if (this.choices.length == 0) {
            result = [{ error: Enums.ValidationError.CollectionCantBeEmpty, message: "An Input.ChoiceSet must have at least one choice defined." }];
        }
        for (var i = 0; i < this.choices.length; i++) {
            if (!this.choices[i].title || !this.choices[i].value) {
                result = result.concat([{ error: Enums.ValidationError.PropertyCantBeNull, message: "All choices in an Input.ChoiceSet must have their title and value properties set." }]);
                break;
            }
        }
        return result;
    };
    ChoiceSetInput.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.isCompact = !(json["style"] === "expanded");
        this.isMultiSelect = json["isMultiSelect"];
        this.placeholder = json["placeholder"];
        if (json["choices"] != undefined) {
            var choiceArray = json["choices"];
            for (var i = 0; i < choiceArray.length; i++) {
                var choice = new Choice();
                choice.title = choiceArray[i]["title"];
                choice.value = choiceArray[i]["value"];
                this.choices.push(choice);
            }
        }
    };
    Object.defineProperty(ChoiceSetInput.prototype, "value", {
        get: function () {
            if (!this.isMultiSelect) {
                if (this.isCompact) {
                    return this._selectElement ? this._selectElement.value : null;
                }
                else {
                    if (!this._toggleInputs || this._toggleInputs.length == 0) {
                        return null;
                    }
                    for (var i = 0; i < this._toggleInputs.length; i++) {
                        if (this._toggleInputs[i].checked) {
                            return this._toggleInputs[i].value;
                        }
                    }
                    return null;
                }
            }
            else {
                if (!this._toggleInputs || this._toggleInputs.length == 0) {
                    return null;
                }
                var result = "";
                for (var i = 0; i < this._toggleInputs.length; i++) {
                    if (this._toggleInputs[i].checked) {
                        if (result != "") {
                            result += ";";
                        }
                        result += this._toggleInputs[i].value;
                    }
                }
                return result == "" ? null : result;
            }
        },
        enumerable: true,
        configurable: true
    });
    return ChoiceSetInput;
}(Input));
exports.ChoiceSetInput = ChoiceSetInput;
var NumberInput = (function (_super) {
    __extends(NumberInput, _super);
    function NumberInput() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    NumberInput.prototype.internalRender = function () {
        this._numberInputElement = document.createElement("input");
        this._numberInputElement.type = "number";
        this._numberInputElement.className = "ac-input ac-numberInput";
        this._numberInputElement.min = this.min;
        this._numberInputElement.max = this.max;
        this._numberInputElement.style.width = "100%";
        if (!Utils.isNullOrEmpty(this.defaultValue)) {
            this._numberInputElement.value = this.defaultValue;
        }
        return this._numberInputElement;
    };
    NumberInput.prototype.getJsonTypeName = function () {
        return "Input.Number";
    };
    NumberInput.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.min = json["min"];
        this.max = json["max"];
    };
    Object.defineProperty(NumberInput.prototype, "value", {
        get: function () {
            return this._numberInputElement ? this._numberInputElement.value : null;
        },
        enumerable: true,
        configurable: true
    });
    return NumberInput;
}(Input));
exports.NumberInput = NumberInput;
var DateInput = (function (_super) {
    __extends(DateInput, _super);
    function DateInput() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    DateInput.prototype.internalRender = function () {
        this._dateInputElement = document.createElement("input");
        this._dateInputElement.type = "date";
        this._dateInputElement.className = "ac-input ac-dateInput";
        this._dateInputElement.style.width = "100%";
        if (!Utils.isNullOrEmpty(this.defaultValue)) {
            this._dateInputElement.value = this.defaultValue;
        }
        return this._dateInputElement;
    };
    DateInput.prototype.getJsonTypeName = function () {
        return "Input.Date";
    };
    Object.defineProperty(DateInput.prototype, "value", {
        get: function () {
            return this._dateInputElement ? this._dateInputElement.value : null;
        },
        enumerable: true,
        configurable: true
    });
    return DateInput;
}(Input));
exports.DateInput = DateInput;
var TimeInput = (function (_super) {
    __extends(TimeInput, _super);
    function TimeInput() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TimeInput.prototype.internalRender = function () {
        this._timeInputElement = document.createElement("input");
        this._timeInputElement.type = "time";
        this._timeInputElement.className = "ac-input ac-timeInput";
        this._timeInputElement.style.width = "100%";
        if (!Utils.isNullOrEmpty(this.defaultValue)) {
            this._timeInputElement.value = this.defaultValue;
        }
        return this._timeInputElement;
    };
    TimeInput.prototype.getJsonTypeName = function () {
        return "Input.Time";
    };
    Object.defineProperty(TimeInput.prototype, "value", {
        get: function () {
            return this._timeInputElement ? this._timeInputElement.value : null;
        },
        enumerable: true,
        configurable: true
    });
    return TimeInput;
}(Input));
exports.TimeInput = TimeInput;
var ActionButtonState;
(function (ActionButtonState) {
    ActionButtonState[ActionButtonState["Normal"] = 0] = "Normal";
    ActionButtonState[ActionButtonState["Expanded"] = 1] = "Expanded";
    ActionButtonState[ActionButtonState["Subdued"] = 2] = "Subdued";
})(ActionButtonState || (ActionButtonState = {}));
var ActionButton = (function () {
    function ActionButton(action, style) {
        var _this = this;
        this._style = "button";
        this._element = null;
        this._state = ActionButtonState.Normal;
        this.onClick = null;
        this._action = action;
        this._style = style;
        this._element = document.createElement("button");
        this._element.style.overflow = "hidden";
        this._element.style.whiteSpace = "nowrap";
        this._element.style.textOverflow = "ellipsis";
        this._element.onclick = function (e) { _this.click(); };
        this.updateCssStyle();
    }
    ActionButton.prototype.click = function () {
        if (this.onClick != null) {
            this.onClick(this);
        }
    };
    ActionButton.prototype.updateCssStyle = function () {
        this._element.className = this._style == "link" ? "ac-linkButton" : "ac-pushButton";
        if (this._action instanceof ShowCardAction) {
            this._element.classList.add("expandable");
        }
        switch (this._state) {
            case ActionButtonState.Expanded:
                this._element.classList.add("expanded");
                break;
            case ActionButtonState.Subdued:
                this._element.classList.add("subdued");
                break;
        }
    };
    Object.defineProperty(ActionButton.prototype, "action", {
        get: function () {
            return this._action;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ActionButton.prototype, "text", {
        get: function () {
            return this._text;
        },
        set: function (value) {
            this._text = value;
            this._element.innerText = this._text;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ActionButton.prototype, "element", {
        get: function () {
            return this._element;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ActionButton.prototype, "state", {
        get: function () {
            return this._state;
        },
        set: function (value) {
            this._state = value;
            this.updateCssStyle();
        },
        enumerable: true,
        configurable: true
    });
    return ActionButton;
}());
var Action = (function () {
    function Action() {
        this._parent = null;
    }
    Action.createAction = function (json) {
        var actionType = json["type"];
        var result = AdaptiveCard.actionTypeRegistry.createInstance(actionType);
        if (result) {
            result.parse(json);
        }
        else {
            raiseParseError({
                error: Enums.ValidationError.UnknownActionType,
                message: "Unknown action type: " + actionType
            });
        }
        return result;
    };
    Action.prototype.setParent = function (value) {
        this._parent = value;
    };
    Action.prototype.validate = function () {
        return [];
    };
    Action.prototype.prepare = function (inputs) {
        // Do nothing in base implementation
    };
    ;
    Action.prototype.parse = function (json) {
        this.title = json["title"];
    };
    Action.prototype.getAllInputs = function () {
        return [];
    };
    Object.defineProperty(Action.prototype, "parent", {
        get: function () {
            return this._parent;
        },
        enumerable: true,
        configurable: true
    });
    return Action;
}());
exports.Action = Action;
var ExternalAction = (function (_super) {
    __extends(ExternalAction, _super);
    function ExternalAction() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ExternalAction.prototype.execute = function () {
        raiseExecuteActionEvent(this);
    };
    return ExternalAction;
}(Action));
exports.ExternalAction = ExternalAction;
var SubmitAction = (function (_super) {
    __extends(SubmitAction, _super);
    function SubmitAction() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isPrepared = false;
        return _this;
    }
    SubmitAction.prototype.getJsonTypeName = function () {
        return "Action.Submit";
    };
    SubmitAction.prototype.prepare = function (inputs) {
        if (this._originalData) {
            this._processedData = JSON.parse(JSON.stringify(this._originalData));
        }
        else {
            this._processedData = {};
        }
        for (var i = 0; i < inputs.length; i++) {
            var inputValue = inputs[i].value;
            if (inputValue != null) {
                this._processedData[inputs[i].id] = inputs[i].value;
            }
        }
        this._isPrepared = true;
    };
    SubmitAction.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.data = json["data"];
    };
    Object.defineProperty(SubmitAction.prototype, "data", {
        get: function () {
            return this._isPrepared ? this._processedData : this._originalData;
        },
        set: function (value) {
            this._originalData = value;
            this._isPrepared = false;
        },
        enumerable: true,
        configurable: true
    });
    return SubmitAction;
}(ExternalAction));
exports.SubmitAction = SubmitAction;
var OpenUrlAction = (function (_super) {
    __extends(OpenUrlAction, _super);
    function OpenUrlAction() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    OpenUrlAction.prototype.getJsonTypeName = function () {
        return "Action.OpenUrl";
    };
    OpenUrlAction.prototype.validate = function () {
        if (!this.url) {
            return [{ error: Enums.ValidationError.PropertyCantBeNull, message: "An Action.OpenUrl must have its url property set." }];
        }
        else {
            return [];
        }
    };
    OpenUrlAction.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.url = json["url"];
    };
    return OpenUrlAction;
}(ExternalAction));
exports.OpenUrlAction = OpenUrlAction;
var HttpHeader = (function () {
    function HttpHeader() {
        this._value = new Utils.StringWithSubstitutions();
    }
    HttpHeader.prototype.prepare = function (inputs) {
        this._value.substituteInputValues(inputs);
    };
    Object.defineProperty(HttpHeader.prototype, "value", {
        get: function () {
            return this._value.get();
        },
        set: function (newValue) {
            this._value.set(newValue);
        },
        enumerable: true,
        configurable: true
    });
    return HttpHeader;
}());
exports.HttpHeader = HttpHeader;
var HttpAction = (function (_super) {
    __extends(HttpAction, _super);
    function HttpAction() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._url = new Utils.StringWithSubstitutions();
        _this._body = new Utils.StringWithSubstitutions();
        _this._headers = [];
        return _this;
    }
    HttpAction.prototype.getJsonTypeName = function () {
        return "Action.Http";
    };
    HttpAction.prototype.validate = function () {
        var result = [];
        if (!this.url) {
            result = [{ error: Enums.ValidationError.PropertyCantBeNull, message: "An Action.Http must have its url property set." }];
        }
        if (this.headers.length > 0) {
            for (var i = 0; i < this.headers.length; i++) {
                if (!this.headers[i].name || !this.headers[i].value) {
                    result = result.concat([{ error: Enums.ValidationError.PropertyCantBeNull, message: "All headers of an Action.Http must have their name and value properties set." }]);
                    break;
                }
            }
        }
        return result;
    };
    HttpAction.prototype.prepare = function (inputs) {
        this._url.substituteInputValues(inputs);
        this._body.substituteInputValues(inputs);
        for (var i = 0; i < this._headers.length; i++) {
            this._headers[i].prepare(inputs);
        }
    };
    ;
    HttpAction.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.url = json["url"];
        this.method = json["method"];
        this.body = json["body"];
        if (json["headers"] != null) {
            var jsonHeaders = json["headers"];
            for (var i = 0; i < jsonHeaders.length; i++) {
                var httpHeader = new HttpHeader();
                httpHeader.name = jsonHeaders[i]["name"];
                httpHeader.value = jsonHeaders[i]["value"];
                this.headers.push(httpHeader);
            }
        }
    };
    Object.defineProperty(HttpAction.prototype, "url", {
        get: function () {
            return this._url.get();
        },
        set: function (value) {
            this._url.set(value);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpAction.prototype, "body", {
        get: function () {
            return this._body.get();
        },
        set: function (value) {
            this._body.set(value);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpAction.prototype, "headers", {
        get: function () {
            return this._headers;
        },
        enumerable: true,
        configurable: true
    });
    return HttpAction;
}(ExternalAction));
exports.HttpAction = HttpAction;
var ShowCardAction = (function (_super) {
    __extends(ShowCardAction, _super);
    function ShowCardAction() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.card = new InlineAdaptiveCard();
        return _this;
    }
    ShowCardAction.prototype.setParent = function (value) {
        _super.prototype.setParent.call(this, value);
        invokeSetParent(this.card, value);
    };
    ShowCardAction.prototype.execute = function () {
        raiseShowPopupCardEvent(this);
    };
    ShowCardAction.prototype.getJsonTypeName = function () {
        return "Action.ShowCard";
    };
    ShowCardAction.prototype.validate = function () {
        return this.card.validate();
    };
    ShowCardAction.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.card.parse(json["card"]);
    };
    ShowCardAction.prototype.getAllInputs = function () {
        return this.card.getAllInputs();
    };
    return ShowCardAction;
}(Action));
exports.ShowCardAction = ShowCardAction;
var ActionCollection = (function () {
    function ActionCollection(owner) {
        this._actionButtons = [];
        this._expandedAction = null;
        this._renderedActionCount = 0;
        this.items = [];
        this.actionStyle = "button";
        this.onHideActionCardPane = null;
        this.onShowActionCardPane = null;
        this._owner = owner;
    }
    ActionCollection.prototype.hideActionCardPane = function () {
        this._actionCardContainer.innerHTML = '';
        this._actionCardContainer.style.padding = "0px";
        this._actionCardContainer.style.marginTop = "0px";
        if (this.onHideActionCardPane) {
            this.onHideActionCardPane();
        }
        if (this._expandedAction) {
            raiseInlineCardExpandedEvent(this._expandedAction, false);
        }
        this._expandedAction = null;
    };
    ActionCollection.prototype.showActionCardPane = function (action) {
        if (this.onShowActionCardPane) {
            this.onShowActionCardPane(action);
        }
        var renderedCard = action.card.render();
        this._actionCardContainer.innerHTML = '';
        this._actionCardContainer.style.marginTop = this._renderedActionCount > 0 ? hostConfig.actions.showCard.inlineTopMargin + "px" : "0px";
        if (hostConfig.actions.showCard.actionMode == "inlineEdgeToEdge") {
            var padding = this._owner.getNonZeroPadding();
            this._actionCardContainer.style.paddingLeft = padding.left + "px";
            this._actionCardContainer.style.paddingRight = padding.right + "px";
            this._actionCardContainer.style.marginLeft = "-" + padding.left + "px";
            this._actionCardContainer.style.marginRight = "-" + padding.right + "px";
            renderedCard.style.paddingLeft = "0px";
            renderedCard.style.paddingRight = "0px";
        }
        Utils.appendChild(this._actionCardContainer, renderedCard);
        raiseInlineCardExpandedEvent(action, true);
        this._expandedAction = action;
    };
    ActionCollection.prototype.actionClicked = function (actionButton) {
        if (!(actionButton.action instanceof ShowCardAction)) {
            for (var i = 0; i < this._actionButtons.length; i++) {
                this._actionButtons[i].state = ActionButtonState.Normal;
            }
            this.hideActionCardPane();
            actionButton.action.execute();
        }
        else {
            if (hostConfig.actions.showCard.actionMode == "popup") {
                actionButton.action.execute();
            }
            else if (actionButton.action === this._expandedAction) {
                for (var i = 0; i < this._actionButtons.length; i++) {
                    this._actionButtons[i].state = ActionButtonState.Normal;
                }
                this.hideActionCardPane();
            }
            else {
                for (var i = 0; i < this._actionButtons.length; i++) {
                    if (this._actionButtons[i] !== actionButton) {
                        this._actionButtons[i].state = ActionButtonState.Subdued;
                    }
                }
                actionButton.state = ActionButtonState.Expanded;
                this.showActionCardPane(actionButton.action);
            }
        }
    };
    ActionCollection.prototype.validate = function () {
        var result = [];
        if (hostConfig.actions.maxActions && this.items.length > hostConfig.actions.maxActions) {
            result.push({
                error: Enums.ValidationError.TooManyActions,
                message: "A maximum of " + hostConfig.actions.maxActions + " actions are allowed."
            });
        }
        if (this.items.length > 0 && !hostConfig.supportsInteractivity) {
            result.push({
                error: Enums.ValidationError.InteractivityNotAllowed,
                message: "Interactivity is not allowed."
            });
        }
        for (var i = 0; i < this.items.length; i++) {
            if (!isActionAllowed(this.items[i], this._owner.getForbiddenActionTypes())) {
                result.push({
                    error: Enums.ValidationError.ActionTypeNotAllowed,
                    message: "Actions of type " + this.items[i].getJsonTypeName() + " are not allowe."
                });
            }
        }
        for (var i = 0; i < this.items.length; i++) {
            result = result.concat(this.items[i].validate());
        }
        return result;
    };
    ActionCollection.prototype.render = function () {
        var _this = this;
        if (!hostConfig.supportsInteractivity) {
            return null;
        }
        var element = document.createElement("div");
        this._actionCardContainer = document.createElement("div");
        this._actionCardContainer.style.backgroundColor = Utils.stringToCssColor(hostConfig.actions.showCard.backgroundColor);
        this._renderedActionCount = 0;
        var maxActions = hostConfig.actions.maxActions ? Math.min(hostConfig.actions.maxActions, this.items.length) : this.items.length;
        var forbiddenActionTypes = this._owner.getForbiddenActionTypes();
        if (hostConfig.actions.preExpandSingleShowCardAction && maxActions == 1 && this.items[0] instanceof ShowCardAction && isActionAllowed(this.items[i], forbiddenActionTypes)) {
            this.showActionCardPane(this.items[0]);
            this._renderedActionCount = 1;
        }
        else {
            var buttonStrip = document.createElement("div");
            buttonStrip.style.display = "flex";
            if (hostConfig.actions.actionsOrientation == "horizontal") {
                buttonStrip.style.flexDirection = "row";
                switch (hostConfig.actions.actionAlignment) {
                    case "center":
                        buttonStrip.style.justifyContent = "center";
                        break;
                    case "right":
                        buttonStrip.style.justifyContent = "flex-end";
                        break;
                    default:
                        buttonStrip.style.justifyContent = "flex-start";
                        break;
                }
            }
            else {
                buttonStrip.style.flexDirection = "column";
                switch (hostConfig.actions.actionAlignment) {
                    case "center":
                        buttonStrip.style.alignItems = "center";
                        break;
                    case "right":
                        buttonStrip.style.alignItems = "flex-end";
                        break;
                    case "stretch":
                        buttonStrip.style.alignItems = "stretch";
                        break;
                    default:
                        buttonStrip.style.alignItems = "flex-start";
                        break;
                }
            }
            for (var i = 0; i < maxActions; i++) {
                if (isActionAllowed(this.items[i], forbiddenActionTypes)) {
                    var actionButton = new ActionButton(this.items[i], this.actionStyle);
                    actionButton.element.style.overflow = "hidden";
                    actionButton.element.style.overflow = "table-cell";
                    actionButton.element.style.flex = hostConfig.actions.actionAlignment == "stretch" ? "0 1 100%" : "0 1 auto";
                    actionButton.text = this.items[i].title;
                    actionButton.onClick = function (ab) { _this.actionClicked(ab); };
                    this._actionButtons.push(actionButton);
                    buttonStrip.appendChild(actionButton.element);
                    if (i < this.items.length - 1 && hostConfig.actions.buttonSpacing > 0) {
                        var spacer = document.createElement("div");
                        if (hostConfig.actions.actionsOrientation == "horizontal") {
                            spacer.style.flex = "0 0 auto";
                            spacer.style.width = hostConfig.actions.buttonSpacing + "px";
                        }
                        else {
                            spacer.style.height = hostConfig.actions.buttonSpacing + "px";
                        }
                        Utils.appendChild(buttonStrip, spacer);
                    }
                    this._renderedActionCount++;
                }
            }
            var buttonStripContainer = document.createElement("div");
            buttonStripContainer.style.overflow = "hidden";
            buttonStripContainer.appendChild(buttonStrip);
            Utils.appendChild(element, buttonStripContainer);
        }
        Utils.appendChild(element, this._actionCardContainer);
        return this._renderedActionCount > 0 ? element : null;
    };
    ActionCollection.prototype.addAction = function (action) {
        if (!action.parent) {
            this.items.push(action);
            invokeSetParent(action, this._owner);
        }
        else {
            throw new Error("The action already belongs to another element.");
        }
    };
    ActionCollection.prototype.clear = function () {
        this.items = [];
    };
    ActionCollection.prototype.getAllInputs = function () {
        var result = [];
        for (var i = 0; i < this.items.length; i++) {
            var action = this.items[i];
            result = result.concat(action.getAllInputs());
        }
        return result;
    };
    return ActionCollection;
}());
var ActionSet = (function (_super) {
    __extends(ActionSet, _super);
    function ActionSet() {
        var _this = _super.call(this) || this;
        _this.actionStyle = "button";
        _this._actionCollection = new ActionCollection(_this);
        _this._actionCollection.onHideActionCardPane = function () { _this.showBottomSpacer(_this); };
        _this._actionCollection.onShowActionCardPane = function (action) { _this.hideBottomSpacer(_this); };
        return _this;
    }
    ActionSet.prototype.internalRender = function () {
        this._actionCollection.actionStyle = this.actionStyle;
        return this._actionCollection.render();
    };
    ActionSet.prototype.getJsonTypeName = function () {
        return "ActionSet";
    };
    ActionSet.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.actions.separation;
    };
    ActionSet.prototype.validate = function () {
        return this._actionCollection.validate();
    };
    ActionSet.prototype.parse = function (json, itemsCollectionPropertyName) {
        if (itemsCollectionPropertyName === void 0) { itemsCollectionPropertyName = "items"; }
        _super.prototype.parse.call(this, json);
        this.actionStyle = Utils.getValueOrDefault(json["actionStyle"], "button");
        if (json["actions"] != undefined) {
            var jsonActions = json["actions"];
            for (var i = 0; i < jsonActions.length; i++) {
                this.addAction(Action.createAction(jsonActions[i]));
            }
        }
    };
    ActionSet.prototype.addAction = function (action) {
        if (action != null) {
            this._actionCollection.addAction(action);
        }
    };
    ActionSet.prototype.getAllInputs = function () {
        return this._actionCollection.getAllInputs();
    };
    ActionSet.prototype.renderSpeech = function () {
        // TODO: What's the right thing to do here?
        return "";
    };
    Object.defineProperty(ActionSet.prototype, "isInteractive", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    return ActionSet;
}(CardElement));
exports.ActionSet = ActionSet;
var ContainerBase = (function (_super) {
    __extends(ContainerBase, _super);
    function ContainerBase() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._items = [];
        return _this;
    }
    ContainerBase.prototype.showBottomSpacer = function (requestingElement) {
        if ((!requestingElement || this.isLastItem(requestingElement)) && hostConfig.actions.showCard.actionMode == "inlineEdgeToEdge") {
            this._element.style.paddingBottom = this.padding.bottom + "px";
            _super.prototype.showBottomSpacer.call(this, this);
        }
    };
    ContainerBase.prototype.hideBottomSpacer = function (requestingElement) {
        if ((!requestingElement || this.isLastItem(requestingElement)) && hostConfig.actions.showCard.actionMode == "inlineEdgeToEdge") {
            this._element.style.paddingBottom = "0px";
            _super.prototype.hideBottomSpacer.call(this, this);
        }
    };
    ContainerBase.prototype.internalRender = function () {
        var _this = this;
        this._element = document.createElement("div");
        this._element.className = "ac-container";
        if (this.backgroundImage) {
            this._element.style.backgroundImage = "url('" + this.backgroundImage + "')";
            this._element.style.backgroundRepeat = "no-repeat";
            this._element.style.backgroundSize = "cover";
        }
        var backgroundColor = this.getBackgroundColor();
        if (backgroundColor) {
            this._element.style.backgroundColor = Utils.stringToCssColor(backgroundColor);
        }
        if (this.selectAction) {
            this._element.classList.add("ac-selectable");
        }
        this._element.style.paddingTop = this.padding.top + "px";
        this._element.style.paddingRight = this.padding.right + "px";
        this._element.style.paddingBottom = this.padding.bottom + "px";
        this._element.style.paddingLeft = this.padding.left + "px";
        this._element.onclick = function (e) {
            if (_this.selectAction != null) {
                _this.selectAction.execute();
                e.cancelBubble = true;
            }
        };
        if (this._items.length > 0) {
            var renderedElementCount = 0;
            for (var i = 0; i < this._items.length; i++) {
                var renderedElement = isElementAllowed(this._items[i], this.getForbiddenElementTypes()) ? this._items[i].render() : null;
                if (renderedElement != null) {
                    if (renderedElementCount > 0 && this._items[i].separation != "none") {
                        var separationDefinition = this._items[i].separation == "default" ? this._items[i].getDefaultSeparationDefinition() : hostConfig.strongSeparation;
                        Utils.appendChild(this._element, Utils.renderSeparation(separationDefinition, "vertical"));
                    }
                    Utils.appendChild(this._element, renderedElement);
                    renderedElementCount++;
                }
            }
        }
        return this._element;
    };
    ContainerBase.prototype.getBackgroundColor = function () {
        return null;
    };
    Object.defineProperty(ContainerBase.prototype, "padding", {
        get: function () {
            return { left: 0, top: 0, right: 0, bottom: 0 };
        },
        enumerable: true,
        configurable: true
    });
    ContainerBase.prototype.isLastItem = function (item) {
        return this._items.indexOf(item) == (this._items.length - 1);
    };
    ContainerBase.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.container.separation;
    };
    ContainerBase.prototype.validate = function () {
        var result = [];
        for (var i = 0; i < this._items.length; i++) {
            if (!hostConfig.supportsInteractivity && this._items[i].isInteractive) {
                result.push({
                    error: Enums.ValidationError.InteractivityNotAllowed,
                    message: "Interactivity is not allowed."
                });
            }
            if (!isElementAllowed(this._items[i], this.getForbiddenElementTypes())) {
                result.push({
                    error: Enums.ValidationError.InteractivityNotAllowed,
                    message: "Elements of type " + this._items[i].getJsonTypeName() + " are not allowed in this container."
                });
            }
            result = result.concat(this._items[i].validate());
        }
        return result;
    };
    ContainerBase.prototype.parse = function (json, itemsCollectionPropertyName) {
        if (itemsCollectionPropertyName === void 0) { itemsCollectionPropertyName = "items"; }
        _super.prototype.parse.call(this, json);
        this.backgroundImage = json["backgroundImage"];
        if (json[itemsCollectionPropertyName] != null) {
            var items = json[itemsCollectionPropertyName];
            for (var i = 0; i < items.length; i++) {
                var elementType = items[i]["type"];
                var element = AdaptiveCard.elementTypeRegistry.createInstance(elementType);
                if (!element) {
                    raiseParseError({
                        error: Enums.ValidationError.UnknownElementType,
                        message: "Unknown element type: " + elementType
                    });
                }
                else {
                    this.addItem(element);
                    element.parse(items[i]);
                }
            }
        }
        var selectActionJson = json["selectAction"];
        if (selectActionJson != undefined) {
            this.selectAction = Action.createAction(selectActionJson);
            invokeSetParent(this.selectAction, this);
        }
    };
    ContainerBase.prototype.addItem = function (item) {
        if (!item.parent) {
            if (item.isStandalone) {
                this._items.push(item);
                invokeSetParent(item, this);
            }
            else {
                throw new Error("Elements of type " + item.getJsonTypeName() + " cannot be used as standalone elements.");
            }
        }
        else {
            throw new Error("The element already belongs to another container.");
        }
    };
    ContainerBase.prototype.clear = function () {
        this._items = [];
    };
    ContainerBase.prototype.getAllInputs = function () {
        var result = [];
        for (var i = 0; i < this._items.length; i++) {
            var item = this._items[i];
            result = result.concat(item.getAllInputs());
        }
        return result;
    };
    ContainerBase.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak;
        }
        // render each item
        var speak = null;
        if (this._items.length > 0) {
            speak = '';
            for (var i = 0; i < this._items.length; i++) {
                var result = this._items[i].renderSpeech();
                if (result) {
                    speak += result;
                }
            }
        }
        return speak;
    };
    return ContainerBase;
}(CardElement));
exports.ContainerBase = ContainerBase;
var Container = (function (_super) {
    __extends(Container, _super);
    function Container() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.style = "normal";
        return _this;
    }
    Container.prototype.getBackgroundColor = function () {
        return this.style == "normal" ? hostConfig.container.normal.backgroundColor : hostConfig.container.emphasis.backgroundColor;
    };
    Container.prototype.internalRender = function () {
        var renderedContainer = _super.prototype.internalRender.call(this);
        if (renderedContainer) {
            var styleDefinition = this.style == "normal" ? hostConfig.container.normal : hostConfig.container.emphasis;
            if (styleDefinition.borderThickness) {
                renderedContainer.style.borderTop = styleDefinition.borderThickness.top + "px solid";
                renderedContainer.style.borderRight = styleDefinition.borderThickness.right + "px solid";
                renderedContainer.style.borderBottom = styleDefinition.borderThickness.bottom + "px solid";
                renderedContainer.style.borderLeft = styleDefinition.borderThickness.left + "px solid";
            }
            if (styleDefinition.borderColor) {
                renderedContainer.style.borderColor = Utils.stringToCssColor(styleDefinition.borderColor);
            }
        }
        return renderedContainer;
    };
    Object.defineProperty(Container.prototype, "padding", {
        get: function () {
            var styleDefinition = this.style == "normal" ? hostConfig.container.normal : hostConfig.container.emphasis;
            return styleDefinition.padding ? styleDefinition.padding : { top: 0, right: 0, bottom: 0, left: 0 };
        },
        enumerable: true,
        configurable: true
    });
    Container.prototype.getJsonTypeName = function () {
        return "Container";
    };
    Container.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        this.style = Utils.getValueOrDefault(json["style"], "normal");
    };
    return Container;
}(ContainerBase));
exports.Container = Container;
var Column = (function (_super) {
    __extends(Column, _super);
    function Column() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._computedWeight = 0;
        _this.size = "auto";
        return _this;
    }
    Object.defineProperty(Column.prototype, "padding", {
        get: function () {
            return { left: 0, top: 0, right: 0, bottom: 0 };
        },
        enumerable: true,
        configurable: true
    });
    Column.prototype.internalRender = function () {
        var element = _super.prototype.internalRender.call(this);
        if (element) {
            element.style.minWidth = "0";
            if (typeof this.size === "number") {
                element.style.flex = "1 1 " + (this._computedWeight > 0 ? this._computedWeight : this.size) + "%";
            }
            else if (this.size === "auto") {
                element.style.flex = "0 1 auto";
            }
            else {
                element.style.flex = "1 1 auto";
            }
        }
        return element;
    };
    Column.prototype.getJsonTypeName = function () {
        return "Column";
    };
    Column.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.column.separation;
    };
    Column.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        var parsedSize = json["size"];
        var invalidSize = false;
        if (typeof parsedSize === "number") {
            if (parsedSize <= 0) {
                invalidSize = true;
            }
        }
        else if (typeof parsedSize === "string") {
            if (parsedSize != "auto" && parsedSize != "stretch") {
                var sizeAsNumber = parseInt(parsedSize);
                if (!isNaN(sizeAsNumber)) {
                    parsedSize = sizeAsNumber;
                }
                else {
                    invalidSize = true;
                }
            }
        }
        else if (parsedSize) {
            invalidSize = true;
        }
        if (invalidSize) {
            raiseParseError({
                error: Enums.ValidationError.InvalidPropertyValue,
                message: "Invalid column size: " + parsedSize
            });
        }
        else {
            this.size = parsedSize;
        }
    };
    Object.defineProperty(Column.prototype, "isStandalone", {
        get: function () {
            return false;
        },
        enumerable: true,
        configurable: true
    });
    return Column;
}(Container));
exports.Column = Column;
var ColumnSet = (function (_super) {
    __extends(ColumnSet, _super);
    function ColumnSet() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._columns = [];
        return _this;
    }
    ColumnSet.prototype.internalRender = function () {
        if (this._columns.length > 0) {
            var element = document.createElement("div");
            element.style.display = "flex";
            element.style.overflow = "hidden";
            switch (this.horizontalAlignment) {
                case "center":
                    element.style.justifyContent = "center";
                    break;
                case "right":
                    element.style.justifyContent = "flex-end";
                    break;
                default:
                    element.style.justifyContent = "flex-start";
                    break;
            }
            var totalWeight = 0;
            for (var i = 0; i < this._columns.length; i++) {
                if (typeof this._columns[i].size === "number") {
                    totalWeight += this._columns[i].size;
                }
            }
            var renderedColumnCount = 0;
            for (var i = 0; i < this._columns.length; i++) {
                if (typeof this._columns[i].size === "number" && totalWeight > 0) {
                    var computedWeight = 100 / totalWeight * this._columns[i].size;
                    // Best way to emulate "internal" access I know of
                    this._columns[i]["_computedWeight"] = computedWeight;
                }
                var renderedColumn = this._columns[i].render();
                if (renderedColumn != null) {
                    Utils.appendChild(element, renderedColumn);
                    if (this._columns.length > 1 && i < this._columns.length - 1 && this._columns[i + 1].separation != "none") {
                        var separationDefinition = this._columns[i + 1].separation == "default" ? this._columns[i + 1].getDefaultSeparationDefinition() : hostConfig.strongSeparation;
                        if (separationDefinition) {
                            var separator = Utils.renderSeparation(separationDefinition, "horizontal");
                            separator.style.flex = "0 0 auto";
                            Utils.appendChild(element, separator);
                        }
                    }
                    renderedColumnCount++;
                }
            }
            return renderedColumnCount > 0 ? element : null;
        }
        else {
            return null;
        }
    };
    ColumnSet.prototype.getJsonTypeName = function () {
        return "ColumnSet";
    };
    ColumnSet.prototype.getDefaultSeparationDefinition = function () {
        return hostConfig.columnSet.separation;
    };
    ColumnSet.prototype.parse = function (json) {
        _super.prototype.parse.call(this, json);
        if (json["columns"] != null) {
            var jsonColumns = json["columns"];
            for (var i = 0; i < jsonColumns.length; i++) {
                var column = new Column();
                column.parse(jsonColumns[i]);
                this.addColumn(column);
            }
        }
    };
    ColumnSet.prototype.addColumn = function (column) {
        if (!column.parent) {
            this._columns.push(column);
            invokeSetParent(column, this);
        }
        else {
            throw new Error("This column already belongs to another ColumnSet.");
        }
    };
    ColumnSet.prototype.getAllInputs = function () {
        var result = [];
        for (var i = 0; i < this._columns.length; i++) {
            result = result.concat(this._columns[i].getAllInputs());
        }
        return result;
    };
    ColumnSet.prototype.renderSpeech = function () {
        if (this.speak != null) {
            return this.speak;
        }
        // render each item
        var speak = '';
        if (this._columns.length > 0) {
            for (var i = 0; i < this._columns.length; i++) {
                speak += this._columns[i].renderSpeech();
            }
        }
        return speak;
    };
    return ColumnSet;
}(CardElement));
exports.ColumnSet = ColumnSet;
function raiseExecuteActionEvent(action) {
    if (AdaptiveCard.onExecuteAction != null) {
        action.prepare(action.parent.getRootElement().getAllInputs());
        AdaptiveCard.onExecuteAction(action);
    }
}
function raiseInlineCardExpandedEvent(action, isExpanded) {
    if (AdaptiveCard.onInlineCardExpanded != null) {
        AdaptiveCard.onInlineCardExpanded(action, isExpanded);
    }
}
function raiseShowPopupCardEvent(action) {
    if (AdaptiveCard.onShowPopupCard != null) {
        AdaptiveCard.onShowPopupCard(action);
    }
}
function raiseParseError(error) {
    if (AdaptiveCard.onParseError != null) {
        AdaptiveCard.onParseError(error);
    }
}
var TypeRegistry = (function () {
    function TypeRegistry() {
        this._items = [];
    }
    TypeRegistry.prototype.findTypeRegistration = function (typeName) {
        for (var i = 0; i < this._items.length; i++) {
            if (this._items[i].typeName === typeName) {
                return this._items[i];
            }
        }
        return null;
    };
    TypeRegistry.prototype.clear = function () {
        this._items = [];
    };
    TypeRegistry.prototype.registerType = function (typeName, createInstance) {
        var registrationInfo = this.findTypeRegistration(typeName);
        if (registrationInfo != null) {
            registrationInfo.createInstance = createInstance;
        }
        else {
            registrationInfo = {
                typeName: typeName,
                createInstance: createInstance
            };
            this._items.push(registrationInfo);
        }
    };
    TypeRegistry.prototype.unregisterType = function (typeName) {
        for (var i = 0; i < this._items.length; i++) {
            if (this._items[i].typeName === typeName) {
                this._items = this._items.splice(i, 1);
                return;
            }
        }
    };
    TypeRegistry.prototype.createInstance = function (typeName) {
        var registrationInfo = this.findTypeRegistration(typeName);
        return registrationInfo ? registrationInfo.createInstance() : null;
    };
    return TypeRegistry;
}());
exports.TypeRegistry = TypeRegistry;
var ContainerWithActions = (function (_super) {
    __extends(ContainerWithActions, _super);
    function ContainerWithActions() {
        var _this = _super.call(this) || this;
        _this.actionStyle = "button";
        _this._actionCollection = new ActionCollection(_this);
        _this._actionCollection.onHideActionCardPane = function () { _this.showBottomSpacer(null); };
        _this._actionCollection.onShowActionCardPane = function (action) { _this.hideBottomSpacer(null); };
        return _this;
    }
    ContainerWithActions.prototype.internalRender = function () {
        _super.prototype.internalRender.call(this);
        this._actionCollection.actionStyle = this.actionStyle;
        var renderedActions = this._actionCollection.render();
        if (renderedActions) {
            Utils.appendChild(this._element, Utils.renderSeparation(hostConfig.actions.separation, "vertical"));
            Utils.appendChild(this._element, renderedActions);
        }
        return this._element.children.length > 0 ? this._element : null;
    };
    ContainerWithActions.prototype.parse = function (json, itemsCollectionPropertyName) {
        if (itemsCollectionPropertyName === void 0) { itemsCollectionPropertyName = "items"; }
        _super.prototype.parse.call(this, json, itemsCollectionPropertyName);
        this.actionStyle = Utils.getValueOrDefault(json["actionStyle"], "button");
        if (json["actions"] != undefined) {
            var jsonActions = json["actions"];
            for (var i = 0; i < jsonActions.length; i++) {
                var action = Action.createAction(jsonActions[i]);
                if (action != null) {
                    this.addAction(action);
                }
            }
        }
    };
    ContainerWithActions.prototype.isLastItem = function (item) {
        return _super.prototype.isLastItem.call(this, item) && this._actionCollection.items.length == 0;
    };
    ContainerWithActions.prototype.addAction = function (action) {
        this._actionCollection.addAction(action);
    };
    ContainerWithActions.prototype.clear = function () {
        _super.prototype.clear.call(this);
        this._actionCollection.clear();
    };
    ContainerWithActions.prototype.getAllInputs = function () {
        return _super.prototype.getAllInputs.call(this).concat(this._actionCollection.getAllInputs());
    };
    Object.defineProperty(ContainerWithActions.prototype, "isStandalone", {
        get: function () {
            return false;
        },
        enumerable: true,
        configurable: true
    });
    return ContainerWithActions;
}(ContainerBase));
exports.ContainerWithActions = ContainerWithActions;
var AdaptiveCard = (function (_super) {
    __extends(AdaptiveCard, _super);
    function AdaptiveCard() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.minVersion = { major: 1, minor: 0 };
        return _this;
    }
    AdaptiveCard.initialize = function () {
        AdaptiveCard.elementTypeRegistry.clear();
        AdaptiveCard.elementTypeRegistry.registerType("Container", function () { return new Container(); });
        AdaptiveCard.elementTypeRegistry.registerType("TextBlock", function () { return new TextBlock(); });
        AdaptiveCard.elementTypeRegistry.registerType("Image", function () { return new Image(); });
        AdaptiveCard.elementTypeRegistry.registerType("ImageSet", function () { return new ImageSet(); });
        AdaptiveCard.elementTypeRegistry.registerType("FactSet", function () { return new FactSet(); });
        AdaptiveCard.elementTypeRegistry.registerType("ColumnSet", function () { return new ColumnSet(); });
        AdaptiveCard.elementTypeRegistry.registerType("ActionSet", function () { return new ActionSet(); });
        AdaptiveCard.elementTypeRegistry.registerType("Input.Text", function () { return new TextInput(); });
        AdaptiveCard.elementTypeRegistry.registerType("Input.Date", function () { return new DateInput(); });
        AdaptiveCard.elementTypeRegistry.registerType("Input.Time", function () { return new TimeInput(); });
        AdaptiveCard.elementTypeRegistry.registerType("Input.Number", function () { return new NumberInput(); });
        AdaptiveCard.elementTypeRegistry.registerType("Input.ChoiceSet", function () { return new ChoiceSetInput(); });
        AdaptiveCard.elementTypeRegistry.registerType("Input.Toggle", function () { return new ToggleInput(); });
        AdaptiveCard.actionTypeRegistry.clear();
        AdaptiveCard.actionTypeRegistry.registerType("Action.Http", function () { return new HttpAction(); });
        AdaptiveCard.actionTypeRegistry.registerType("Action.OpenUrl", function () { return new OpenUrlAction(); });
        AdaptiveCard.actionTypeRegistry.registerType("Action.Submit", function () { return new SubmitAction(); });
        AdaptiveCard.actionTypeRegistry.registerType("Action.ShowCard", function () { return new ShowCardAction(); });
    };
    AdaptiveCard.prototype.isVersionSupported = function () {
        var unsupportedVersion = (AdaptiveCard.currentVersion.major < this.minVersion.major) ||
            (AdaptiveCard.currentVersion.major == this.minVersion.major && AdaptiveCard.currentVersion.minor < this.minVersion.minor);
        return !unsupportedVersion;
    };
    AdaptiveCard.prototype.getBackgroundColor = function () {
        return hostConfig.adaptiveCard.backgroundColor;
    };
    Object.defineProperty(AdaptiveCard.prototype, "padding", {
        get: function () {
            return hostConfig.adaptiveCard.padding;
        },
        enumerable: true,
        configurable: true
    });
    AdaptiveCard.prototype.getJsonTypeName = function () {
        return "AdaptiveCard";
    };
    AdaptiveCard.prototype.validate = function () {
        var result = [];
        if (this._cardTypeName != "AdaptiveCard") {
            result.push({
                error: Enums.ValidationError.MissingCardType,
                message: "Invalid or missing card type. Make sure the card's type property is set to \"AdaptiveCard\"."
            });
        }
        if (!this.isVersionSupported()) {
            result.push({
                error: Enums.ValidationError.UnsupportedCardVersion,
                message: "The specified card version is not supported."
            });
        }
        return result.concat(_super.prototype.validate.call(this));
    };
    AdaptiveCard.prototype.parse = function (json) {
        this._cardTypeName = json["type"];
        var minVersion = json["minVersion"];
        var regEx = /(\d+).(\d+)/gi;
        var matches = regEx.exec(minVersion);
        if (matches != null && matches.length == 3) {
            this.minVersion.major = parseInt(matches[1]);
            this.minVersion.minor = parseInt(matches[2]);
        }
        this.fallbackText = json["fallbackText"];
        _super.prototype.parse.call(this, json, "body");
    };
    AdaptiveCard.prototype.render = function () {
        var renderedCard;
        if (!this.isVersionSupported()) {
            renderedCard = document.createElement("div");
            renderedCard.innerHTML = this.fallbackText ? this.fallbackText : "The specified card version is not supported.";
            return renderedCard;
        }
        else {
            return _super.prototype.render.call(this);
        }
    };
    return AdaptiveCard;
}(ContainerWithActions));
AdaptiveCard.currentVersion = { major: 1, minor: 0 };
AdaptiveCard.elementTypeRegistry = new TypeRegistry();
AdaptiveCard.actionTypeRegistry = new TypeRegistry();
AdaptiveCard.onExecuteAction = null;
AdaptiveCard.onShowPopupCard = null;
AdaptiveCard.onInlineCardExpanded = null;
AdaptiveCard.onParseError = null;
exports.AdaptiveCard = AdaptiveCard;
// This calls acts as a static constructor (see https://github.com/Microsoft/TypeScript/issues/265)
AdaptiveCard.initialize();
var InlineAdaptiveCard = (function (_super) {
    __extends(InlineAdaptiveCard, _super);
    function InlineAdaptiveCard() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(InlineAdaptiveCard.prototype, "padding", {
        get: function () {
            return hostConfig.actions.showCard.padding;
        },
        enumerable: true,
        configurable: true
    });
    InlineAdaptiveCard.prototype.getBackgroundColor = function () {
        return null;
    };
    InlineAdaptiveCard.prototype.getForbiddenActionTypes = function () {
        return [ShowCardAction];
    };
    return InlineAdaptiveCard;
}(AdaptiveCard));
var defaultHostConfig = {
    supportsInteractivity: true,
    strongSeparation: {
        spacing: 40,
        lineThickness: 1,
        lineColor: "#EEEEEE"
    },
    fontFamily: "Segoe UI",
    fontSizes: {
        small: 8,
        normal: 10,
        medium: 12,
        large: 14,
        extraLarge: 16
    },
    fontWeights: {
        lighter: 200,
        normal: 400,
        bolder: 600
    },
    colors: {
        dark: {
            normal: "#0000FF",
            subtle: "#222222"
        },
        light: {
            normal: "#FFFFFF",
            subtle: "#DDDDDD"
        },
        accent: {
            normal: "#0000FF",
            subtle: "#0000DD"
        },
        attention: {
            normal: "#FF6600",
            subtle: "#DD4400"
        },
        good: {
            normal: "#00FF00",
            subtle: "#00DD00"
        },
        warning: {
            normal: "#FF0000",
            subtle: "#DD0000"
        }
    },
    imageSizes: {
        small: 40,
        medium: 80,
        large: 160
    },
    actions: {
        maxActions: 5,
        separation: {
            spacing: 20
        },
        buttonSpacing: 20,
        showCard: {
            actionMode: "inlineEdgeToEdge",
            inlineTopMargin: 16,
            backgroundColor: "#22000000",
            padding: {
                top: 16,
                right: 16,
                bottom: 16,
                left: 16
            }
        },
        actionsOrientation: "horizontal",
        actionAlignment: "left"
    },
    adaptiveCard: {
        backgroundColor: "#00000000",
        padding: {
            left: 20,
            top: 20,
            right: 20,
            bottom: 20
        }
    },
    container: {
        separation: {
            spacing: 20
        },
        normal: {},
        emphasis: {
            backgroundColor: "#EEEEEE",
            borderColor: "#AAAAAA",
            borderThickness: {
                top: 1,
                right: 1,
                bottom: 1,
                left: 1
            },
            padding: {
                top: 10,
                right: 10,
                bottom: 10,
                left: 10
            }
        }
    },
    textBlock: {
        color: "dark",
        separations: {
            small: {
                spacing: 20,
            },
            normal: {
                spacing: 20
            },
            medium: {
                spacing: 20
            },
            large: {
                spacing: 20
            },
            extraLarge: {
                spacing: 20
            }
        }
    },
    image: {
        size: "medium",
        separation: {
            spacing: 20
        }
    },
    imageSet: {
        imageSize: "medium",
        separation: {
            spacing: 20
        }
    },
    factSet: {
        separation: {
            spacing: 20
        },
        title: {
            color: "dark",
            size: "normal",
            isSubtle: false,
            weight: "bolder",
            wrap: true,
            maxWidth: 150
        },
        value: {
            color: "dark",
            size: "normal",
            isSubtle: false,
            weight: "normal",
            wrap: true
        },
        spacing: 10
    },
    input: {
        separation: {
            spacing: 20
        }
    },
    columnSet: {
        separation: {
            spacing: 20
        }
    },
    column: {
        separation: {
            spacing: 20
        }
    }
};
var hostConfig = defaultHostConfig;
function setHostConfig(configuration) {
    hostConfig = configuration;
}
exports.setHostConfig = setHostConfig;
function resetHostConfig() {
    hostConfig = defaultHostConfig;
}
exports.resetHostConfig = resetHostConfig;


/***/ }),
/* 15 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var Utils = __webpack_require__(10);
function parseSpacingDefinition(obj) {
    return obj ? {
        top: obj["top"],
        right: obj["right"],
        bottom: obj["bottom"],
        left: obj["left"]
    } : null;
}
function parseColorDefinition(obj) {
    return obj ? {
        normal: obj["normal"],
        subtle: obj["subtle"]
    } : null;
}
function parseSeparationDefinition(obj) {
    return obj ? {
        spacing: obj["spacing"],
        lineThickness: obj["lineThickness"],
        lineColor: obj["lineColor"]
    } : null;
}
function parseAdaptiveCardConfiguration(obj) {
    return obj ? {
        backgroundColor: obj["backgroundColor"],
        padding: parseSpacingDefinition(obj["padding"])
    } : null;
}
function parseTextBlockConfiguration(obj) {
    return obj ? {
        color: obj["color"],
        separations: {
            small: parseSeparationDefinition(obj["separations"]["small"]),
            normal: parseSeparationDefinition(obj["separations"]["normal"]),
            medium: parseSeparationDefinition(obj["separations"]["medium"]),
            large: parseSeparationDefinition(obj["separations"]["large"]),
            extraLarge: parseSeparationDefinition(obj["separations"]["extraLarge"])
        }
    } : null;
}
function parseContainerStyleDefinition(obj) {
    return obj ? {
        backgroundColor: obj["backgroundColor"],
        padding: parseSpacingDefinition(obj["padding"]),
        borderColor: obj["borderColor"],
        borderThickness: parseSpacingDefinition(obj["borderThickness"])
    } : null;
}
function parseContainerConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"]),
        normal: parseContainerStyleDefinition(obj["normal"]),
        emphasis: parseContainerStyleDefinition(obj["emphasis"])
    } : null;
}
function parseImageConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"]),
        size: obj["size"]
    } : null;
}
function parseImageSetConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"]),
        imageSize: obj["imageSize"]
    } : null;
}
function parseFactTextDefinition(obj) {
    return obj ? {
        size: Utils.getValueOrDefault(obj["size"], "normal"),
        color: Utils.getValueOrDefault(obj["color"], "dark"),
        isSubtle: obj["isSubtle"],
        weight: Utils.getValueOrDefault(obj["weight"], "normal"),
        wrap: obj["wrap"]
    } : null;
}
function parseFactTitleDefinition(obj) {
    var result = parseFactTextDefinition(obj);
    if (result) {
        result.maxWidth = obj["maxWidth"];
    }
    return result;
}
function parseFactSetConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"]),
        title: parseFactTitleDefinition(obj["title"]),
        value: parseFactTextDefinition(obj["value"]),
        spacing: obj["spacing"]
    } : null;
}
function parseColumnSetConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"])
    } : null;
}
function parseColumnConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"])
    } : null;
}
function parseShowCardActionConfiguration(obj) {
    return obj ? {
        actionMode: Utils.getValueOrDefault(obj["actionMode"], "inlineEdgeToEdge"),
        inlineTopMargin: obj["inlineTopMargin"],
        backgroundColor: obj["backgroundColor"],
        padding: parseSpacingDefinition(obj["padding"])
    } : null;
}
function parseActionsConfiguration(obj) {
    return obj ? {
        maxActions: obj["maxActions"],
        separation: parseSeparationDefinition(obj["separation"]),
        buttonSpacing: obj["buttonSpacing"],
        showCard: parseShowCardActionConfiguration(obj["showCard"]),
        preExpandSingleShowCardAction: Utils.getValueOrDefault(obj["preExpandSingleShowCardAction"], false),
        actionsOrientation: Utils.getValueOrDefault(obj["actionsOrientation"], "horizontal"),
        actionAlignment: Utils.getValueOrDefault(obj["actionAlignment"], "left"),
    } : null;
}
function parseInputConfiguration(obj) {
    return obj ? {
        separation: parseSeparationDefinition(obj["separation"])
    } : null;
}
function parseHostConfig(serializedConfiguration) {
    var obj = JSON.parse(serializedConfiguration);
    return obj ? {
        supportsInteractivity: obj["supportsInteractivity"],
        fontFamily: obj["fontFamily"],
        fontSizes: {
            small: obj["fontSizes"]["small"],
            normal: obj["fontSizes"]["normal"],
            medium: obj["fontSizes"]["medium"],
            large: obj["fontSizes"]["large"],
            extraLarge: obj["fontSizes"]["extraLarge"]
        },
        fontWeights: {
            lighter: obj["fontWeights"]["lighter"],
            normal: obj["fontWeights"]["normal"],
            bolder: obj["fontWeights"]["bolder"]
        },
        imageSizes: {
            small: obj["imageSizes"]["small"],
            medium: obj["imageSizes"]["medium"],
            large: obj["imageSizes"]["large"],
        },
        colors: {
            dark: parseColorDefinition(obj["colors"]["dark"]),
            light: parseColorDefinition(obj["colors"]["light"]),
            accent: parseColorDefinition(obj["colors"]["accent"]),
            good: parseColorDefinition(obj["colors"]["good"]),
            warning: parseColorDefinition(obj["colors"]["warning"]),
            attention: parseColorDefinition(obj["colors"]["attention"])
        },
        strongSeparation: parseSeparationDefinition(obj["strongSeparation"]),
        actions: parseActionsConfiguration(obj["actions"]),
        adaptiveCard: parseAdaptiveCardConfiguration(obj["adaptiveCard"]),
        container: parseContainerConfiguration(obj["container"]),
        textBlock: parseTextBlockConfiguration(obj["textBlock"]),
        image: parseImageConfiguration(obj["image"]),
        imageSet: parseImageSetConfiguration(obj["imageSet"]),
        factSet: parseFactSetConfiguration(obj["factSet"]),
        column: parseColumnConfiguration(obj["column"]),
        columnSet: parseColumnSetConfiguration(obj["columnSet"]),
        input: parseInputConfiguration(obj["input"])
    } : null;
}
exports.parseHostConfig = parseHostConfig;


/***/ }),
/* 16 */
/***/ (function(module, exports) {

module.exports = {
	"Aacute": "Ã",
	"aacute": "Ã¡",
	"Abreve": "Ä‚",
	"abreve": "Äƒ",
	"ac": "âˆ¾",
	"acd": "âˆ¿",
	"acE": "âˆ¾Ì³",
	"Acirc": "Ã‚",
	"acirc": "Ã¢",
	"acute": "Â´",
	"Acy": "Ð",
	"acy": "Ð°",
	"AElig": "Ã†",
	"aelig": "Ã¦",
	"af": "â¡",
	"Afr": "ð”„",
	"afr": "ð”ž",
	"Agrave": "Ã€",
	"agrave": "Ã ",
	"alefsym": "â„µ",
	"aleph": "â„µ",
	"Alpha": "Î‘",
	"alpha": "Î±",
	"Amacr": "Ä€",
	"amacr": "Ä",
	"amalg": "â¨¿",
	"amp": "&",
	"AMP": "&",
	"andand": "â©•",
	"And": "â©“",
	"and": "âˆ§",
	"andd": "â©œ",
	"andslope": "â©˜",
	"andv": "â©š",
	"ang": "âˆ ",
	"ange": "â¦¤",
	"angle": "âˆ ",
	"angmsdaa": "â¦¨",
	"angmsdab": "â¦©",
	"angmsdac": "â¦ª",
	"angmsdad": "â¦«",
	"angmsdae": "â¦¬",
	"angmsdaf": "â¦­",
	"angmsdag": "â¦®",
	"angmsdah": "â¦¯",
	"angmsd": "âˆ¡",
	"angrt": "âˆŸ",
	"angrtvb": "âŠ¾",
	"angrtvbd": "â¦",
	"angsph": "âˆ¢",
	"angst": "Ã…",
	"angzarr": "â¼",
	"Aogon": "Ä„",
	"aogon": "Ä…",
	"Aopf": "ð”¸",
	"aopf": "ð•’",
	"apacir": "â©¯",
	"ap": "â‰ˆ",
	"apE": "â©°",
	"ape": "â‰Š",
	"apid": "â‰‹",
	"apos": "'",
	"ApplyFunction": "â¡",
	"approx": "â‰ˆ",
	"approxeq": "â‰Š",
	"Aring": "Ã…",
	"aring": "Ã¥",
	"Ascr": "ð’œ",
	"ascr": "ð’¶",
	"Assign": "â‰”",
	"ast": "*",
	"asymp": "â‰ˆ",
	"asympeq": "â‰",
	"Atilde": "Ãƒ",
	"atilde": "Ã£",
	"Auml": "Ã„",
	"auml": "Ã¤",
	"awconint": "âˆ³",
	"awint": "â¨‘",
	"backcong": "â‰Œ",
	"backepsilon": "Ï¶",
	"backprime": "â€µ",
	"backsim": "âˆ½",
	"backsimeq": "â‹",
	"Backslash": "âˆ–",
	"Barv": "â«§",
	"barvee": "âŠ½",
	"barwed": "âŒ…",
	"Barwed": "âŒ†",
	"barwedge": "âŒ…",
	"bbrk": "âŽµ",
	"bbrktbrk": "âŽ¶",
	"bcong": "â‰Œ",
	"Bcy": "Ð‘",
	"bcy": "Ð±",
	"bdquo": "â€ž",
	"becaus": "âˆµ",
	"because": "âˆµ",
	"Because": "âˆµ",
	"bemptyv": "â¦°",
	"bepsi": "Ï¶",
	"bernou": "â„¬",
	"Bernoullis": "â„¬",
	"Beta": "Î’",
	"beta": "Î²",
	"beth": "â„¶",
	"between": "â‰¬",
	"Bfr": "ð”…",
	"bfr": "ð”Ÿ",
	"bigcap": "â‹‚",
	"bigcirc": "â—¯",
	"bigcup": "â‹ƒ",
	"bigodot": "â¨€",
	"bigoplus": "â¨",
	"bigotimes": "â¨‚",
	"bigsqcup": "â¨†",
	"bigstar": "â˜…",
	"bigtriangledown": "â–½",
	"bigtriangleup": "â–³",
	"biguplus": "â¨„",
	"bigvee": "â‹",
	"bigwedge": "â‹€",
	"bkarow": "â¤",
	"blacklozenge": "â§«",
	"blacksquare": "â–ª",
	"blacktriangle": "â–´",
	"blacktriangledown": "â–¾",
	"blacktriangleleft": "â—‚",
	"blacktriangleright": "â–¸",
	"blank": "â£",
	"blk12": "â–’",
	"blk14": "â–‘",
	"blk34": "â–“",
	"block": "â–ˆ",
	"bne": "=âƒ¥",
	"bnequiv": "â‰¡âƒ¥",
	"bNot": "â«­",
	"bnot": "âŒ",
	"Bopf": "ð”¹",
	"bopf": "ð•“",
	"bot": "âŠ¥",
	"bottom": "âŠ¥",
	"bowtie": "â‹ˆ",
	"boxbox": "â§‰",
	"boxdl": "â”",
	"boxdL": "â••",
	"boxDl": "â•–",
	"boxDL": "â•—",
	"boxdr": "â”Œ",
	"boxdR": "â•’",
	"boxDr": "â•“",
	"boxDR": "â•”",
	"boxh": "â”€",
	"boxH": "â•",
	"boxhd": "â”¬",
	"boxHd": "â•¤",
	"boxhD": "â•¥",
	"boxHD": "â•¦",
	"boxhu": "â”´",
	"boxHu": "â•§",
	"boxhU": "â•¨",
	"boxHU": "â•©",
	"boxminus": "âŠŸ",
	"boxplus": "âŠž",
	"boxtimes": "âŠ ",
	"boxul": "â”˜",
	"boxuL": "â•›",
	"boxUl": "â•œ",
	"boxUL": "â•",
	"boxur": "â””",
	"boxuR": "â•˜",
	"boxUr": "â•™",
	"boxUR": "â•š",
	"boxv": "â”‚",
	"boxV": "â•‘",
	"boxvh": "â”¼",
	"boxvH": "â•ª",
	"boxVh": "â•«",
	"boxVH": "â•¬",
	"boxvl": "â”¤",
	"boxvL": "â•¡",
	"boxVl": "â•¢",
	"boxVL": "â•£",
	"boxvr": "â”œ",
	"boxvR": "â•ž",
	"boxVr": "â•Ÿ",
	"boxVR": "â• ",
	"bprime": "â€µ",
	"breve": "Ë˜",
	"Breve": "Ë˜",
	"brvbar": "Â¦",
	"bscr": "ð’·",
	"Bscr": "â„¬",
	"bsemi": "â",
	"bsim": "âˆ½",
	"bsime": "â‹",
	"bsolb": "â§…",
	"bsol": "\\",
	"bsolhsub": "âŸˆ",
	"bull": "â€¢",
	"bullet": "â€¢",
	"bump": "â‰Ž",
	"bumpE": "âª®",
	"bumpe": "â‰",
	"Bumpeq": "â‰Ž",
	"bumpeq": "â‰",
	"Cacute": "Ä†",
	"cacute": "Ä‡",
	"capand": "â©„",
	"capbrcup": "â©‰",
	"capcap": "â©‹",
	"cap": "âˆ©",
	"Cap": "â‹’",
	"capcup": "â©‡",
	"capdot": "â©€",
	"CapitalDifferentialD": "â……",
	"caps": "âˆ©ï¸€",
	"caret": "â",
	"caron": "Ë‡",
	"Cayleys": "â„­",
	"ccaps": "â©",
	"Ccaron": "ÄŒ",
	"ccaron": "Ä",
	"Ccedil": "Ã‡",
	"ccedil": "Ã§",
	"Ccirc": "Äˆ",
	"ccirc": "Ä‰",
	"Cconint": "âˆ°",
	"ccups": "â©Œ",
	"ccupssm": "â©",
	"Cdot": "ÄŠ",
	"cdot": "Ä‹",
	"cedil": "Â¸",
	"Cedilla": "Â¸",
	"cemptyv": "â¦²",
	"cent": "Â¢",
	"centerdot": "Â·",
	"CenterDot": "Â·",
	"cfr": "ð” ",
	"Cfr": "â„­",
	"CHcy": "Ð§",
	"chcy": "Ñ‡",
	"check": "âœ“",
	"checkmark": "âœ“",
	"Chi": "Î§",
	"chi": "Ï‡",
	"circ": "Ë†",
	"circeq": "â‰—",
	"circlearrowleft": "â†º",
	"circlearrowright": "â†»",
	"circledast": "âŠ›",
	"circledcirc": "âŠš",
	"circleddash": "âŠ",
	"CircleDot": "âŠ™",
	"circledR": "Â®",
	"circledS": "â“ˆ",
	"CircleMinus": "âŠ–",
	"CirclePlus": "âŠ•",
	"CircleTimes": "âŠ—",
	"cir": "â—‹",
	"cirE": "â§ƒ",
	"cire": "â‰—",
	"cirfnint": "â¨",
	"cirmid": "â«¯",
	"cirscir": "â§‚",
	"ClockwiseContourIntegral": "âˆ²",
	"CloseCurlyDoubleQuote": "â€",
	"CloseCurlyQuote": "â€™",
	"clubs": "â™£",
	"clubsuit": "â™£",
	"colon": ":",
	"Colon": "âˆ·",
	"Colone": "â©´",
	"colone": "â‰”",
	"coloneq": "â‰”",
	"comma": ",",
	"commat": "@",
	"comp": "âˆ",
	"compfn": "âˆ˜",
	"complement": "âˆ",
	"complexes": "â„‚",
	"cong": "â‰…",
	"congdot": "â©­",
	"Congruent": "â‰¡",
	"conint": "âˆ®",
	"Conint": "âˆ¯",
	"ContourIntegral": "âˆ®",
	"copf": "ð•”",
	"Copf": "â„‚",
	"coprod": "âˆ",
	"Coproduct": "âˆ",
	"copy": "Â©",
	"COPY": "Â©",
	"copysr": "â„—",
	"CounterClockwiseContourIntegral": "âˆ³",
	"crarr": "â†µ",
	"cross": "âœ—",
	"Cross": "â¨¯",
	"Cscr": "ð’ž",
	"cscr": "ð’¸",
	"csub": "â«",
	"csube": "â«‘",
	"csup": "â«",
	"csupe": "â«’",
	"ctdot": "â‹¯",
	"cudarrl": "â¤¸",
	"cudarrr": "â¤µ",
	"cuepr": "â‹ž",
	"cuesc": "â‹Ÿ",
	"cularr": "â†¶",
	"cularrp": "â¤½",
	"cupbrcap": "â©ˆ",
	"cupcap": "â©†",
	"CupCap": "â‰",
	"cup": "âˆª",
	"Cup": "â‹“",
	"cupcup": "â©Š",
	"cupdot": "âŠ",
	"cupor": "â©…",
	"cups": "âˆªï¸€",
	"curarr": "â†·",
	"curarrm": "â¤¼",
	"curlyeqprec": "â‹ž",
	"curlyeqsucc": "â‹Ÿ",
	"curlyvee": "â‹Ž",
	"curlywedge": "â‹",
	"curren": "Â¤",
	"curvearrowleft": "â†¶",
	"curvearrowright": "â†·",
	"cuvee": "â‹Ž",
	"cuwed": "â‹",
	"cwconint": "âˆ²",
	"cwint": "âˆ±",
	"cylcty": "âŒ­",
	"dagger": "â€ ",
	"Dagger": "â€¡",
	"daleth": "â„¸",
	"darr": "â†“",
	"Darr": "â†¡",
	"dArr": "â‡“",
	"dash": "â€",
	"Dashv": "â«¤",
	"dashv": "âŠ£",
	"dbkarow": "â¤",
	"dblac": "Ë",
	"Dcaron": "ÄŽ",
	"dcaron": "Ä",
	"Dcy": "Ð”",
	"dcy": "Ð´",
	"ddagger": "â€¡",
	"ddarr": "â‡Š",
	"DD": "â……",
	"dd": "â…†",
	"DDotrahd": "â¤‘",
	"ddotseq": "â©·",
	"deg": "Â°",
	"Del": "âˆ‡",
	"Delta": "Î”",
	"delta": "Î´",
	"demptyv": "â¦±",
	"dfisht": "â¥¿",
	"Dfr": "ð”‡",
	"dfr": "ð”¡",
	"dHar": "â¥¥",
	"dharl": "â‡ƒ",
	"dharr": "â‡‚",
	"DiacriticalAcute": "Â´",
	"DiacriticalDot": "Ë™",
	"DiacriticalDoubleAcute": "Ë",
	"DiacriticalGrave": "`",
	"DiacriticalTilde": "Ëœ",
	"diam": "â‹„",
	"diamond": "â‹„",
	"Diamond": "â‹„",
	"diamondsuit": "â™¦",
	"diams": "â™¦",
	"die": "Â¨",
	"DifferentialD": "â…†",
	"digamma": "Ï",
	"disin": "â‹²",
	"div": "Ã·",
	"divide": "Ã·",
	"divideontimes": "â‹‡",
	"divonx": "â‹‡",
	"DJcy": "Ð‚",
	"djcy": "Ñ’",
	"dlcorn": "âŒž",
	"dlcrop": "âŒ",
	"dollar": "$",
	"Dopf": "ð”»",
	"dopf": "ð••",
	"Dot": "Â¨",
	"dot": "Ë™",
	"DotDot": "âƒœ",
	"doteq": "â‰",
	"doteqdot": "â‰‘",
	"DotEqual": "â‰",
	"dotminus": "âˆ¸",
	"dotplus": "âˆ”",
	"dotsquare": "âŠ¡",
	"doublebarwedge": "âŒ†",
	"DoubleContourIntegral": "âˆ¯",
	"DoubleDot": "Â¨",
	"DoubleDownArrow": "â‡“",
	"DoubleLeftArrow": "â‡",
	"DoubleLeftRightArrow": "â‡”",
	"DoubleLeftTee": "â«¤",
	"DoubleLongLeftArrow": "âŸ¸",
	"DoubleLongLeftRightArrow": "âŸº",
	"DoubleLongRightArrow": "âŸ¹",
	"DoubleRightArrow": "â‡’",
	"DoubleRightTee": "âŠ¨",
	"DoubleUpArrow": "â‡‘",
	"DoubleUpDownArrow": "â‡•",
	"DoubleVerticalBar": "âˆ¥",
	"DownArrowBar": "â¤“",
	"downarrow": "â†“",
	"DownArrow": "â†“",
	"Downarrow": "â‡“",
	"DownArrowUpArrow": "â‡µ",
	"DownBreve": "Ì‘",
	"downdownarrows": "â‡Š",
	"downharpoonleft": "â‡ƒ",
	"downharpoonright": "â‡‚",
	"DownLeftRightVector": "â¥",
	"DownLeftTeeVector": "â¥ž",
	"DownLeftVectorBar": "â¥–",
	"DownLeftVector": "â†½",
	"DownRightTeeVector": "â¥Ÿ",
	"DownRightVectorBar": "â¥—",
	"DownRightVector": "â‡",
	"DownTeeArrow": "â†§",
	"DownTee": "âŠ¤",
	"drbkarow": "â¤",
	"drcorn": "âŒŸ",
	"drcrop": "âŒŒ",
	"Dscr": "ð’Ÿ",
	"dscr": "ð’¹",
	"DScy": "Ð…",
	"dscy": "Ñ•",
	"dsol": "â§¶",
	"Dstrok": "Ä",
	"dstrok": "Ä‘",
	"dtdot": "â‹±",
	"dtri": "â–¿",
	"dtrif": "â–¾",
	"duarr": "â‡µ",
	"duhar": "â¥¯",
	"dwangle": "â¦¦",
	"DZcy": "Ð",
	"dzcy": "ÑŸ",
	"dzigrarr": "âŸ¿",
	"Eacute": "Ã‰",
	"eacute": "Ã©",
	"easter": "â©®",
	"Ecaron": "Äš",
	"ecaron": "Ä›",
	"Ecirc": "ÃŠ",
	"ecirc": "Ãª",
	"ecir": "â‰–",
	"ecolon": "â‰•",
	"Ecy": "Ð­",
	"ecy": "Ñ",
	"eDDot": "â©·",
	"Edot": "Ä–",
	"edot": "Ä—",
	"eDot": "â‰‘",
	"ee": "â…‡",
	"efDot": "â‰’",
	"Efr": "ð”ˆ",
	"efr": "ð”¢",
	"eg": "âªš",
	"Egrave": "Ãˆ",
	"egrave": "Ã¨",
	"egs": "âª–",
	"egsdot": "âª˜",
	"el": "âª™",
	"Element": "âˆˆ",
	"elinters": "â§",
	"ell": "â„“",
	"els": "âª•",
	"elsdot": "âª—",
	"Emacr": "Ä’",
	"emacr": "Ä“",
	"empty": "âˆ…",
	"emptyset": "âˆ…",
	"EmptySmallSquare": "â—»",
	"emptyv": "âˆ…",
	"EmptyVerySmallSquare": "â–«",
	"emsp13": "â€„",
	"emsp14": "â€…",
	"emsp": "â€ƒ",
	"ENG": "ÅŠ",
	"eng": "Å‹",
	"ensp": "â€‚",
	"Eogon": "Ä˜",
	"eogon": "Ä™",
	"Eopf": "ð”¼",
	"eopf": "ð•–",
	"epar": "â‹•",
	"eparsl": "â§£",
	"eplus": "â©±",
	"epsi": "Îµ",
	"Epsilon": "Î•",
	"epsilon": "Îµ",
	"epsiv": "Ïµ",
	"eqcirc": "â‰–",
	"eqcolon": "â‰•",
	"eqsim": "â‰‚",
	"eqslantgtr": "âª–",
	"eqslantless": "âª•",
	"Equal": "â©µ",
	"equals": "=",
	"EqualTilde": "â‰‚",
	"equest": "â‰Ÿ",
	"Equilibrium": "â‡Œ",
	"equiv": "â‰¡",
	"equivDD": "â©¸",
	"eqvparsl": "â§¥",
	"erarr": "â¥±",
	"erDot": "â‰“",
	"escr": "â„¯",
	"Escr": "â„°",
	"esdot": "â‰",
	"Esim": "â©³",
	"esim": "â‰‚",
	"Eta": "Î—",
	"eta": "Î·",
	"ETH": "Ã",
	"eth": "Ã°",
	"Euml": "Ã‹",
	"euml": "Ã«",
	"euro": "â‚¬",
	"excl": "!",
	"exist": "âˆƒ",
	"Exists": "âˆƒ",
	"expectation": "â„°",
	"exponentiale": "â…‡",
	"ExponentialE": "â…‡",
	"fallingdotseq": "â‰’",
	"Fcy": "Ð¤",
	"fcy": "Ñ„",
	"female": "â™€",
	"ffilig": "ï¬ƒ",
	"fflig": "ï¬€",
	"ffllig": "ï¬„",
	"Ffr": "ð”‰",
	"ffr": "ð”£",
	"filig": "ï¬",
	"FilledSmallSquare": "â—¼",
	"FilledVerySmallSquare": "â–ª",
	"fjlig": "fj",
	"flat": "â™­",
	"fllig": "ï¬‚",
	"fltns": "â–±",
	"fnof": "Æ’",
	"Fopf": "ð”½",
	"fopf": "ð•—",
	"forall": "âˆ€",
	"ForAll": "âˆ€",
	"fork": "â‹”",
	"forkv": "â«™",
	"Fouriertrf": "â„±",
	"fpartint": "â¨",
	"frac12": "Â½",
	"frac13": "â…“",
	"frac14": "Â¼",
	"frac15": "â…•",
	"frac16": "â…™",
	"frac18": "â…›",
	"frac23": "â…”",
	"frac25": "â…–",
	"frac34": "Â¾",
	"frac35": "â…—",
	"frac38": "â…œ",
	"frac45": "â…˜",
	"frac56": "â…š",
	"frac58": "â…",
	"frac78": "â…ž",
	"frasl": "â„",
	"frown": "âŒ¢",
	"fscr": "ð’»",
	"Fscr": "â„±",
	"gacute": "Çµ",
	"Gamma": "Î“",
	"gamma": "Î³",
	"Gammad": "Ïœ",
	"gammad": "Ï",
	"gap": "âª†",
	"Gbreve": "Äž",
	"gbreve": "ÄŸ",
	"Gcedil": "Ä¢",
	"Gcirc": "Äœ",
	"gcirc": "Ä",
	"Gcy": "Ð“",
	"gcy": "Ð³",
	"Gdot": "Ä ",
	"gdot": "Ä¡",
	"ge": "â‰¥",
	"gE": "â‰§",
	"gEl": "âªŒ",
	"gel": "â‹›",
	"geq": "â‰¥",
	"geqq": "â‰§",
	"geqslant": "â©¾",
	"gescc": "âª©",
	"ges": "â©¾",
	"gesdot": "âª€",
	"gesdoto": "âª‚",
	"gesdotol": "âª„",
	"gesl": "â‹›ï¸€",
	"gesles": "âª”",
	"Gfr": "ð”Š",
	"gfr": "ð”¤",
	"gg": "â‰«",
	"Gg": "â‹™",
	"ggg": "â‹™",
	"gimel": "â„·",
	"GJcy": "Ðƒ",
	"gjcy": "Ñ“",
	"gla": "âª¥",
	"gl": "â‰·",
	"glE": "âª’",
	"glj": "âª¤",
	"gnap": "âªŠ",
	"gnapprox": "âªŠ",
	"gne": "âªˆ",
	"gnE": "â‰©",
	"gneq": "âªˆ",
	"gneqq": "â‰©",
	"gnsim": "â‹§",
	"Gopf": "ð”¾",
	"gopf": "ð•˜",
	"grave": "`",
	"GreaterEqual": "â‰¥",
	"GreaterEqualLess": "â‹›",
	"GreaterFullEqual": "â‰§",
	"GreaterGreater": "âª¢",
	"GreaterLess": "â‰·",
	"GreaterSlantEqual": "â©¾",
	"GreaterTilde": "â‰³",
	"Gscr": "ð’¢",
	"gscr": "â„Š",
	"gsim": "â‰³",
	"gsime": "âªŽ",
	"gsiml": "âª",
	"gtcc": "âª§",
	"gtcir": "â©º",
	"gt": ">",
	"GT": ">",
	"Gt": "â‰«",
	"gtdot": "â‹—",
	"gtlPar": "â¦•",
	"gtquest": "â©¼",
	"gtrapprox": "âª†",
	"gtrarr": "â¥¸",
	"gtrdot": "â‹—",
	"gtreqless": "â‹›",
	"gtreqqless": "âªŒ",
	"gtrless": "â‰·",
	"gtrsim": "â‰³",
	"gvertneqq": "â‰©ï¸€",
	"gvnE": "â‰©ï¸€",
	"Hacek": "Ë‡",
	"hairsp": "â€Š",
	"half": "Â½",
	"hamilt": "â„‹",
	"HARDcy": "Ðª",
	"hardcy": "ÑŠ",
	"harrcir": "â¥ˆ",
	"harr": "â†”",
	"hArr": "â‡”",
	"harrw": "â†­",
	"Hat": "^",
	"hbar": "â„",
	"Hcirc": "Ä¤",
	"hcirc": "Ä¥",
	"hearts": "â™¥",
	"heartsuit": "â™¥",
	"hellip": "â€¦",
	"hercon": "âŠ¹",
	"hfr": "ð”¥",
	"Hfr": "â„Œ",
	"HilbertSpace": "â„‹",
	"hksearow": "â¤¥",
	"hkswarow": "â¤¦",
	"hoarr": "â‡¿",
	"homtht": "âˆ»",
	"hookleftarrow": "â†©",
	"hookrightarrow": "â†ª",
	"hopf": "ð•™",
	"Hopf": "â„",
	"horbar": "â€•",
	"HorizontalLine": "â”€",
	"hscr": "ð’½",
	"Hscr": "â„‹",
	"hslash": "â„",
	"Hstrok": "Ä¦",
	"hstrok": "Ä§",
	"HumpDownHump": "â‰Ž",
	"HumpEqual": "â‰",
	"hybull": "âƒ",
	"hyphen": "â€",
	"Iacute": "Ã",
	"iacute": "Ã­",
	"ic": "â£",
	"Icirc": "ÃŽ",
	"icirc": "Ã®",
	"Icy": "Ð˜",
	"icy": "Ð¸",
	"Idot": "Ä°",
	"IEcy": "Ð•",
	"iecy": "Ðµ",
	"iexcl": "Â¡",
	"iff": "â‡”",
	"ifr": "ð”¦",
	"Ifr": "â„‘",
	"Igrave": "ÃŒ",
	"igrave": "Ã¬",
	"ii": "â…ˆ",
	"iiiint": "â¨Œ",
	"iiint": "âˆ­",
	"iinfin": "â§œ",
	"iiota": "â„©",
	"IJlig": "Ä²",
	"ijlig": "Ä³",
	"Imacr": "Äª",
	"imacr": "Ä«",
	"image": "â„‘",
	"ImaginaryI": "â…ˆ",
	"imagline": "â„",
	"imagpart": "â„‘",
	"imath": "Ä±",
	"Im": "â„‘",
	"imof": "âŠ·",
	"imped": "Æµ",
	"Implies": "â‡’",
	"incare": "â„…",
	"in": "âˆˆ",
	"infin": "âˆž",
	"infintie": "â§",
	"inodot": "Ä±",
	"intcal": "âŠº",
	"int": "âˆ«",
	"Int": "âˆ¬",
	"integers": "â„¤",
	"Integral": "âˆ«",
	"intercal": "âŠº",
	"Intersection": "â‹‚",
	"intlarhk": "â¨—",
	"intprod": "â¨¼",
	"InvisibleComma": "â£",
	"InvisibleTimes": "â¢",
	"IOcy": "Ð",
	"iocy": "Ñ‘",
	"Iogon": "Ä®",
	"iogon": "Ä¯",
	"Iopf": "ð•€",
	"iopf": "ð•š",
	"Iota": "Î™",
	"iota": "Î¹",
	"iprod": "â¨¼",
	"iquest": "Â¿",
	"iscr": "ð’¾",
	"Iscr": "â„",
	"isin": "âˆˆ",
	"isindot": "â‹µ",
	"isinE": "â‹¹",
	"isins": "â‹´",
	"isinsv": "â‹³",
	"isinv": "âˆˆ",
	"it": "â¢",
	"Itilde": "Ä¨",
	"itilde": "Ä©",
	"Iukcy": "Ð†",
	"iukcy": "Ñ–",
	"Iuml": "Ã",
	"iuml": "Ã¯",
	"Jcirc": "Ä´",
	"jcirc": "Äµ",
	"Jcy": "Ð™",
	"jcy": "Ð¹",
	"Jfr": "ð”",
	"jfr": "ð”§",
	"jmath": "È·",
	"Jopf": "ð•",
	"jopf": "ð•›",
	"Jscr": "ð’¥",
	"jscr": "ð’¿",
	"Jsercy": "Ðˆ",
	"jsercy": "Ñ˜",
	"Jukcy": "Ð„",
	"jukcy": "Ñ”",
	"Kappa": "Îš",
	"kappa": "Îº",
	"kappav": "Ï°",
	"Kcedil": "Ä¶",
	"kcedil": "Ä·",
	"Kcy": "Ðš",
	"kcy": "Ðº",
	"Kfr": "ð”Ž",
	"kfr": "ð”¨",
	"kgreen": "Ä¸",
	"KHcy": "Ð¥",
	"khcy": "Ñ…",
	"KJcy": "ÐŒ",
	"kjcy": "Ñœ",
	"Kopf": "ð•‚",
	"kopf": "ð•œ",
	"Kscr": "ð’¦",
	"kscr": "ð“€",
	"lAarr": "â‡š",
	"Lacute": "Ä¹",
	"lacute": "Äº",
	"laemptyv": "â¦´",
	"lagran": "â„’",
	"Lambda": "Î›",
	"lambda": "Î»",
	"lang": "âŸ¨",
	"Lang": "âŸª",
	"langd": "â¦‘",
	"langle": "âŸ¨",
	"lap": "âª…",
	"Laplacetrf": "â„’",
	"laquo": "Â«",
	"larrb": "â‡¤",
	"larrbfs": "â¤Ÿ",
	"larr": "â†",
	"Larr": "â†ž",
	"lArr": "â‡",
	"larrfs": "â¤",
	"larrhk": "â†©",
	"larrlp": "â†«",
	"larrpl": "â¤¹",
	"larrsim": "â¥³",
	"larrtl": "â†¢",
	"latail": "â¤™",
	"lAtail": "â¤›",
	"lat": "âª«",
	"late": "âª­",
	"lates": "âª­ï¸€",
	"lbarr": "â¤Œ",
	"lBarr": "â¤Ž",
	"lbbrk": "â²",
	"lbrace": "{",
	"lbrack": "[",
	"lbrke": "â¦‹",
	"lbrksld": "â¦",
	"lbrkslu": "â¦",
	"Lcaron": "Ä½",
	"lcaron": "Ä¾",
	"Lcedil": "Ä»",
	"lcedil": "Ä¼",
	"lceil": "âŒˆ",
	"lcub": "{",
	"Lcy": "Ð›",
	"lcy": "Ð»",
	"ldca": "â¤¶",
	"ldquo": "â€œ",
	"ldquor": "â€ž",
	"ldrdhar": "â¥§",
	"ldrushar": "â¥‹",
	"ldsh": "â†²",
	"le": "â‰¤",
	"lE": "â‰¦",
	"LeftAngleBracket": "âŸ¨",
	"LeftArrowBar": "â‡¤",
	"leftarrow": "â†",
	"LeftArrow": "â†",
	"Leftarrow": "â‡",
	"LeftArrowRightArrow": "â‡†",
	"leftarrowtail": "â†¢",
	"LeftCeiling": "âŒˆ",
	"LeftDoubleBracket": "âŸ¦",
	"LeftDownTeeVector": "â¥¡",
	"LeftDownVectorBar": "â¥™",
	"LeftDownVector": "â‡ƒ",
	"LeftFloor": "âŒŠ",
	"leftharpoondown": "â†½",
	"leftharpoonup": "â†¼",
	"leftleftarrows": "â‡‡",
	"leftrightarrow": "â†”",
	"LeftRightArrow": "â†”",
	"Leftrightarrow": "â‡”",
	"leftrightarrows": "â‡†",
	"leftrightharpoons": "â‡‹",
	"leftrightsquigarrow": "â†­",
	"LeftRightVector": "â¥Ž",
	"LeftTeeArrow": "â†¤",
	"LeftTee": "âŠ£",
	"LeftTeeVector": "â¥š",
	"leftthreetimes": "â‹‹",
	"LeftTriangleBar": "â§",
	"LeftTriangle": "âŠ²",
	"LeftTriangleEqual": "âŠ´",
	"LeftUpDownVector": "â¥‘",
	"LeftUpTeeVector": "â¥ ",
	"LeftUpVectorBar": "â¥˜",
	"LeftUpVector": "â†¿",
	"LeftVectorBar": "â¥’",
	"LeftVector": "â†¼",
	"lEg": "âª‹",
	"leg": "â‹š",
	"leq": "â‰¤",
	"leqq": "â‰¦",
	"leqslant": "â©½",
	"lescc": "âª¨",
	"les": "â©½",
	"lesdot": "â©¿",
	"lesdoto": "âª",
	"lesdotor": "âªƒ",
	"lesg": "â‹šï¸€",
	"lesges": "âª“",
	"lessapprox": "âª…",
	"lessdot": "â‹–",
	"lesseqgtr": "â‹š",
	"lesseqqgtr": "âª‹",
	"LessEqualGreater": "â‹š",
	"LessFullEqual": "â‰¦",
	"LessGreater": "â‰¶",
	"lessgtr": "â‰¶",
	"LessLess": "âª¡",
	"lesssim": "â‰²",
	"LessSlantEqual": "â©½",
	"LessTilde": "â‰²",
	"lfisht": "â¥¼",
	"lfloor": "âŒŠ",
	"Lfr": "ð”",
	"lfr": "ð”©",
	"lg": "â‰¶",
	"lgE": "âª‘",
	"lHar": "â¥¢",
	"lhard": "â†½",
	"lharu": "â†¼",
	"lharul": "â¥ª",
	"lhblk": "â–„",
	"LJcy": "Ð‰",
	"ljcy": "Ñ™",
	"llarr": "â‡‡",
	"ll": "â‰ª",
	"Ll": "â‹˜",
	"llcorner": "âŒž",
	"Lleftarrow": "â‡š",
	"llhard": "â¥«",
	"lltri": "â—º",
	"Lmidot": "Ä¿",
	"lmidot": "Å€",
	"lmoustache": "âŽ°",
	"lmoust": "âŽ°",
	"lnap": "âª‰",
	"lnapprox": "âª‰",
	"lne": "âª‡",
	"lnE": "â‰¨",
	"lneq": "âª‡",
	"lneqq": "â‰¨",
	"lnsim": "â‹¦",
	"loang": "âŸ¬",
	"loarr": "â‡½",
	"lobrk": "âŸ¦",
	"longleftarrow": "âŸµ",
	"LongLeftArrow": "âŸµ",
	"Longleftarrow": "âŸ¸",
	"longleftrightarrow": "âŸ·",
	"LongLeftRightArrow": "âŸ·",
	"Longleftrightarrow": "âŸº",
	"longmapsto": "âŸ¼",
	"longrightarrow": "âŸ¶",
	"LongRightArrow": "âŸ¶",
	"Longrightarrow": "âŸ¹",
	"looparrowleft": "â†«",
	"looparrowright": "â†¬",
	"lopar": "â¦…",
	"Lopf": "ð•ƒ",
	"lopf": "ð•",
	"loplus": "â¨­",
	"lotimes": "â¨´",
	"lowast": "âˆ—",
	"lowbar": "_",
	"LowerLeftArrow": "â†™",
	"LowerRightArrow": "â†˜",
	"loz": "â—Š",
	"lozenge": "â—Š",
	"lozf": "â§«",
	"lpar": "(",
	"lparlt": "â¦“",
	"lrarr": "â‡†",
	"lrcorner": "âŒŸ",
	"lrhar": "â‡‹",
	"lrhard": "â¥­",
	"lrm": "â€Ž",
	"lrtri": "âŠ¿",
	"lsaquo": "â€¹",
	"lscr": "ð“",
	"Lscr": "â„’",
	"lsh": "â†°",
	"Lsh": "â†°",
	"lsim": "â‰²",
	"lsime": "âª",
	"lsimg": "âª",
	"lsqb": "[",
	"lsquo": "â€˜",
	"lsquor": "â€š",
	"Lstrok": "Å",
	"lstrok": "Å‚",
	"ltcc": "âª¦",
	"ltcir": "â©¹",
	"lt": "<",
	"LT": "<",
	"Lt": "â‰ª",
	"ltdot": "â‹–",
	"lthree": "â‹‹",
	"ltimes": "â‹‰",
	"ltlarr": "â¥¶",
	"ltquest": "â©»",
	"ltri": "â—ƒ",
	"ltrie": "âŠ´",
	"ltrif": "â—‚",
	"ltrPar": "â¦–",
	"lurdshar": "â¥Š",
	"luruhar": "â¥¦",
	"lvertneqq": "â‰¨ï¸€",
	"lvnE": "â‰¨ï¸€",
	"macr": "Â¯",
	"male": "â™‚",
	"malt": "âœ ",
	"maltese": "âœ ",
	"Map": "â¤…",
	"map": "â†¦",
	"mapsto": "â†¦",
	"mapstodown": "â†§",
	"mapstoleft": "â†¤",
	"mapstoup": "â†¥",
	"marker": "â–®",
	"mcomma": "â¨©",
	"Mcy": "Ðœ",
	"mcy": "Ð¼",
	"mdash": "â€”",
	"mDDot": "âˆº",
	"measuredangle": "âˆ¡",
	"MediumSpace": "âŸ",
	"Mellintrf": "â„³",
	"Mfr": "ð”",
	"mfr": "ð”ª",
	"mho": "â„§",
	"micro": "Âµ",
	"midast": "*",
	"midcir": "â«°",
	"mid": "âˆ£",
	"middot": "Â·",
	"minusb": "âŠŸ",
	"minus": "âˆ’",
	"minusd": "âˆ¸",
	"minusdu": "â¨ª",
	"MinusPlus": "âˆ“",
	"mlcp": "â«›",
	"mldr": "â€¦",
	"mnplus": "âˆ“",
	"models": "âŠ§",
	"Mopf": "ð•„",
	"mopf": "ð•ž",
	"mp": "âˆ“",
	"mscr": "ð“‚",
	"Mscr": "â„³",
	"mstpos": "âˆ¾",
	"Mu": "Îœ",
	"mu": "Î¼",
	"multimap": "âŠ¸",
	"mumap": "âŠ¸",
	"nabla": "âˆ‡",
	"Nacute": "Åƒ",
	"nacute": "Å„",
	"nang": "âˆ âƒ’",
	"nap": "â‰‰",
	"napE": "â©°Ì¸",
	"napid": "â‰‹Ì¸",
	"napos": "Å‰",
	"napprox": "â‰‰",
	"natural": "â™®",
	"naturals": "â„•",
	"natur": "â™®",
	"nbsp": "Â ",
	"nbump": "â‰ŽÌ¸",
	"nbumpe": "â‰Ì¸",
	"ncap": "â©ƒ",
	"Ncaron": "Å‡",
	"ncaron": "Åˆ",
	"Ncedil": "Å…",
	"ncedil": "Å†",
	"ncong": "â‰‡",
	"ncongdot": "â©­Ì¸",
	"ncup": "â©‚",
	"Ncy": "Ð",
	"ncy": "Ð½",
	"ndash": "â€“",
	"nearhk": "â¤¤",
	"nearr": "â†—",
	"neArr": "â‡—",
	"nearrow": "â†—",
	"ne": "â‰ ",
	"nedot": "â‰Ì¸",
	"NegativeMediumSpace": "â€‹",
	"NegativeThickSpace": "â€‹",
	"NegativeThinSpace": "â€‹",
	"NegativeVeryThinSpace": "â€‹",
	"nequiv": "â‰¢",
	"nesear": "â¤¨",
	"nesim": "â‰‚Ì¸",
	"NestedGreaterGreater": "â‰«",
	"NestedLessLess": "â‰ª",
	"NewLine": "\n",
	"nexist": "âˆ„",
	"nexists": "âˆ„",
	"Nfr": "ð”‘",
	"nfr": "ð”«",
	"ngE": "â‰§Ì¸",
	"nge": "â‰±",
	"ngeq": "â‰±",
	"ngeqq": "â‰§Ì¸",
	"ngeqslant": "â©¾Ì¸",
	"nges": "â©¾Ì¸",
	"nGg": "â‹™Ì¸",
	"ngsim": "â‰µ",
	"nGt": "â‰«âƒ’",
	"ngt": "â‰¯",
	"ngtr": "â‰¯",
	"nGtv": "â‰«Ì¸",
	"nharr": "â†®",
	"nhArr": "â‡Ž",
	"nhpar": "â«²",
	"ni": "âˆ‹",
	"nis": "â‹¼",
	"nisd": "â‹º",
	"niv": "âˆ‹",
	"NJcy": "ÐŠ",
	"njcy": "Ñš",
	"nlarr": "â†š",
	"nlArr": "â‡",
	"nldr": "â€¥",
	"nlE": "â‰¦Ì¸",
	"nle": "â‰°",
	"nleftarrow": "â†š",
	"nLeftarrow": "â‡",
	"nleftrightarrow": "â†®",
	"nLeftrightarrow": "â‡Ž",
	"nleq": "â‰°",
	"nleqq": "â‰¦Ì¸",
	"nleqslant": "â©½Ì¸",
	"nles": "â©½Ì¸",
	"nless": "â‰®",
	"nLl": "â‹˜Ì¸",
	"nlsim": "â‰´",
	"nLt": "â‰ªâƒ’",
	"nlt": "â‰®",
	"nltri": "â‹ª",
	"nltrie": "â‹¬",
	"nLtv": "â‰ªÌ¸",
	"nmid": "âˆ¤",
	"NoBreak": "â ",
	"NonBreakingSpace": "Â ",
	"nopf": "ð•Ÿ",
	"Nopf": "â„•",
	"Not": "â«¬",
	"not": "Â¬",
	"NotCongruent": "â‰¢",
	"NotCupCap": "â‰­",
	"NotDoubleVerticalBar": "âˆ¦",
	"NotElement": "âˆ‰",
	"NotEqual": "â‰ ",
	"NotEqualTilde": "â‰‚Ì¸",
	"NotExists": "âˆ„",
	"NotGreater": "â‰¯",
	"NotGreaterEqual": "â‰±",
	"NotGreaterFullEqual": "â‰§Ì¸",
	"NotGreaterGreater": "â‰«Ì¸",
	"NotGreaterLess": "â‰¹",
	"NotGreaterSlantEqual": "â©¾Ì¸",
	"NotGreaterTilde": "â‰µ",
	"NotHumpDownHump": "â‰ŽÌ¸",
	"NotHumpEqual": "â‰Ì¸",
	"notin": "âˆ‰",
	"notindot": "â‹µÌ¸",
	"notinE": "â‹¹Ì¸",
	"notinva": "âˆ‰",
	"notinvb": "â‹·",
	"notinvc": "â‹¶",
	"NotLeftTriangleBar": "â§Ì¸",
	"NotLeftTriangle": "â‹ª",
	"NotLeftTriangleEqual": "â‹¬",
	"NotLess": "â‰®",
	"NotLessEqual": "â‰°",
	"NotLessGreater": "â‰¸",
	"NotLessLess": "â‰ªÌ¸",
	"NotLessSlantEqual": "â©½Ì¸",
	"NotLessTilde": "â‰´",
	"NotNestedGreaterGreater": "âª¢Ì¸",
	"NotNestedLessLess": "âª¡Ì¸",
	"notni": "âˆŒ",
	"notniva": "âˆŒ",
	"notnivb": "â‹¾",
	"notnivc": "â‹½",
	"NotPrecedes": "âŠ€",
	"NotPrecedesEqual": "âª¯Ì¸",
	"NotPrecedesSlantEqual": "â‹ ",
	"NotReverseElement": "âˆŒ",
	"NotRightTriangleBar": "â§Ì¸",
	"NotRightTriangle": "â‹«",
	"NotRightTriangleEqual": "â‹­",
	"NotSquareSubset": "âŠÌ¸",
	"NotSquareSubsetEqual": "â‹¢",
	"NotSquareSuperset": "âŠÌ¸",
	"NotSquareSupersetEqual": "â‹£",
	"NotSubset": "âŠ‚âƒ’",
	"NotSubsetEqual": "âŠˆ",
	"NotSucceeds": "âŠ",
	"NotSucceedsEqual": "âª°Ì¸",
	"NotSucceedsSlantEqual": "â‹¡",
	"NotSucceedsTilde": "â‰¿Ì¸",
	"NotSuperset": "âŠƒâƒ’",
	"NotSupersetEqual": "âŠ‰",
	"NotTilde": "â‰",
	"NotTildeEqual": "â‰„",
	"NotTildeFullEqual": "â‰‡",
	"NotTildeTilde": "â‰‰",
	"NotVerticalBar": "âˆ¤",
	"nparallel": "âˆ¦",
	"npar": "âˆ¦",
	"nparsl": "â«½âƒ¥",
	"npart": "âˆ‚Ì¸",
	"npolint": "â¨”",
	"npr": "âŠ€",
	"nprcue": "â‹ ",
	"nprec": "âŠ€",
	"npreceq": "âª¯Ì¸",
	"npre": "âª¯Ì¸",
	"nrarrc": "â¤³Ì¸",
	"nrarr": "â†›",
	"nrArr": "â‡",
	"nrarrw": "â†Ì¸",
	"nrightarrow": "â†›",
	"nRightarrow": "â‡",
	"nrtri": "â‹«",
	"nrtrie": "â‹­",
	"nsc": "âŠ",
	"nsccue": "â‹¡",
	"nsce": "âª°Ì¸",
	"Nscr": "ð’©",
	"nscr": "ð“ƒ",
	"nshortmid": "âˆ¤",
	"nshortparallel": "âˆ¦",
	"nsim": "â‰",
	"nsime": "â‰„",
	"nsimeq": "â‰„",
	"nsmid": "âˆ¤",
	"nspar": "âˆ¦",
	"nsqsube": "â‹¢",
	"nsqsupe": "â‹£",
	"nsub": "âŠ„",
	"nsubE": "â«…Ì¸",
	"nsube": "âŠˆ",
	"nsubset": "âŠ‚âƒ’",
	"nsubseteq": "âŠˆ",
	"nsubseteqq": "â«…Ì¸",
	"nsucc": "âŠ",
	"nsucceq": "âª°Ì¸",
	"nsup": "âŠ…",
	"nsupE": "â«†Ì¸",
	"nsupe": "âŠ‰",
	"nsupset": "âŠƒâƒ’",
	"nsupseteq": "âŠ‰",
	"nsupseteqq": "â«†Ì¸",
	"ntgl": "â‰¹",
	"Ntilde": "Ã‘",
	"ntilde": "Ã±",
	"ntlg": "â‰¸",
	"ntriangleleft": "â‹ª",
	"ntrianglelefteq": "â‹¬",
	"ntriangleright": "â‹«",
	"ntrianglerighteq": "â‹­",
	"Nu": "Î",
	"nu": "Î½",
	"num": "#",
	"numero": "â„–",
	"numsp": "â€‡",
	"nvap": "â‰âƒ’",
	"nvdash": "âŠ¬",
	"nvDash": "âŠ­",
	"nVdash": "âŠ®",
	"nVDash": "âŠ¯",
	"nvge": "â‰¥âƒ’",
	"nvgt": ">âƒ’",
	"nvHarr": "â¤„",
	"nvinfin": "â§ž",
	"nvlArr": "â¤‚",
	"nvle": "â‰¤âƒ’",
	"nvlt": "<âƒ’",
	"nvltrie": "âŠ´âƒ’",
	"nvrArr": "â¤ƒ",
	"nvrtrie": "âŠµâƒ’",
	"nvsim": "âˆ¼âƒ’",
	"nwarhk": "â¤£",
	"nwarr": "â†–",
	"nwArr": "â‡–",
	"nwarrow": "â†–",
	"nwnear": "â¤§",
	"Oacute": "Ã“",
	"oacute": "Ã³",
	"oast": "âŠ›",
	"Ocirc": "Ã”",
	"ocirc": "Ã´",
	"ocir": "âŠš",
	"Ocy": "Ðž",
	"ocy": "Ð¾",
	"odash": "âŠ",
	"Odblac": "Å",
	"odblac": "Å‘",
	"odiv": "â¨¸",
	"odot": "âŠ™",
	"odsold": "â¦¼",
	"OElig": "Å’",
	"oelig": "Å“",
	"ofcir": "â¦¿",
	"Ofr": "ð”’",
	"ofr": "ð”¬",
	"ogon": "Ë›",
	"Ograve": "Ã’",
	"ograve": "Ã²",
	"ogt": "â§",
	"ohbar": "â¦µ",
	"ohm": "Î©",
	"oint": "âˆ®",
	"olarr": "â†º",
	"olcir": "â¦¾",
	"olcross": "â¦»",
	"oline": "â€¾",
	"olt": "â§€",
	"Omacr": "ÅŒ",
	"omacr": "Å",
	"Omega": "Î©",
	"omega": "Ï‰",
	"Omicron": "ÎŸ",
	"omicron": "Î¿",
	"omid": "â¦¶",
	"ominus": "âŠ–",
	"Oopf": "ð•†",
	"oopf": "ð• ",
	"opar": "â¦·",
	"OpenCurlyDoubleQuote": "â€œ",
	"OpenCurlyQuote": "â€˜",
	"operp": "â¦¹",
	"oplus": "âŠ•",
	"orarr": "â†»",
	"Or": "â©”",
	"or": "âˆ¨",
	"ord": "â©",
	"order": "â„´",
	"orderof": "â„´",
	"ordf": "Âª",
	"ordm": "Âº",
	"origof": "âŠ¶",
	"oror": "â©–",
	"orslope": "â©—",
	"orv": "â©›",
	"oS": "â“ˆ",
	"Oscr": "ð’ª",
	"oscr": "â„´",
	"Oslash": "Ã˜",
	"oslash": "Ã¸",
	"osol": "âŠ˜",
	"Otilde": "Ã•",
	"otilde": "Ãµ",
	"otimesas": "â¨¶",
	"Otimes": "â¨·",
	"otimes": "âŠ—",
	"Ouml": "Ã–",
	"ouml": "Ã¶",
	"ovbar": "âŒ½",
	"OverBar": "â€¾",
	"OverBrace": "âž",
	"OverBracket": "âŽ´",
	"OverParenthesis": "âœ",
	"para": "Â¶",
	"parallel": "âˆ¥",
	"par": "âˆ¥",
	"parsim": "â«³",
	"parsl": "â«½",
	"part": "âˆ‚",
	"PartialD": "âˆ‚",
	"Pcy": "ÐŸ",
	"pcy": "Ð¿",
	"percnt": "%",
	"period": ".",
	"permil": "â€°",
	"perp": "âŠ¥",
	"pertenk": "â€±",
	"Pfr": "ð”“",
	"pfr": "ð”­",
	"Phi": "Î¦",
	"phi": "Ï†",
	"phiv": "Ï•",
	"phmmat": "â„³",
	"phone": "â˜Ž",
	"Pi": "Î ",
	"pi": "Ï€",
	"pitchfork": "â‹”",
	"piv": "Ï–",
	"planck": "â„",
	"planckh": "â„Ž",
	"plankv": "â„",
	"plusacir": "â¨£",
	"plusb": "âŠž",
	"pluscir": "â¨¢",
	"plus": "+",
	"plusdo": "âˆ”",
	"plusdu": "â¨¥",
	"pluse": "â©²",
	"PlusMinus": "Â±",
	"plusmn": "Â±",
	"plussim": "â¨¦",
	"plustwo": "â¨§",
	"pm": "Â±",
	"Poincareplane": "â„Œ",
	"pointint": "â¨•",
	"popf": "ð•¡",
	"Popf": "â„™",
	"pound": "Â£",
	"prap": "âª·",
	"Pr": "âª»",
	"pr": "â‰º",
	"prcue": "â‰¼",
	"precapprox": "âª·",
	"prec": "â‰º",
	"preccurlyeq": "â‰¼",
	"Precedes": "â‰º",
	"PrecedesEqual": "âª¯",
	"PrecedesSlantEqual": "â‰¼",
	"PrecedesTilde": "â‰¾",
	"preceq": "âª¯",
	"precnapprox": "âª¹",
	"precneqq": "âªµ",
	"precnsim": "â‹¨",
	"pre": "âª¯",
	"prE": "âª³",
	"precsim": "â‰¾",
	"prime": "â€²",
	"Prime": "â€³",
	"primes": "â„™",
	"prnap": "âª¹",
	"prnE": "âªµ",
	"prnsim": "â‹¨",
	"prod": "âˆ",
	"Product": "âˆ",
	"profalar": "âŒ®",
	"profline": "âŒ’",
	"profsurf": "âŒ“",
	"prop": "âˆ",
	"Proportional": "âˆ",
	"Proportion": "âˆ·",
	"propto": "âˆ",
	"prsim": "â‰¾",
	"prurel": "âŠ°",
	"Pscr": "ð’«",
	"pscr": "ð“…",
	"Psi": "Î¨",
	"psi": "Ïˆ",
	"puncsp": "â€ˆ",
	"Qfr": "ð””",
	"qfr": "ð”®",
	"qint": "â¨Œ",
	"qopf": "ð•¢",
	"Qopf": "â„š",
	"qprime": "â—",
	"Qscr": "ð’¬",
	"qscr": "ð“†",
	"quaternions": "â„",
	"quatint": "â¨–",
	"quest": "?",
	"questeq": "â‰Ÿ",
	"quot": "\"",
	"QUOT": "\"",
	"rAarr": "â‡›",
	"race": "âˆ½Ì±",
	"Racute": "Å”",
	"racute": "Å•",
	"radic": "âˆš",
	"raemptyv": "â¦³",
	"rang": "âŸ©",
	"Rang": "âŸ«",
	"rangd": "â¦’",
	"range": "â¦¥",
	"rangle": "âŸ©",
	"raquo": "Â»",
	"rarrap": "â¥µ",
	"rarrb": "â‡¥",
	"rarrbfs": "â¤ ",
	"rarrc": "â¤³",
	"rarr": "â†’",
	"Rarr": "â† ",
	"rArr": "â‡’",
	"rarrfs": "â¤ž",
	"rarrhk": "â†ª",
	"rarrlp": "â†¬",
	"rarrpl": "â¥…",
	"rarrsim": "â¥´",
	"Rarrtl": "â¤–",
	"rarrtl": "â†£",
	"rarrw": "â†",
	"ratail": "â¤š",
	"rAtail": "â¤œ",
	"ratio": "âˆ¶",
	"rationals": "â„š",
	"rbarr": "â¤",
	"rBarr": "â¤",
	"RBarr": "â¤",
	"rbbrk": "â³",
	"rbrace": "}",
	"rbrack": "]",
	"rbrke": "â¦Œ",
	"rbrksld": "â¦Ž",
	"rbrkslu": "â¦",
	"Rcaron": "Å˜",
	"rcaron": "Å™",
	"Rcedil": "Å–",
	"rcedil": "Å—",
	"rceil": "âŒ‰",
	"rcub": "}",
	"Rcy": "Ð ",
	"rcy": "Ñ€",
	"rdca": "â¤·",
	"rdldhar": "â¥©",
	"rdquo": "â€",
	"rdquor": "â€",
	"rdsh": "â†³",
	"real": "â„œ",
	"realine": "â„›",
	"realpart": "â„œ",
	"reals": "â„",
	"Re": "â„œ",
	"rect": "â–­",
	"reg": "Â®",
	"REG": "Â®",
	"ReverseElement": "âˆ‹",
	"ReverseEquilibrium": "â‡‹",
	"ReverseUpEquilibrium": "â¥¯",
	"rfisht": "â¥½",
	"rfloor": "âŒ‹",
	"rfr": "ð”¯",
	"Rfr": "â„œ",
	"rHar": "â¥¤",
	"rhard": "â‡",
	"rharu": "â‡€",
	"rharul": "â¥¬",
	"Rho": "Î¡",
	"rho": "Ï",
	"rhov": "Ï±",
	"RightAngleBracket": "âŸ©",
	"RightArrowBar": "â‡¥",
	"rightarrow": "â†’",
	"RightArrow": "â†’",
	"Rightarrow": "â‡’",
	"RightArrowLeftArrow": "â‡„",
	"rightarrowtail": "â†£",
	"RightCeiling": "âŒ‰",
	"RightDoubleBracket": "âŸ§",
	"RightDownTeeVector": "â¥",
	"RightDownVectorBar": "â¥•",
	"RightDownVector": "â‡‚",
	"RightFloor": "âŒ‹",
	"rightharpoondown": "â‡",
	"rightharpoonup": "â‡€",
	"rightleftarrows": "â‡„",
	"rightleftharpoons": "â‡Œ",
	"rightrightarrows": "â‡‰",
	"rightsquigarrow": "â†",
	"RightTeeArrow": "â†¦",
	"RightTee": "âŠ¢",
	"RightTeeVector": "â¥›",
	"rightthreetimes": "â‹Œ",
	"RightTriangleBar": "â§",
	"RightTriangle": "âŠ³",
	"RightTriangleEqual": "âŠµ",
	"RightUpDownVector": "â¥",
	"RightUpTeeVector": "â¥œ",
	"RightUpVectorBar": "â¥”",
	"RightUpVector": "â†¾",
	"RightVectorBar": "â¥“",
	"RightVector": "â‡€",
	"ring": "Ëš",
	"risingdotseq": "â‰“",
	"rlarr": "â‡„",
	"rlhar": "â‡Œ",
	"rlm": "â€",
	"rmoustache": "âŽ±",
	"rmoust": "âŽ±",
	"rnmid": "â«®",
	"roang": "âŸ­",
	"roarr": "â‡¾",
	"robrk": "âŸ§",
	"ropar": "â¦†",
	"ropf": "ð•£",
	"Ropf": "â„",
	"roplus": "â¨®",
	"rotimes": "â¨µ",
	"RoundImplies": "â¥°",
	"rpar": ")",
	"rpargt": "â¦”",
	"rppolint": "â¨’",
	"rrarr": "â‡‰",
	"Rrightarrow": "â‡›",
	"rsaquo": "â€º",
	"rscr": "ð“‡",
	"Rscr": "â„›",
	"rsh": "â†±",
	"Rsh": "â†±",
	"rsqb": "]",
	"rsquo": "â€™",
	"rsquor": "â€™",
	"rthree": "â‹Œ",
	"rtimes": "â‹Š",
	"rtri": "â–¹",
	"rtrie": "âŠµ",
	"rtrif": "â–¸",
	"rtriltri": "â§Ž",
	"RuleDelayed": "â§´",
	"ruluhar": "â¥¨",
	"rx": "â„ž",
	"Sacute": "Åš",
	"sacute": "Å›",
	"sbquo": "â€š",
	"scap": "âª¸",
	"Scaron": "Å ",
	"scaron": "Å¡",
	"Sc": "âª¼",
	"sc": "â‰»",
	"sccue": "â‰½",
	"sce": "âª°",
	"scE": "âª´",
	"Scedil": "Åž",
	"scedil": "ÅŸ",
	"Scirc": "Åœ",
	"scirc": "Å",
	"scnap": "âªº",
	"scnE": "âª¶",
	"scnsim": "â‹©",
	"scpolint": "â¨“",
	"scsim": "â‰¿",
	"Scy": "Ð¡",
	"scy": "Ñ",
	"sdotb": "âŠ¡",
	"sdot": "â‹…",
	"sdote": "â©¦",
	"searhk": "â¤¥",
	"searr": "â†˜",
	"seArr": "â‡˜",
	"searrow": "â†˜",
	"sect": "Â§",
	"semi": ";",
	"seswar": "â¤©",
	"setminus": "âˆ–",
	"setmn": "âˆ–",
	"sext": "âœ¶",
	"Sfr": "ð”–",
	"sfr": "ð”°",
	"sfrown": "âŒ¢",
	"sharp": "â™¯",
	"SHCHcy": "Ð©",
	"shchcy": "Ñ‰",
	"SHcy": "Ð¨",
	"shcy": "Ñˆ",
	"ShortDownArrow": "â†“",
	"ShortLeftArrow": "â†",
	"shortmid": "âˆ£",
	"shortparallel": "âˆ¥",
	"ShortRightArrow": "â†’",
	"ShortUpArrow": "â†‘",
	"shy": "Â­",
	"Sigma": "Î£",
	"sigma": "Ïƒ",
	"sigmaf": "Ï‚",
	"sigmav": "Ï‚",
	"sim": "âˆ¼",
	"simdot": "â©ª",
	"sime": "â‰ƒ",
	"simeq": "â‰ƒ",
	"simg": "âªž",
	"simgE": "âª ",
	"siml": "âª",
	"simlE": "âªŸ",
	"simne": "â‰†",
	"simplus": "â¨¤",
	"simrarr": "â¥²",
	"slarr": "â†",
	"SmallCircle": "âˆ˜",
	"smallsetminus": "âˆ–",
	"smashp": "â¨³",
	"smeparsl": "â§¤",
	"smid": "âˆ£",
	"smile": "âŒ£",
	"smt": "âªª",
	"smte": "âª¬",
	"smtes": "âª¬ï¸€",
	"SOFTcy": "Ð¬",
	"softcy": "ÑŒ",
	"solbar": "âŒ¿",
	"solb": "â§„",
	"sol": "/",
	"Sopf": "ð•Š",
	"sopf": "ð•¤",
	"spades": "â™ ",
	"spadesuit": "â™ ",
	"spar": "âˆ¥",
	"sqcap": "âŠ“",
	"sqcaps": "âŠ“ï¸€",
	"sqcup": "âŠ”",
	"sqcups": "âŠ”ï¸€",
	"Sqrt": "âˆš",
	"sqsub": "âŠ",
	"sqsube": "âŠ‘",
	"sqsubset": "âŠ",
	"sqsubseteq": "âŠ‘",
	"sqsup": "âŠ",
	"sqsupe": "âŠ’",
	"sqsupset": "âŠ",
	"sqsupseteq": "âŠ’",
	"square": "â–¡",
	"Square": "â–¡",
	"SquareIntersection": "âŠ“",
	"SquareSubset": "âŠ",
	"SquareSubsetEqual": "âŠ‘",
	"SquareSuperset": "âŠ",
	"SquareSupersetEqual": "âŠ’",
	"SquareUnion": "âŠ”",
	"squarf": "â–ª",
	"squ": "â–¡",
	"squf": "â–ª",
	"srarr": "â†’",
	"Sscr": "ð’®",
	"sscr": "ð“ˆ",
	"ssetmn": "âˆ–",
	"ssmile": "âŒ£",
	"sstarf": "â‹†",
	"Star": "â‹†",
	"star": "â˜†",
	"starf": "â˜…",
	"straightepsilon": "Ïµ",
	"straightphi": "Ï•",
	"strns": "Â¯",
	"sub": "âŠ‚",
	"Sub": "â‹",
	"subdot": "âª½",
	"subE": "â«…",
	"sube": "âŠ†",
	"subedot": "â«ƒ",
	"submult": "â«",
	"subnE": "â«‹",
	"subne": "âŠŠ",
	"subplus": "âª¿",
	"subrarr": "â¥¹",
	"subset": "âŠ‚",
	"Subset": "â‹",
	"subseteq": "âŠ†",
	"subseteqq": "â«…",
	"SubsetEqual": "âŠ†",
	"subsetneq": "âŠŠ",
	"subsetneqq": "â«‹",
	"subsim": "â«‡",
	"subsub": "â«•",
	"subsup": "â«“",
	"succapprox": "âª¸",
	"succ": "â‰»",
	"succcurlyeq": "â‰½",
	"Succeeds": "â‰»",
	"SucceedsEqual": "âª°",
	"SucceedsSlantEqual": "â‰½",
	"SucceedsTilde": "â‰¿",
	"succeq": "âª°",
	"succnapprox": "âªº",
	"succneqq": "âª¶",
	"succnsim": "â‹©",
	"succsim": "â‰¿",
	"SuchThat": "âˆ‹",
	"sum": "âˆ‘",
	"Sum": "âˆ‘",
	"sung": "â™ª",
	"sup1": "Â¹",
	"sup2": "Â²",
	"sup3": "Â³",
	"sup": "âŠƒ",
	"Sup": "â‹‘",
	"supdot": "âª¾",
	"supdsub": "â«˜",
	"supE": "â«†",
	"supe": "âŠ‡",
	"supedot": "â«„",
	"Superset": "âŠƒ",
	"SupersetEqual": "âŠ‡",
	"suphsol": "âŸ‰",
	"suphsub": "â«—",
	"suplarr": "â¥»",
	"supmult": "â«‚",
	"supnE": "â«Œ",
	"supne": "âŠ‹",
	"supplus": "â«€",
	"supset": "âŠƒ",
	"Supset": "â‹‘",
	"supseteq": "âŠ‡",
	"supseteqq": "â«†",
	"supsetneq": "âŠ‹",
	"supsetneqq": "â«Œ",
	"supsim": "â«ˆ",
	"supsub": "â«”",
	"supsup": "â«–",
	"swarhk": "â¤¦",
	"swarr": "â†™",
	"swArr": "â‡™",
	"swarrow": "â†™",
	"swnwar": "â¤ª",
	"szlig": "ÃŸ",
	"Tab": "\t",
	"target": "âŒ–",
	"Tau": "Î¤",
	"tau": "Ï„",
	"tbrk": "âŽ´",
	"Tcaron": "Å¤",
	"tcaron": "Å¥",
	"Tcedil": "Å¢",
	"tcedil": "Å£",
	"Tcy": "Ð¢",
	"tcy": "Ñ‚",
	"tdot": "âƒ›",
	"telrec": "âŒ•",
	"Tfr": "ð”—",
	"tfr": "ð”±",
	"there4": "âˆ´",
	"therefore": "âˆ´",
	"Therefore": "âˆ´",
	"Theta": "Î˜",
	"theta": "Î¸",
	"thetasym": "Ï‘",
	"thetav": "Ï‘",
	"thickapprox": "â‰ˆ",
	"thicksim": "âˆ¼",
	"ThickSpace": "âŸâ€Š",
	"ThinSpace": "â€‰",
	"thinsp": "â€‰",
	"thkap": "â‰ˆ",
	"thksim": "âˆ¼",
	"THORN": "Ãž",
	"thorn": "Ã¾",
	"tilde": "Ëœ",
	"Tilde": "âˆ¼",
	"TildeEqual": "â‰ƒ",
	"TildeFullEqual": "â‰…",
	"TildeTilde": "â‰ˆ",
	"timesbar": "â¨±",
	"timesb": "âŠ ",
	"times": "Ã—",
	"timesd": "â¨°",
	"tint": "âˆ­",
	"toea": "â¤¨",
	"topbot": "âŒ¶",
	"topcir": "â«±",
	"top": "âŠ¤",
	"Topf": "ð•‹",
	"topf": "ð•¥",
	"topfork": "â«š",
	"tosa": "â¤©",
	"tprime": "â€´",
	"trade": "â„¢",
	"TRADE": "â„¢",
	"triangle": "â–µ",
	"triangledown": "â–¿",
	"triangleleft": "â—ƒ",
	"trianglelefteq": "âŠ´",
	"triangleq": "â‰œ",
	"triangleright": "â–¹",
	"trianglerighteq": "âŠµ",
	"tridot": "â—¬",
	"trie": "â‰œ",
	"triminus": "â¨º",
	"TripleDot": "âƒ›",
	"triplus": "â¨¹",
	"trisb": "â§",
	"tritime": "â¨»",
	"trpezium": "â¢",
	"Tscr": "ð’¯",
	"tscr": "ð“‰",
	"TScy": "Ð¦",
	"tscy": "Ñ†",
	"TSHcy": "Ð‹",
	"tshcy": "Ñ›",
	"Tstrok": "Å¦",
	"tstrok": "Å§",
	"twixt": "â‰¬",
	"twoheadleftarrow": "â†ž",
	"twoheadrightarrow": "â† ",
	"Uacute": "Ãš",
	"uacute": "Ãº",
	"uarr": "â†‘",
	"Uarr": "â†Ÿ",
	"uArr": "â‡‘",
	"Uarrocir": "â¥‰",
	"Ubrcy": "ÐŽ",
	"ubrcy": "Ñž",
	"Ubreve": "Å¬",
	"ubreve": "Å­",
	"Ucirc": "Ã›",
	"ucirc": "Ã»",
	"Ucy": "Ð£",
	"ucy": "Ñƒ",
	"udarr": "â‡…",
	"Udblac": "Å°",
	"udblac": "Å±",
	"udhar": "â¥®",
	"ufisht": "â¥¾",
	"Ufr": "ð”˜",
	"ufr": "ð”²",
	"Ugrave": "Ã™",
	"ugrave": "Ã¹",
	"uHar": "â¥£",
	"uharl": "â†¿",
	"uharr": "â†¾",
	"uhblk": "â–€",
	"ulcorn": "âŒœ",
	"ulcorner": "âŒœ",
	"ulcrop": "âŒ",
	"ultri": "â—¸",
	"Umacr": "Åª",
	"umacr": "Å«",
	"uml": "Â¨",
	"UnderBar": "_",
	"UnderBrace": "âŸ",
	"UnderBracket": "âŽµ",
	"UnderParenthesis": "â",
	"Union": "â‹ƒ",
	"UnionPlus": "âŠŽ",
	"Uogon": "Å²",
	"uogon": "Å³",
	"Uopf": "ð•Œ",
	"uopf": "ð•¦",
	"UpArrowBar": "â¤’",
	"uparrow": "â†‘",
	"UpArrow": "â†‘",
	"Uparrow": "â‡‘",
	"UpArrowDownArrow": "â‡…",
	"updownarrow": "â†•",
	"UpDownArrow": "â†•",
	"Updownarrow": "â‡•",
	"UpEquilibrium": "â¥®",
	"upharpoonleft": "â†¿",
	"upharpoonright": "â†¾",
	"uplus": "âŠŽ",
	"UpperLeftArrow": "â†–",
	"UpperRightArrow": "â†—",
	"upsi": "Ï…",
	"Upsi": "Ï’",
	"upsih": "Ï’",
	"Upsilon": "Î¥",
	"upsilon": "Ï…",
	"UpTeeArrow": "â†¥",
	"UpTee": "âŠ¥",
	"upuparrows": "â‡ˆ",
	"urcorn": "âŒ",
	"urcorner": "âŒ",
	"urcrop": "âŒŽ",
	"Uring": "Å®",
	"uring": "Å¯",
	"urtri": "â—¹",
	"Uscr": "ð’°",
	"uscr": "ð“Š",
	"utdot": "â‹°",
	"Utilde": "Å¨",
	"utilde": "Å©",
	"utri": "â–µ",
	"utrif": "â–´",
	"uuarr": "â‡ˆ",
	"Uuml": "Ãœ",
	"uuml": "Ã¼",
	"uwangle": "â¦§",
	"vangrt": "â¦œ",
	"varepsilon": "Ïµ",
	"varkappa": "Ï°",
	"varnothing": "âˆ…",
	"varphi": "Ï•",
	"varpi": "Ï–",
	"varpropto": "âˆ",
	"varr": "â†•",
	"vArr": "â‡•",
	"varrho": "Ï±",
	"varsigma": "Ï‚",
	"varsubsetneq": "âŠŠï¸€",
	"varsubsetneqq": "â«‹ï¸€",
	"varsupsetneq": "âŠ‹ï¸€",
	"varsupsetneqq": "â«Œï¸€",
	"vartheta": "Ï‘",
	"vartriangleleft": "âŠ²",
	"vartriangleright": "âŠ³",
	"vBar": "â«¨",
	"Vbar": "â««",
	"vBarv": "â«©",
	"Vcy": "Ð’",
	"vcy": "Ð²",
	"vdash": "âŠ¢",
	"vDash": "âŠ¨",
	"Vdash": "âŠ©",
	"VDash": "âŠ«",
	"Vdashl": "â«¦",
	"veebar": "âŠ»",
	"vee": "âˆ¨",
	"Vee": "â‹",
	"veeeq": "â‰š",
	"vellip": "â‹®",
	"verbar": "|",
	"Verbar": "â€–",
	"vert": "|",
	"Vert": "â€–",
	"VerticalBar": "âˆ£",
	"VerticalLine": "|",
	"VerticalSeparator": "â˜",
	"VerticalTilde": "â‰€",
	"VeryThinSpace": "â€Š",
	"Vfr": "ð”™",
	"vfr": "ð”³",
	"vltri": "âŠ²",
	"vnsub": "âŠ‚âƒ’",
	"vnsup": "âŠƒâƒ’",
	"Vopf": "ð•",
	"vopf": "ð•§",
	"vprop": "âˆ",
	"vrtri": "âŠ³",
	"Vscr": "ð’±",
	"vscr": "ð“‹",
	"vsubnE": "â«‹ï¸€",
	"vsubne": "âŠŠï¸€",
	"vsupnE": "â«Œï¸€",
	"vsupne": "âŠ‹ï¸€",
	"Vvdash": "âŠª",
	"vzigzag": "â¦š",
	"Wcirc": "Å´",
	"wcirc": "Åµ",
	"wedbar": "â©Ÿ",
	"wedge": "âˆ§",
	"Wedge": "â‹€",
	"wedgeq": "â‰™",
	"weierp": "â„˜",
	"Wfr": "ð”š",
	"wfr": "ð”´",
	"Wopf": "ð•Ž",
	"wopf": "ð•¨",
	"wp": "â„˜",
	"wr": "â‰€",
	"wreath": "â‰€",
	"Wscr": "ð’²",
	"wscr": "ð“Œ",
	"xcap": "â‹‚",
	"xcirc": "â—¯",
	"xcup": "â‹ƒ",
	"xdtri": "â–½",
	"Xfr": "ð”›",
	"xfr": "ð”µ",
	"xharr": "âŸ·",
	"xhArr": "âŸº",
	"Xi": "Îž",
	"xi": "Î¾",
	"xlarr": "âŸµ",
	"xlArr": "âŸ¸",
	"xmap": "âŸ¼",
	"xnis": "â‹»",
	"xodot": "â¨€",
	"Xopf": "ð•",
	"xopf": "ð•©",
	"xoplus": "â¨",
	"xotime": "â¨‚",
	"xrarr": "âŸ¶",
	"xrArr": "âŸ¹",
	"Xscr": "ð’³",
	"xscr": "ð“",
	"xsqcup": "â¨†",
	"xuplus": "â¨„",
	"xutri": "â–³",
	"xvee": "â‹",
	"xwedge": "â‹€",
	"Yacute": "Ã",
	"yacute": "Ã½",
	"YAcy": "Ð¯",
	"yacy": "Ñ",
	"Ycirc": "Å¶",
	"ycirc": "Å·",
	"Ycy": "Ð«",
	"ycy": "Ñ‹",
	"yen": "Â¥",
	"Yfr": "ð”œ",
	"yfr": "ð”¶",
	"YIcy": "Ð‡",
	"yicy": "Ñ—",
	"Yopf": "ð•",
	"yopf": "ð•ª",
	"Yscr": "ð’´",
	"yscr": "ð“Ž",
	"YUcy": "Ð®",
	"yucy": "ÑŽ",
	"yuml": "Ã¿",
	"Yuml": "Å¸",
	"Zacute": "Å¹",
	"zacute": "Åº",
	"Zcaron": "Å½",
	"zcaron": "Å¾",
	"Zcy": "Ð—",
	"zcy": "Ð·",
	"Zdot": "Å»",
	"zdot": "Å¼",
	"zeetrf": "â„¨",
	"ZeroWidthSpace": "â€‹",
	"Zeta": "Î–",
	"zeta": "Î¶",
	"zfr": "ð”·",
	"Zfr": "â„¨",
	"ZHcy": "Ð–",
	"zhcy": "Ð¶",
	"zigrarr": "â‡",
	"zopf": "ð•«",
	"Zopf": "â„¤",
	"Zscr": "ð’µ",
	"zscr": "ð“",
	"zwj": "â€",
	"zwnj": "â€Œ"
};

/***/ }),
/* 17 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";



////////////////////////////////////////////////////////////////////////////////
// Helpers

// Merge objects
//
function assign(obj /*from1, from2, from3, ...*/) {
  var sources = Array.prototype.slice.call(arguments, 1);

  sources.forEach(function (source) {
    if (!source) { return; }

    Object.keys(source).forEach(function (key) {
      obj[key] = source[key];
    });
  });

  return obj;
}

function _class(obj) { return Object.prototype.toString.call(obj); }
function isString(obj) { return _class(obj) === '[object String]'; }
function isObject(obj) { return _class(obj) === '[object Object]'; }
function isRegExp(obj) { return _class(obj) === '[object RegExp]'; }
function isFunction(obj) { return _class(obj) === '[object Function]'; }


function escapeRE(str) { return str.replace(/[.?*+^$[\]\\(){}|-]/g, '\\$&'); }

////////////////////////////////////////////////////////////////////////////////


var defaultOptions = {
  fuzzyLink: true,
  fuzzyEmail: true,
  fuzzyIP: false
};


function isOptionsObj(obj) {
  return Object.keys(obj || {}).reduce(function (acc, k) {
    return acc || defaultOptions.hasOwnProperty(k);
  }, false);
}


var defaultSchemas = {
  'http:': {
    validate: function (text, pos, self) {
      var tail = text.slice(pos);

      if (!self.re.http) {
        // compile lazily, because "host"-containing variables can change on tlds update.
        self.re.http =  new RegExp(
          '^\\/\\/' + self.re.src_auth + self.re.src_host_port_strict + self.re.src_path, 'i'
        );
      }
      if (self.re.http.test(tail)) {
        return tail.match(self.re.http)[0].length;
      }
      return 0;
    }
  },
  'https:':  'http:',
  'ftp:':    'http:',
  '//':      {
    validate: function (text, pos, self) {
      var tail = text.slice(pos);

      if (!self.re.no_http) {
      // compile lazily, because "host"-containing variables can change on tlds update.
        self.re.no_http =  new RegExp(
          '^' +
          self.re.src_auth +
          // Don't allow single-level domains, because of false positives like '//test'
          // with code comments
          '(?:localhost|(?:(?:' + self.re.src_domain + ')\\.)+' + self.re.src_domain_root + ')' +
          self.re.src_port +
          self.re.src_host_terminator +
          self.re.src_path,

          'i'
        );
      }

      if (self.re.no_http.test(tail)) {
        // should not be `://` & `///`, that protects from errors in protocol name
        if (pos >= 3 && text[pos - 3] === ':') { return 0; }
        if (pos >= 3 && text[pos - 3] === '/') { return 0; }
        return tail.match(self.re.no_http)[0].length;
      }
      return 0;
    }
  },
  'mailto:': {
    validate: function (text, pos, self) {
      var tail = text.slice(pos);

      if (!self.re.mailto) {
        self.re.mailto =  new RegExp(
          '^' + self.re.src_email_name + '@' + self.re.src_host_strict, 'i'
        );
      }
      if (self.re.mailto.test(tail)) {
        return tail.match(self.re.mailto)[0].length;
      }
      return 0;
    }
  }
};

/*eslint-disable max-len*/

// RE pattern for 2-character tlds (autogenerated by ./support/tlds_2char_gen.js)
var tlds_2ch_src_re = 'a[cdefgilmnoqrstuwxz]|b[abdefghijmnorstvwyz]|c[acdfghiklmnoruvwxyz]|d[ejkmoz]|e[cegrstu]|f[ijkmor]|g[abdefghilmnpqrstuwy]|h[kmnrtu]|i[delmnoqrst]|j[emop]|k[eghimnprwyz]|l[abcikrstuvy]|m[acdeghklmnopqrstuvwxyz]|n[acefgilopruz]|om|p[aefghklmnrstwy]|qa|r[eosuw]|s[abcdeghijklmnortuvxyz]|t[cdfghjklmnortvwz]|u[agksyz]|v[aceginu]|w[fs]|y[et]|z[amw]';

// DON'T try to make PRs with changes. Extend TLDs with LinkifyIt.tlds() instead
var tlds_default = 'biz|com|edu|gov|net|org|pro|web|xxx|aero|asia|coop|info|museum|name|shop|Ñ€Ñ„'.split('|');

/*eslint-enable max-len*/

////////////////////////////////////////////////////////////////////////////////

function resetScanCache(self) {
  self.__index__ = -1;
  self.__text_cache__   = '';
}

function createValidator(re) {
  return function (text, pos) {
    var tail = text.slice(pos);

    if (re.test(tail)) {
      return tail.match(re)[0].length;
    }
    return 0;
  };
}

function createNormalizer() {
  return function (match, self) {
    self.normalize(match);
  };
}

// Schemas compiler. Build regexps.
//
function compile(self) {

  // Load & clone RE patterns.
  var re = self.re = __webpack_require__(18)(self.__opts__);

  // Define dynamic patterns
  var tlds = self.__tlds__.slice();

  self.onCompile();

  if (!self.__tlds_replaced__) {
    tlds.push(tlds_2ch_src_re);
  }
  tlds.push(re.src_xn);

  re.src_tlds = tlds.join('|');

  function untpl(tpl) { return tpl.replace('%TLDS%', re.src_tlds); }

  re.email_fuzzy      = RegExp(untpl(re.tpl_email_fuzzy), 'i');
  re.link_fuzzy       = RegExp(untpl(re.tpl_link_fuzzy), 'i');
  re.link_no_ip_fuzzy = RegExp(untpl(re.tpl_link_no_ip_fuzzy), 'i');
  re.host_fuzzy_test  = RegExp(untpl(re.tpl_host_fuzzy_test), 'i');

  //
  // Compile each schema
  //

  var aliases = [];

  self.__compiled__ = {}; // Reset compiled data

  function schemaError(name, val) {
    throw new Error('(LinkifyIt) Invalid schema "' + name + '": ' + val);
  }

  Object.keys(self.__schemas__).forEach(function (name) {
    var val = self.__schemas__[name];

    // skip disabled methods
    if (val === null) { return; }

    var compiled = { validate: null, link: null };

    self.__compiled__[name] = compiled;

    if (isObject(val)) {
      if (isRegExp(val.validate)) {
        compiled.validate = createValidator(val.validate);
      } else if (isFunction(val.validate)) {
        compiled.validate = val.validate;
      } else {
        schemaError(name, val);
      }

      if (isFunction(val.normalize)) {
        compiled.normalize = val.normalize;
      } else if (!val.normalize) {
        compiled.normalize = createNormalizer();
      } else {
        schemaError(name, val);
      }

      return;
    }

    if (isString(val)) {
      aliases.push(name);
      return;
    }

    schemaError(name, val);
  });

  //
  // Compile postponed aliases
  //

  aliases.forEach(function (alias) {
    if (!self.__compiled__[self.__schemas__[alias]]) {
      // Silently fail on missed schemas to avoid errons on disable.
      // schemaError(alias, self.__schemas__[alias]);
      return;
    }

    self.__compiled__[alias].validate =
      self.__compiled__[self.__schemas__[alias]].validate;
    self.__compiled__[alias].normalize =
      self.__compiled__[self.__schemas__[alias]].normalize;
  });

  //
  // Fake record for guessed links
  //
  self.__compiled__[''] = { validate: null, normalize: createNormalizer() };

  //
  // Build schema condition
  //
  var slist = Object.keys(self.__compiled__)
                      .filter(function (name) {
                        // Filter disabled & fake schemas
                        return name.length > 0 && self.__compiled__[name];
                      })
                      .map(escapeRE)
                      .join('|');
  // (?!_) cause 1.5x slowdown
  self.re.schema_test   = RegExp('(^|(?!_)(?:[><\uff5c]|' + re.src_ZPCc + '))(' + slist + ')', 'i');
  self.re.schema_search = RegExp('(^|(?!_)(?:[><\uff5c]|' + re.src_ZPCc + '))(' + slist + ')', 'ig');

  self.re.pretest       = RegExp(
                            '(' + self.re.schema_test.source + ')|' +
                            '(' + self.re.host_fuzzy_test.source + ')|' +
                            '@',
                            'i');

  //
  // Cleanup
  //

  resetScanCache(self);
}

/**
 * class Match
 *
 * Match result. Single element of array, returned by [[LinkifyIt#match]]
 **/
function Match(self, shift) {
  var start = self.__index__,
      end   = self.__last_index__,
      text  = self.__text_cache__.slice(start, end);

  /**
   * Match#schema -> String
   *
   * Prefix (protocol) for matched string.
   **/
  this.schema    = self.__schema__.toLowerCase();
  /**
   * Match#index -> Number
   *
   * First position of matched string.
   **/
  this.index     = start + shift;
  /**
   * Match#lastIndex -> Number
   *
   * Next position after matched string.
   **/
  this.lastIndex = end + shift;
  /**
   * Match#raw -> String
   *
   * Matched string.
   **/
  this.raw       = text;
  /**
   * Match#text -> String
   *
   * Notmalized text of matched string.
   **/
  this.text      = text;
  /**
   * Match#url -> String
   *
   * Normalized url of matched string.
   **/
  this.url       = text;
}

function createMatch(self, shift) {
  var match = new Match(self, shift);

  self.__compiled__[match.schema].normalize(match, self);

  return match;
}


/**
 * class LinkifyIt
 **/

/**
 * new LinkifyIt(schemas, options)
 * - schemas (Object): Optional. Additional schemas to validate (prefix/validator)
 * - options (Object): { fuzzyLink|fuzzyEmail|fuzzyIP: true|false }
 *
 * Creates new linkifier instance with optional additional schemas.
 * Can be called without `new` keyword for convenience.
 *
 * By default understands:
 *
 * - `http(s)://...` , `ftp://...`, `mailto:...` & `//...` links
 * - "fuzzy" links and emails (example.com, foo@bar.com).
 *
 * `schemas` is an object, where each key/value describes protocol/rule:
 *
 * - __key__ - link prefix (usually, protocol name with `:` at the end, `skype:`
 *   for example). `linkify-it` makes shure that prefix is not preceeded with
 *   alphanumeric char and symbols. Only whitespaces and punctuation allowed.
 * - __value__ - rule to check tail after link prefix
 *   - _String_ - just alias to existing rule
 *   - _Object_
 *     - _validate_ - validator function (should return matched length on success),
 *       or `RegExp`.
 *     - _normalize_ - optional function to normalize text & url of matched result
 *       (for example, for @twitter mentions).
 *
 * `options`:
 *
 * - __fuzzyLink__ - recognige URL-s without `http(s):` prefix. Default `true`.
 * - __fuzzyIP__ - allow IPs in fuzzy links above. Can conflict with some texts
 *   like version numbers. Default `false`.
 * - __fuzzyEmail__ - recognize emails without `mailto:` prefix.
 *
 **/
function LinkifyIt(schemas, options) {
  if (!(this instanceof LinkifyIt)) {
    return new LinkifyIt(schemas, options);
  }

  if (!options) {
    if (isOptionsObj(schemas)) {
      options = schemas;
      schemas = {};
    }
  }

  this.__opts__           = assign({}, defaultOptions, options);

  // Cache last tested result. Used to skip repeating steps on next `match` call.
  this.__index__          = -1;
  this.__last_index__     = -1; // Next scan position
  this.__schema__         = '';
  this.__text_cache__     = '';

  this.__schemas__        = assign({}, defaultSchemas, schemas);
  this.__compiled__       = {};

  this.__tlds__           = tlds_default;
  this.__tlds_replaced__  = false;

  this.re = {};

  compile(this);
}


/** chainable
 * LinkifyIt#add(schema, definition)
 * - schema (String): rule name (fixed pattern prefix)
 * - definition (String|RegExp|Object): schema definition
 *
 * Add new rule definition. See constructor description for details.
 **/
LinkifyIt.prototype.add = function add(schema, definition) {
  this.__schemas__[schema] = definition;
  compile(this);
  return this;
};


/** chainable
 * LinkifyIt#set(options)
 * - options (Object): { fuzzyLink|fuzzyEmail|fuzzyIP: true|false }
 *
 * Set recognition options for links without schema.
 **/
LinkifyIt.prototype.set = function set(options) {
  this.__opts__ = assign(this.__opts__, options);
  return this;
};


/**
 * LinkifyIt#test(text) -> Boolean
 *
 * Searches linkifiable pattern and returns `true` on success or `false` on fail.
 **/
LinkifyIt.prototype.test = function test(text) {
  // Reset scan cache
  this.__text_cache__ = text;
  this.__index__      = -1;

  if (!text.length) { return false; }

  var m, ml, me, len, shift, next, re, tld_pos, at_pos;

  // try to scan for link with schema - that's the most simple rule
  if (this.re.schema_test.test(text)) {
    re = this.re.schema_search;
    re.lastIndex = 0;
    while ((m = re.exec(text)) !== null) {
      len = this.testSchemaAt(text, m[2], re.lastIndex);
      if (len) {
        this.__schema__     = m[2];
        this.__index__      = m.index + m[1].length;
        this.__last_index__ = m.index + m[0].length + len;
        break;
      }
    }
  }

  if (this.__opts__.fuzzyLink && this.__compiled__['http:']) {
    // guess schemaless links
    tld_pos = text.search(this.re.host_fuzzy_test);
    if (tld_pos >= 0) {
      // if tld is located after found link - no need to check fuzzy pattern
      if (this.__index__ < 0 || tld_pos < this.__index__) {
        if ((ml = text.match(this.__opts__.fuzzyIP ? this.re.link_fuzzy : this.re.link_no_ip_fuzzy)) !== null) {

          shift = ml.index + ml[1].length;

          if (this.__index__ < 0 || shift < this.__index__) {
            this.__schema__     = '';
            this.__index__      = shift;
            this.__last_index__ = ml.index + ml[0].length;
          }
        }
      }
    }
  }

  if (this.__opts__.fuzzyEmail && this.__compiled__['mailto:']) {
    // guess schemaless emails
    at_pos = text.indexOf('@');
    if (at_pos >= 0) {
      // We can't skip this check, because this cases are possible:
      // 192.168.1.1@gmail.com, my.in@example.com
      if ((me = text.match(this.re.email_fuzzy)) !== null) {

        shift = me.index + me[1].length;
        next  = me.index + me[0].length;

        if (this.__index__ < 0 || shift < this.__index__ ||
            (shift === this.__index__ && next > this.__last_index__)) {
          this.__schema__     = 'mailto:';
          this.__index__      = shift;
          this.__last_index__ = next;
        }
      }
    }
  }

  return this.__index__ >= 0;
};


/**
 * LinkifyIt#pretest(text) -> Boolean
 *
 * Very quick check, that can give false positives. Returns true if link MAY BE
 * can exists. Can be used for speed optimization, when you need to check that
 * link NOT exists.
 **/
LinkifyIt.prototype.pretest = function pretest(text) {
  return this.re.pretest.test(text);
};


/**
 * LinkifyIt#testSchemaAt(text, name, position) -> Number
 * - text (String): text to scan
 * - name (String): rule (schema) name
 * - position (Number): text offset to check from
 *
 * Similar to [[LinkifyIt#test]] but checks only specific protocol tail exactly
 * at given position. Returns length of found pattern (0 on fail).
 **/
LinkifyIt.prototype.testSchemaAt = function testSchemaAt(text, schema, pos) {
  // If not supported schema check requested - terminate
  if (!this.__compiled__[schema.toLowerCase()]) {
    return 0;
  }
  return this.__compiled__[schema.toLowerCase()].validate(text, pos, this);
};


/**
 * LinkifyIt#match(text) -> Array|null
 *
 * Returns array of found link descriptions or `null` on fail. We strongly
 * recommend to use [[LinkifyIt#test]] first, for best speed.
 *
 * ##### Result match description
 *
 * - __schema__ - link schema, can be empty for fuzzy links, or `//` for
 *   protocol-neutral  links.
 * - __index__ - offset of matched text
 * - __lastIndex__ - index of next char after mathch end
 * - __raw__ - matched text
 * - __text__ - normalized text
 * - __url__ - link, generated from matched text
 **/
LinkifyIt.prototype.match = function match(text) {
  var shift = 0, result = [];

  // Try to take previous element from cache, if .test() called before
  if (this.__index__ >= 0 && this.__text_cache__ === text) {
    result.push(createMatch(this, shift));
    shift = this.__last_index__;
  }

  // Cut head if cache was used
  var tail = shift ? text.slice(shift) : text;

  // Scan string until end reached
  while (this.test(tail)) {
    result.push(createMatch(this, shift));

    tail = tail.slice(this.__last_index__);
    shift += this.__last_index__;
  }

  if (result.length) {
    return result;
  }

  return null;
};


/** chainable
 * LinkifyIt#tlds(list [, keepOld]) -> this
 * - list (Array): list of tlds
 * - keepOld (Boolean): merge with current list if `true` (`false` by default)
 *
 * Load (or merge) new tlds list. Those are user for fuzzy links (without prefix)
 * to avoid false positives. By default this algorythm used:
 *
 * - hostname with any 2-letter root zones are ok.
 * - biz|com|edu|gov|net|org|pro|web|xxx|aero|asia|coop|info|museum|name|shop|Ñ€Ñ„
 *   are ok.
 * - encoded (`xn--...`) root zones are ok.
 *
 * If list is replaced, then exact match for 2-chars root zones will be checked.
 **/
LinkifyIt.prototype.tlds = function tlds(list, keepOld) {
  list = Array.isArray(list) ? list : [ list ];

  if (!keepOld) {
    this.__tlds__ = list.slice();
    this.__tlds_replaced__ = true;
    compile(this);
    return this;
  }

  this.__tlds__ = this.__tlds__.concat(list)
                                  .sort()
                                  .filter(function (el, idx, arr) {
                                    return el !== arr[idx - 1];
                                  })
                                  .reverse();

  compile(this);
  return this;
};

/**
 * LinkifyIt#normalize(match)
 *
 * Default normalizer (if schema does not define it's own).
 **/
LinkifyIt.prototype.normalize = function normalize(match) {

  // Do minimal possible changes by default. Need to collect feedback prior
  // to move forward https://github.com/markdown-it/linkify-it/issues/1

  if (!match.schema) { match.url = 'http://' + match.url; }

  if (match.schema === 'mailto:' && !/^mailto:/i.test(match.url)) {
    match.url = 'mailto:' + match.url;
  }
};


/**
 * LinkifyIt#onCompile()
 *
 * Override to modify basic RegExp-s.
 **/
LinkifyIt.prototype.onCompile = function onCompile() {
};


module.exports = LinkifyIt;


/***/ }),
/* 18 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";



module.exports = function (opts) {
  var re = {};

  // Use direct extract instead of `regenerate` to reduse browserified size
  re.src_Any = __webpack_require__(13).source;
  re.src_Cc  = __webpack_require__(11).source;
  re.src_Z   = __webpack_require__(12).source;
  re.src_P   = __webpack_require__(3).source;

  // \p{\Z\P\Cc\CF} (white spaces + control + format + punctuation)
  re.src_ZPCc = [ re.src_Z, re.src_P, re.src_Cc ].join('|');

  // \p{\Z\Cc} (white spaces + control)
  re.src_ZCc = [ re.src_Z, re.src_Cc ].join('|');

  // Experimental. List of chars, completely prohibited in links
  // because can separate it from other part of text
  var text_separators = '[><\uff5c]';

  // All possible word characters (everything without punctuation, spaces & controls)
  // Defined via punctuation & spaces to save space
  // Should be something like \p{\L\N\S\M} (\w but without `_`)
  re.src_pseudo_letter       = '(?:(?!' + text_separators + '|' + re.src_ZPCc + ')' + re.src_Any + ')';
  // The same as abothe but without [0-9]
  // var src_pseudo_letter_non_d = '(?:(?![0-9]|' + src_ZPCc + ')' + src_Any + ')';

  ////////////////////////////////////////////////////////////////////////////////

  re.src_ip4 =

    '(?:(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)';

  // Prohibit any of "@/[]()" in user/pass to avoid wrong domain fetch.
  re.src_auth    = '(?:(?:(?!' + re.src_ZCc + '|[@/\\[\\]()]).)+@)?';

  re.src_port =

    '(?::(?:6(?:[0-4]\\d{3}|5(?:[0-4]\\d{2}|5(?:[0-2]\\d|3[0-5])))|[1-5]?\\d{1,4}))?';

  re.src_host_terminator =

    '(?=$|' + text_separators + '|' + re.src_ZPCc + ')(?!-|_|:\\d|\\.-|\\.(?!$|' + re.src_ZPCc + '))';

  re.src_path =

    '(?:' +
      '[/?#]' +
        '(?:' +
          '(?!' + re.src_ZCc + '|' + text_separators + '|[()[\\]{}.,"\'?!\\-]).|' +
          '\\[(?:(?!' + re.src_ZCc + '|\\]).)*\\]|' +
          '\\((?:(?!' + re.src_ZCc + '|[)]).)*\\)|' +
          '\\{(?:(?!' + re.src_ZCc + '|[}]).)*\\}|' +
          '\\"(?:(?!' + re.src_ZCc + '|["]).)+\\"|' +
          "\\'(?:(?!" + re.src_ZCc + "|[']).)+\\'|" +
          "\\'(?=" + re.src_pseudo_letter + '|[-]).|' +  // allow `I'm_king` if no pair found
          '\\.{2,3}[a-zA-Z0-9%/]|' + // github has ... in commit range links. Restrict to
                                     // - english
                                     // - percent-encoded
                                     // - parts of file path
                                     // until more examples found.
          '\\.(?!' + re.src_ZCc + '|[.]).|' +
          (opts && opts['---'] ?
            '\\-(?!--(?:[^-]|$))(?:-*)|' // `---` => long dash, terminate
          :
            '\\-+|'
          ) +
          '\\,(?!' + re.src_ZCc + ').|' +      // allow `,,,` in paths
          '\\!(?!' + re.src_ZCc + '|[!]).|' +
          '\\?(?!' + re.src_ZCc + '|[?]).' +
        ')+' +
      '|\\/' +
    ')?';

  re.src_email_name =

    '[\\-;:&=\\+\\$,\\"\\.a-zA-Z0-9_]+';

  re.src_xn =

    'xn--[a-z0-9\\-]{1,59}';

  // More to read about domain names
  // http://serverfault.com/questions/638260/

  re.src_domain_root =

    // Allow letters & digits (http://test1)
    '(?:' +
      re.src_xn +
      '|' +
      re.src_pseudo_letter + '{1,63}' +
    ')';

  re.src_domain =

    '(?:' +
      re.src_xn +
      '|' +
      '(?:' + re.src_pseudo_letter + ')' +
      '|' +
      // don't allow `--` in domain names, because:
      // - that can conflict with markdown &mdash; / &ndash;
      // - nobody use those anyway
      '(?:' + re.src_pseudo_letter + '(?:-(?!-)|' + re.src_pseudo_letter + '){0,61}' + re.src_pseudo_letter + ')' +
    ')';

  re.src_host =

    '(?:' +
    // Don't need IP check, because digits are already allowed in normal domain names
    //   src_ip4 +
    // '|' +
      '(?:(?:(?:' + re.src_domain + ')\\.)*' + re.src_domain/*_root*/ + ')' +
    ')';

  re.tpl_host_fuzzy =

    '(?:' +
      re.src_ip4 +
    '|' +
      '(?:(?:(?:' + re.src_domain + ')\\.)+(?:%TLDS%))' +
    ')';

  re.tpl_host_no_ip_fuzzy =

    '(?:(?:(?:' + re.src_domain + ')\\.)+(?:%TLDS%))';

  re.src_host_strict =

    re.src_host + re.src_host_terminator;

  re.tpl_host_fuzzy_strict =

    re.tpl_host_fuzzy + re.src_host_terminator;

  re.src_host_port_strict =

    re.src_host + re.src_port + re.src_host_terminator;

  re.tpl_host_port_fuzzy_strict =

    re.tpl_host_fuzzy + re.src_port + re.src_host_terminator;

  re.tpl_host_port_no_ip_fuzzy_strict =

    re.tpl_host_no_ip_fuzzy + re.src_port + re.src_host_terminator;


  ////////////////////////////////////////////////////////////////////////////////
  // Main rules

  // Rude test fuzzy links by host, for quick deny
  re.tpl_host_fuzzy_test =

    'localhost|www\\.|\\.\\d{1,3}\\.|(?:\\.(?:%TLDS%)(?:' + re.src_ZPCc + '|>|$))';

  re.tpl_email_fuzzy =

      '(^|' + text_separators + '|\\(|' + re.src_ZCc + ')(' + re.src_email_name + '@' + re.tpl_host_fuzzy_strict + ')';

  re.tpl_link_fuzzy =
      // Fuzzy link can't be prepended with .:/\- and non punctuation.
      // but can start with > (markdown blockquote)
      '(^|(?![.:/\\-_@])(?:[$+<=>^`|\uff5c]|' + re.src_ZPCc + '))' +
      '((?![$+<=>^`|\uff5c])' + re.tpl_host_port_fuzzy_strict + re.src_path + ')';

  re.tpl_link_no_ip_fuzzy =
      // Fuzzy link can't be prepended with .:/\- and non punctuation.
      // but can start with > (markdown blockquote)
      '(^|(?![.:/\\-_@])(?:[$+<=>^`|\uff5c]|' + re.src_ZPCc + '))' +
      '((?![$+<=>^`|\uff5c])' + re.tpl_host_port_no_ip_fuzzy_strict + re.src_path + ')';

  return re;
};


/***/ }),
/* 19 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";



module.exports = __webpack_require__(25);


/***/ }),
/* 20 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// List of valid html blocks names, accorting to commonmark spec
// http://jgm.github.io/CommonMark/spec.html#html-blocks




module.exports = [
  'address',
  'article',
  'aside',
  'base',
  'basefont',
  'blockquote',
  'body',
  'caption',
  'center',
  'col',
  'colgroup',
  'dd',
  'details',
  'dialog',
  'dir',
  'div',
  'dl',
  'dt',
  'fieldset',
  'figcaption',
  'figure',
  'footer',
  'form',
  'frame',
  'frameset',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'head',
  'header',
  'hr',
  'html',
  'iframe',
  'legend',
  'li',
  'link',
  'main',
  'menu',
  'menuitem',
  'meta',
  'nav',
  'noframes',
  'ol',
  'optgroup',
  'option',
  'p',
  'param',
  'pre',
  'section',
  'source',
  'title',
  'summary',
  'table',
  'tbody',
  'td',
  'tfoot',
  'th',
  'thead',
  'title',
  'tr',
  'track',
  'ul'
];


/***/ }),
/* 21 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Just a shortcut for bulk export



exports.parseLinkLabel       = __webpack_require__(23);
exports.parseLinkDestination = __webpack_require__(22);
exports.parseLinkTitle       = __webpack_require__(24);


/***/ }),
/* 22 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Parse link destination
//



var isSpace     = __webpack_require__(0).isSpace;
var unescapeAll = __webpack_require__(0).unescapeAll;


module.exports = function parseLinkDestination(str, pos, max) {
  var code, level,
      lines = 0,
      start = pos,
      result = {
        ok: false,
        pos: 0,
        lines: 0,
        str: ''
      };

  if (str.charCodeAt(pos) === 0x3C /* < */) {
    pos++;
    while (pos < max) {
      code = str.charCodeAt(pos);
      if (code === 0x0A /* \n */ || isSpace(code)) { return result; }
      if (code === 0x3E /* > */) {
        result.pos = pos + 1;
        result.str = unescapeAll(str.slice(start + 1, pos));
        result.ok = true;
        return result;
      }
      if (code === 0x5C /* \ */ && pos + 1 < max) {
        pos += 2;
        continue;
      }

      pos++;
    }

    // no closing '>'
    return result;
  }

  // this should be ... } else { ... branch

  level = 0;
  while (pos < max) {
    code = str.charCodeAt(pos);

    if (code === 0x20) { break; }

    // ascii control characters
    if (code < 0x20 || code === 0x7F) { break; }

    if (code === 0x5C /* \ */ && pos + 1 < max) {
      pos += 2;
      continue;
    }

    if (code === 0x28 /* ( */) {
      level++;
      if (level > 1) { break; }
    }

    if (code === 0x29 /* ) */) {
      level--;
      if (level < 0) { break; }
    }

    pos++;
  }

  if (start === pos) { return result; }

  result.str = unescapeAll(str.slice(start, pos));
  result.lines = lines;
  result.pos = pos;
  result.ok = true;
  return result;
};


/***/ }),
/* 23 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Parse link label
//
// this function assumes that first character ("[") already matches;
// returns the end of the label
//


module.exports = function parseLinkLabel(state, start, disableNested) {
  var level, found, marker, prevPos,
      labelEnd = -1,
      max = state.posMax,
      oldPos = state.pos;

  state.pos = start + 1;
  level = 1;

  while (state.pos < max) {
    marker = state.src.charCodeAt(state.pos);
    if (marker === 0x5D /* ] */) {
      level--;
      if (level === 0) {
        found = true;
        break;
      }
    }

    prevPos = state.pos;
    state.md.inline.skipToken(state);
    if (marker === 0x5B /* [ */) {
      if (prevPos === state.pos - 1) {
        // increase level if we find text `[`, which is not a part of any token
        level++;
      } else if (disableNested) {
        state.pos = oldPos;
        return -1;
      }
    }
  }

  if (found) {
    labelEnd = state.pos;
  }

  // restore old state
  state.pos = oldPos;

  return labelEnd;
};


/***/ }),
/* 24 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Parse link title
//



var unescapeAll = __webpack_require__(0).unescapeAll;


module.exports = function parseLinkTitle(str, pos, max) {
  var code,
      marker,
      lines = 0,
      start = pos,
      result = {
        ok: false,
        pos: 0,
        lines: 0,
        str: ''
      };

  if (pos >= max) { return result; }

  marker = str.charCodeAt(pos);

  if (marker !== 0x22 /* " */ && marker !== 0x27 /* ' */ && marker !== 0x28 /* ( */) { return result; }

  pos++;

  // if opening marker is "(", switch it to closing marker ")"
  if (marker === 0x28) { marker = 0x29; }

  while (pos < max) {
    code = str.charCodeAt(pos);
    if (code === marker) {
      result.pos = pos + 1;
      result.lines = lines;
      result.str = unescapeAll(str.slice(start + 1, pos));
      result.ok = true;
      return result;
    } else if (code === 0x0A) {
      lines++;
    } else if (code === 0x5C /* \ */ && pos + 1 < max) {
      pos++;
      if (str.charCodeAt(pos) === 0x0A) {
        lines++;
      }
    }

    pos++;
  }

  return result;
};


/***/ }),
/* 25 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Main parser class




var utils        = __webpack_require__(0);
var helpers      = __webpack_require__(21);
var Renderer     = __webpack_require__(32);
var ParserCore   = __webpack_require__(27);
var ParserBlock  = __webpack_require__(26);
var ParserInline = __webpack_require__(28);
var LinkifyIt    = __webpack_require__(17);
var mdurl        = __webpack_require__(9);
var punycode     = __webpack_require__(68);


var config = {
  'default': __webpack_require__(30),
  zero: __webpack_require__(31),
  commonmark: __webpack_require__(29)
};

////////////////////////////////////////////////////////////////////////////////
//
// This validator can prohibit more than really needed to prevent XSS. It's a
// tradeoff to keep code simple and to be secure by default.
//
// If you need different setup - override validator method as you wish. Or
// replace it with dummy function and use external sanitizer.
//

var BAD_PROTO_RE = /^(vbscript|javascript|file|data):/;
var GOOD_DATA_RE = /^data:image\/(gif|png|jpeg|webp);/;

function validateLink(url) {
  // url should be normalized at this point, and existing entities are decoded
  var str = url.trim().toLowerCase();

  return BAD_PROTO_RE.test(str) ? (GOOD_DATA_RE.test(str) ? true : false) : true;
}

////////////////////////////////////////////////////////////////////////////////


var RECODE_HOSTNAME_FOR = [ 'http:', 'https:', 'mailto:' ];

function normalizeLink(url) {
  var parsed = mdurl.parse(url, true);

  if (parsed.hostname) {
    // Encode hostnames in urls like:
    // `http://host/`, `https://host/`, `mailto:user@host`, `//host/`
    //
    // We don't encode unknown schemas, because it's likely that we encode
    // something we shouldn't (e.g. `skype:name` treated as `skype:host`)
    //
    if (!parsed.protocol || RECODE_HOSTNAME_FOR.indexOf(parsed.protocol) >= 0) {
      try {
        parsed.hostname = punycode.toASCII(parsed.hostname);
      } catch (er) { /**/ }
    }
  }

  return mdurl.encode(mdurl.format(parsed));
}

function normalizeLinkText(url) {
  var parsed = mdurl.parse(url, true);

  if (parsed.hostname) {
    // Encode hostnames in urls like:
    // `http://host/`, `https://host/`, `mailto:user@host`, `//host/`
    //
    // We don't encode unknown schemas, because it's likely that we encode
    // something we shouldn't (e.g. `skype:name` treated as `skype:host`)
    //
    if (!parsed.protocol || RECODE_HOSTNAME_FOR.indexOf(parsed.protocol) >= 0) {
      try {
        parsed.hostname = punycode.toUnicode(parsed.hostname);
      } catch (er) { /**/ }
    }
  }

  return mdurl.decode(mdurl.format(parsed));
}


/**
 * class MarkdownIt
 *
 * Main parser/renderer class.
 *
 * ##### Usage
 *
 * ```javascript
 * // node.js, "classic" way:
 * var MarkdownIt = require('markdown-it'),
 *     md = new MarkdownIt();
 * var result = md.render('# markdown-it rulezz!');
 *
 * // node.js, the same, but with sugar:
 * var md = require('markdown-it')();
 * var result = md.render('# markdown-it rulezz!');
 *
 * // browser without AMD, added to "window" on script load
 * // Note, there are no dash.
 * var md = window.markdownit();
 * var result = md.render('# markdown-it rulezz!');
 * ```
 *
 * Single line rendering, without paragraph wrap:
 *
 * ```javascript
 * var md = require('markdown-it')();
 * var result = md.renderInline('__markdown-it__ rulezz!');
 * ```
 **/

/**
 * new MarkdownIt([presetName, options])
 * - presetName (String): optional, `commonmark` / `zero`
 * - options (Object)
 *
 * Creates parser instanse with given config. Can be called without `new`.
 *
 * ##### presetName
 *
 * MarkdownIt provides named presets as a convenience to quickly
 * enable/disable active syntax rules and options for common use cases.
 *
 * - ["commonmark"](https://github.com/markdown-it/markdown-it/blob/master/lib/presets/commonmark.js) -
 *   configures parser to strict [CommonMark](http://commonmark.org/) mode.
 * - [default](https://github.com/markdown-it/markdown-it/blob/master/lib/presets/default.js) -
 *   similar to GFM, used when no preset name given. Enables all available rules,
 *   but still without html, typographer & autolinker.
 * - ["zero"](https://github.com/markdown-it/markdown-it/blob/master/lib/presets/zero.js) -
 *   all rules disabled. Useful to quickly setup your config via `.enable()`.
 *   For example, when you need only `bold` and `italic` markup and nothing else.
 *
 * ##### options:
 *
 * - __html__ - `false`. Set `true` to enable HTML tags in source. Be careful!
 *   That's not safe! You may need external sanitizer to protect output from XSS.
 *   It's better to extend features via plugins, instead of enabling HTML.
 * - __xhtmlOut__ - `false`. Set `true` to add '/' when closing single tags
 *   (`<br />`). This is needed only for full CommonMark compatibility. In real
 *   world you will need HTML output.
 * - __breaks__ - `false`. Set `true` to convert `\n` in paragraphs into `<br>`.
 * - __langPrefix__ - `language-`. CSS language class prefix for fenced blocks.
 *   Can be useful for external highlighters.
 * - __linkify__ - `false`. Set `true` to autoconvert URL-like text to links.
 * - __typographer__  - `false`. Set `true` to enable [some language-neutral
 *   replacement](https://github.com/markdown-it/markdown-it/blob/master/lib/rules_core/replacements.js) +
 *   quotes beautification (smartquotes).
 * - __quotes__ - `â€œâ€â€˜â€™`, String or Array. Double + single quotes replacement
 *   pairs, when typographer enabled and smartquotes on. For example, you can
 *   use `'Â«Â»â€žâ€œ'` for Russian, `'â€žâ€œâ€šâ€˜'` for German, and
 *   `['Â«\xA0', '\xA0Â»', 'â€¹\xA0', '\xA0â€º']` for French (including nbsp).
 * - __highlight__ - `null`. Highlighter function for fenced code blocks.
 *   Highlighter `function (str, lang)` should return escaped HTML. It can also
 *   return empty string if the source was not changed and should be escaped
 *   externaly. If result starts with <pre... internal wrapper is skipped.
 *
 * ##### Example
 *
 * ```javascript
 * // commonmark mode
 * var md = require('markdown-it')('commonmark');
 *
 * // default mode
 * var md = require('markdown-it')();
 *
 * // enable everything
 * var md = require('markdown-it')({
 *   html: true,
 *   linkify: true,
 *   typographer: true
 * });
 * ```
 *
 * ##### Syntax highlighting
 *
 * ```js
 * var hljs = require('highlight.js') // https://highlightjs.org/
 *
 * var md = require('markdown-it')({
 *   highlight: function (str, lang) {
 *     if (lang && hljs.getLanguage(lang)) {
 *       try {
 *         return hljs.highlight(lang, str, true).value;
 *       } catch (__) {}
 *     }
 *
 *     return ''; // use external default escaping
 *   }
 * });
 * ```
 *
 * Or with full wrapper override (if you need assign class to `<pre>`):
 *
 * ```javascript
 * var hljs = require('highlight.js') // https://highlightjs.org/
 *
 * // Actual default values
 * var md = require('markdown-it')({
 *   highlight: function (str, lang) {
 *     if (lang && hljs.getLanguage(lang)) {
 *       try {
 *         return '<pre class="hljs"><code>' +
 *                hljs.highlight(lang, str, true).value +
 *                '</code></pre>';
 *       } catch (__) {}
 *     }
 *
 *     return '<pre class="hljs"><code>' + md.utils.escapeHtml(str) + '</code></pre>';
 *   }
 * });
 * ```
 *
 **/
function MarkdownIt(presetName, options) {
  if (!(this instanceof MarkdownIt)) {
    return new MarkdownIt(presetName, options);
  }

  if (!options) {
    if (!utils.isString(presetName)) {
      options = presetName || {};
      presetName = 'default';
    }
  }

  /**
   * MarkdownIt#inline -> ParserInline
   *
   * Instance of [[ParserInline]]. You may need it to add new rules when
   * writing plugins. For simple rules control use [[MarkdownIt.disable]] and
   * [[MarkdownIt.enable]].
   **/
  this.inline = new ParserInline();

  /**
   * MarkdownIt#block -> ParserBlock
   *
   * Instance of [[ParserBlock]]. You may need it to add new rules when
   * writing plugins. For simple rules control use [[MarkdownIt.disable]] and
   * [[MarkdownIt.enable]].
   **/
  this.block = new ParserBlock();

  /**
   * MarkdownIt#core -> Core
   *
   * Instance of [[Core]] chain executor. You may need it to add new rules when
   * writing plugins. For simple rules control use [[MarkdownIt.disable]] and
   * [[MarkdownIt.enable]].
   **/
  this.core = new ParserCore();

  /**
   * MarkdownIt#renderer -> Renderer
   *
   * Instance of [[Renderer]]. Use it to modify output look. Or to add rendering
   * rules for new token types, generated by plugins.
   *
   * ##### Example
   *
   * ```javascript
   * var md = require('markdown-it')();
   *
   * function myToken(tokens, idx, options, env, self) {
   *   //...
   *   return result;
   * };
   *
   * md.renderer.rules['my_token'] = myToken
   * ```
   *
   * See [[Renderer]] docs and [source code](https://github.com/markdown-it/markdown-it/blob/master/lib/renderer.js).
   **/
  this.renderer = new Renderer();

  /**
   * MarkdownIt#linkify -> LinkifyIt
   *
   * [linkify-it](https://github.com/markdown-it/linkify-it) instance.
   * Used by [linkify](https://github.com/markdown-it/markdown-it/blob/master/lib/rules_core/linkify.js)
   * rule.
   **/
  this.linkify = new LinkifyIt();

  /**
   * MarkdownIt#validateLink(url) -> Boolean
   *
   * Link validation function. CommonMark allows too much in links. By default
   * we disable `javascript:`, `vbscript:`, `file:` schemas, and almost all `data:...` schemas
   * except some embedded image types.
   *
   * You can change this behaviour:
   *
   * ```javascript
   * var md = require('markdown-it')();
   * // enable everything
   * md.validateLink = function () { return true; }
   * ```
   **/
  this.validateLink = validateLink;

  /**
   * MarkdownIt#normalizeLink(url) -> String
   *
   * Function used to encode link url to a machine-readable format,
   * which includes url-encoding, punycode, etc.
   **/
  this.normalizeLink = normalizeLink;

  /**
   * MarkdownIt#normalizeLinkText(url) -> String
   *
   * Function used to decode link url to a human-readable format`
   **/
  this.normalizeLinkText = normalizeLinkText;


  // Expose utils & helpers for easy acces from plugins

  /**
   * MarkdownIt#utils -> utils
   *
   * Assorted utility functions, useful to write plugins. See details
   * [here](https://github.com/markdown-it/markdown-it/blob/master/lib/common/utils.js).
   **/
  this.utils = utils;

  /**
   * MarkdownIt#helpers -> helpers
   *
   * Link components parser functions, useful to write plugins. See details
   * [here](https://github.com/markdown-it/markdown-it/blob/master/lib/helpers).
   **/
  this.helpers = utils.assign({}, helpers);


  this.options = {};
  this.configure(presetName);

  if (options) { this.set(options); }
}


/** chainable
 * MarkdownIt.set(options)
 *
 * Set parser options (in the same format as in constructor). Probably, you
 * will never need it, but you can change options after constructor call.
 *
 * ##### Example
 *
 * ```javascript
 * var md = require('markdown-it')()
 *             .set({ html: true, breaks: true })
 *             .set({ typographer, true });
 * ```
 *
 * __Note:__ To achieve the best possible performance, don't modify a
 * `markdown-it` instance options on the fly. If you need multiple configurations
 * it's best to create multiple instances and initialize each with separate
 * config.
 **/
MarkdownIt.prototype.set = function (options) {
  utils.assign(this.options, options);
  return this;
};


/** chainable, internal
 * MarkdownIt.configure(presets)
 *
 * Batch load of all options and compenent settings. This is internal method,
 * and you probably will not need it. But if you with - see available presets
 * and data structure [here](https://github.com/markdown-it/markdown-it/tree/master/lib/presets)
 *
 * We strongly recommend to use presets instead of direct config loads. That
 * will give better compatibility with next versions.
 **/
MarkdownIt.prototype.configure = function (presets) {
  var self = this, presetName;

  if (utils.isString(presets)) {
    presetName = presets;
    presets = config[presetName];
    if (!presets) { throw new Error('Wrong `markdown-it` preset "' + presetName + '", check name'); }
  }

  if (!presets) { throw new Error('Wrong `markdown-it` preset, can\'t be empty'); }

  if (presets.options) { self.set(presets.options); }

  if (presets.components) {
    Object.keys(presets.components).forEach(function (name) {
      if (presets.components[name].rules) {
        self[name].ruler.enableOnly(presets.components[name].rules);
      }
      if (presets.components[name].rules2) {
        self[name].ruler2.enableOnly(presets.components[name].rules2);
      }
    });
  }
  return this;
};


/** chainable
 * MarkdownIt.enable(list, ignoreInvalid)
 * - list (String|Array): rule name or list of rule names to enable
 * - ignoreInvalid (Boolean): set `true` to ignore errors when rule not found.
 *
 * Enable list or rules. It will automatically find appropriate components,
 * containing rules with given names. If rule not found, and `ignoreInvalid`
 * not set - throws exception.
 *
 * ##### Example
 *
 * ```javascript
 * var md = require('markdown-it')()
 *             .enable(['sub', 'sup'])
 *             .disable('smartquotes');
 * ```
 **/
MarkdownIt.prototype.enable = function (list, ignoreInvalid) {
  var result = [];

  if (!Array.isArray(list)) { list = [ list ]; }

  [ 'core', 'block', 'inline' ].forEach(function (chain) {
    result = result.concat(this[chain].ruler.enable(list, true));
  }, this);

  result = result.concat(this.inline.ruler2.enable(list, true));

  var missed = list.filter(function (name) { return result.indexOf(name) < 0; });

  if (missed.length && !ignoreInvalid) {
    throw new Error('MarkdownIt. Failed to enable unknown rule(s): ' + missed);
  }

  return this;
};


/** chainable
 * MarkdownIt.disable(list, ignoreInvalid)
 * - list (String|Array): rule name or list of rule names to disable.
 * - ignoreInvalid (Boolean): set `true` to ignore errors when rule not found.
 *
 * The same as [[MarkdownIt.enable]], but turn specified rules off.
 **/
MarkdownIt.prototype.disable = function (list, ignoreInvalid) {
  var result = [];

  if (!Array.isArray(list)) { list = [ list ]; }

  [ 'core', 'block', 'inline' ].forEach(function (chain) {
    result = result.concat(this[chain].ruler.disable(list, true));
  }, this);

  result = result.concat(this.inline.ruler2.disable(list, true));

  var missed = list.filter(function (name) { return result.indexOf(name) < 0; });

  if (missed.length && !ignoreInvalid) {
    throw new Error('MarkdownIt. Failed to disable unknown rule(s): ' + missed);
  }
  return this;
};


/** chainable
 * MarkdownIt.use(plugin, params)
 *
 * Load specified plugin with given params into current parser instance.
 * It's just a sugar to call `plugin(md, params)` with curring.
 *
 * ##### Example
 *
 * ```javascript
 * var iterator = require('markdown-it-for-inline');
 * var md = require('markdown-it')()
 *             .use(iterator, 'foo_replace', 'text', function (tokens, idx) {
 *               tokens[idx].content = tokens[idx].content.replace(/foo/g, 'bar');
 *             });
 * ```
 **/
MarkdownIt.prototype.use = function (plugin /*, params, ... */) {
  var args = [ this ].concat(Array.prototype.slice.call(arguments, 1));
  plugin.apply(plugin, args);
  return this;
};


/** internal
 * MarkdownIt.parse(src, env) -> Array
 * - src (String): source string
 * - env (Object): environment sandbox
 *
 * Parse input string and returns list of block tokens (special token type
 * "inline" will contain list of inline tokens). You should not call this
 * method directly, until you write custom renderer (for example, to produce
 * AST).
 *
 * `env` is used to pass data between "distributed" rules and return additional
 * metadata like reference info, needed for the renderer. It also can be used to
 * inject data in specific cases. Usually, you will be ok to pass `{}`,
 * and then pass updated object to renderer.
 **/
MarkdownIt.prototype.parse = function (src, env) {
  if (typeof src !== 'string') {
    throw new Error('Input data should be a String');
  }

  var state = new this.core.State(src, this, env);

  this.core.process(state);

  return state.tokens;
};


/**
 * MarkdownIt.render(src [, env]) -> String
 * - src (String): source string
 * - env (Object): environment sandbox
 *
 * Render markdown string into html. It does all magic for you :).
 *
 * `env` can be used to inject additional metadata (`{}` by default).
 * But you will not need it with high probability. See also comment
 * in [[MarkdownIt.parse]].
 **/
MarkdownIt.prototype.render = function (src, env) {
  env = env || {};

  return this.renderer.render(this.parse(src, env), this.options, env);
};


/** internal
 * MarkdownIt.parseInline(src, env) -> Array
 * - src (String): source string
 * - env (Object): environment sandbox
 *
 * The same as [[MarkdownIt.parse]] but skip all block rules. It returns the
 * block tokens list with the single `inline` element, containing parsed inline
 * tokens in `children` property. Also updates `env` object.
 **/
MarkdownIt.prototype.parseInline = function (src, env) {
  var state = new this.core.State(src, this, env);

  state.inlineMode = true;
  this.core.process(state);

  return state.tokens;
};


/**
 * MarkdownIt.renderInline(src [, env]) -> String
 * - src (String): source string
 * - env (Object): environment sandbox
 *
 * Similar to [[MarkdownIt.render]] but for single paragraph content. Result
 * will NOT be wrapped into `<p>` tags.
 **/
MarkdownIt.prototype.renderInline = function (src, env) {
  env = env || {};

  return this.renderer.render(this.parseInline(src, env), this.options, env);
};


module.exports = MarkdownIt;


/***/ }),
/* 26 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/** internal
 * class ParserBlock
 *
 * Block-level tokenizer.
 **/



var Ruler           = __webpack_require__(1);


var _rules = [
  // First 2 params - rule name & source. Secondary array - list of rules,
  // which can be terminated by this one.
  [ 'table',      __webpack_require__(44),      [ 'paragraph', 'reference' ] ],
  [ 'code',       __webpack_require__(34) ],
  [ 'fence',      __webpack_require__(35),      [ 'paragraph', 'reference', 'blockquote', 'list' ] ],
  [ 'blockquote', __webpack_require__(33), [ 'paragraph', 'reference', 'list' ] ],
  [ 'hr',         __webpack_require__(37),         [ 'paragraph', 'reference', 'blockquote', 'list' ] ],
  [ 'list',       __webpack_require__(40),       [ 'paragraph', 'reference', 'blockquote' ] ],
  [ 'reference',  __webpack_require__(42) ],
  [ 'heading',    __webpack_require__(36),    [ 'paragraph', 'reference', 'blockquote' ] ],
  [ 'lheading',   __webpack_require__(39) ],
  [ 'html_block', __webpack_require__(38), [ 'paragraph', 'reference', 'blockquote' ] ],
  [ 'paragraph',  __webpack_require__(41) ]
];


/**
 * new ParserBlock()
 **/
function ParserBlock() {
  /**
   * ParserBlock#ruler -> Ruler
   *
   * [[Ruler]] instance. Keep configuration of block rules.
   **/
  this.ruler = new Ruler();

  for (var i = 0; i < _rules.length; i++) {
    this.ruler.push(_rules[i][0], _rules[i][1], { alt: (_rules[i][2] || []).slice() });
  }
}


// Generate tokens for input range
//
ParserBlock.prototype.tokenize = function (state, startLine, endLine) {
  var ok, i,
      rules = this.ruler.getRules(''),
      len = rules.length,
      line = startLine,
      hasEmptyLines = false,
      maxNesting = state.md.options.maxNesting;

  while (line < endLine) {
    state.line = line = state.skipEmptyLines(line);
    if (line >= endLine) { break; }

    // Termination condition for nested calls.
    // Nested calls currently used for blockquotes & lists
    if (state.sCount[line] < state.blkIndent) { break; }

    // If nesting level exceeded - skip tail to the end. That's not ordinary
    // situation and we should not care about content.
    if (state.level >= maxNesting) {
      state.line = endLine;
      break;
    }

    // Try all possible rules.
    // On success, rule should:
    //
    // - update `state.line`
    // - update `state.tokens`
    // - return true

    for (i = 0; i < len; i++) {
      ok = rules[i](state, line, endLine, false);
      if (ok) { break; }
    }

    // set state.tight iff we had an empty line before current tag
    // i.e. latest empty line should not count
    state.tight = !hasEmptyLines;

    // paragraph might "eat" one newline after it in nested lists
    if (state.isEmpty(state.line - 1)) {
      hasEmptyLines = true;
    }

    line = state.line;

    if (line < endLine && state.isEmpty(line)) {
      hasEmptyLines = true;
      line++;
      state.line = line;
    }
  }
};


/**
 * ParserBlock.parse(str, md, env, outTokens)
 *
 * Process input string and push block tokens into `outTokens`
 **/
ParserBlock.prototype.parse = function (src, md, env, outTokens) {
  var state;

  if (!src) { return; }

  state = new this.State(src, md, env, outTokens);

  this.tokenize(state, state.line, state.lineMax);
};


ParserBlock.prototype.State = __webpack_require__(43);


module.exports = ParserBlock;


/***/ }),
/* 27 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/** internal
 * class Core
 *
 * Top-level rules executor. Glues block/inline parsers and does intermediate
 * transformations.
 **/



var Ruler  = __webpack_require__(1);


var _rules = [
  [ 'normalize',      __webpack_require__(48)      ],
  [ 'block',          __webpack_require__(45)          ],
  [ 'inline',         __webpack_require__(46)         ],
  [ 'linkify',        __webpack_require__(47)        ],
  [ 'replacements',   __webpack_require__(49)   ],
  [ 'smartquotes',    __webpack_require__(50)    ]
];


/**
 * new Core()
 **/
function Core() {
  /**
   * Core#ruler -> Ruler
   *
   * [[Ruler]] instance. Keep configuration of core rules.
   **/
  this.ruler = new Ruler();

  for (var i = 0; i < _rules.length; i++) {
    this.ruler.push(_rules[i][0], _rules[i][1]);
  }
}


/**
 * Core.process(state)
 *
 * Executes core chain rules.
 **/
Core.prototype.process = function (state) {
  var i, l, rules;

  rules = this.ruler.getRules('');

  for (i = 0, l = rules.length; i < l; i++) {
    rules[i](state);
  }
};

Core.prototype.State = __webpack_require__(51);


module.exports = Core;


/***/ }),
/* 28 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/** internal
 * class ParserInline
 *
 * Tokenizes paragraph content.
 **/



var Ruler           = __webpack_require__(1);


////////////////////////////////////////////////////////////////////////////////
// Parser rules

var _rules = [
  [ 'text',            __webpack_require__(62) ],
  [ 'newline',         __webpack_require__(60) ],
  [ 'escape',          __webpack_require__(56) ],
  [ 'backticks',       __webpack_require__(53) ],
  [ 'strikethrough',   __webpack_require__(8).tokenize ],
  [ 'emphasis',        __webpack_require__(7).tokenize ],
  [ 'link',            __webpack_require__(59) ],
  [ 'image',           __webpack_require__(58) ],
  [ 'autolink',        __webpack_require__(52) ],
  [ 'html_inline',     __webpack_require__(57) ],
  [ 'entity',          __webpack_require__(55) ]
];

var _rules2 = [
  [ 'balance_pairs',   __webpack_require__(54) ],
  [ 'strikethrough',   __webpack_require__(8).postProcess ],
  [ 'emphasis',        __webpack_require__(7).postProcess ],
  [ 'text_collapse',   __webpack_require__(63) ]
];


/**
 * new ParserInline()
 **/
function ParserInline() {
  var i;

  /**
   * ParserInline#ruler -> Ruler
   *
   * [[Ruler]] instance. Keep configuration of inline rules.
   **/
  this.ruler = new Ruler();

  for (i = 0; i < _rules.length; i++) {
    this.ruler.push(_rules[i][0], _rules[i][1]);
  }

  /**
   * ParserInline#ruler2 -> Ruler
   *
   * [[Ruler]] instance. Second ruler used for post-processing
   * (e.g. in emphasis-like rules).
   **/
  this.ruler2 = new Ruler();

  for (i = 0; i < _rules2.length; i++) {
    this.ruler2.push(_rules2[i][0], _rules2[i][1]);
  }
}


// Skip single token by running all rules in validation mode;
// returns `true` if any rule reported success
//
ParserInline.prototype.skipToken = function (state) {
  var ok, i, pos = state.pos,
      rules = this.ruler.getRules(''),
      len = rules.length,
      maxNesting = state.md.options.maxNesting,
      cache = state.cache;


  if (typeof cache[pos] !== 'undefined') {
    state.pos = cache[pos];
    return;
  }

  if (state.level < maxNesting) {
    for (i = 0; i < len; i++) {
      // Increment state.level and decrement it later to limit recursion.
      // It's harmless to do here, because no tokens are created. But ideally,
      // we'd need a separate private state variable for this purpose.
      //
      state.level++;
      ok = rules[i](state, true);
      state.level--;

      if (ok) { break; }
    }
  } else {
    // Too much nesting, just skip until the end of the paragraph.
    //
    // NOTE: this will cause links to behave incorrectly in the following case,
    //       when an amount of `[` is exactly equal to `maxNesting + 1`:
    //
    //       [[[[[[[[[[[[[[[[[[[[[foo]()
    //
    // TODO: remove this workaround when CM standard will allow nested links
    //       (we can replace it by preventing links from being parsed in
    //       validation mode)
    //
    state.pos = state.posMax;
  }

  if (!ok) { state.pos++; }
  cache[pos] = state.pos;
};


// Generate tokens for input range
//
ParserInline.prototype.tokenize = function (state) {
  var ok, i,
      rules = this.ruler.getRules(''),
      len = rules.length,
      end = state.posMax,
      maxNesting = state.md.options.maxNesting;

  while (state.pos < end) {
    // Try all possible rules.
    // On success, rule should:
    //
    // - update `state.pos`
    // - update `state.tokens`
    // - return true

    if (state.level < maxNesting) {
      for (i = 0; i < len; i++) {
        ok = rules[i](state, false);
        if (ok) { break; }
      }
    }

    if (ok) {
      if (state.pos >= end) { break; }
      continue;
    }

    state.pending += state.src[state.pos++];
  }

  if (state.pending) {
    state.pushPending();
  }
};


/**
 * ParserInline.parse(str, md, env, outTokens)
 *
 * Process input string and push inline tokens into `outTokens`
 **/
ParserInline.prototype.parse = function (str, md, env, outTokens) {
  var i, rules, len;
  var state = new this.State(str, md, env, outTokens);

  this.tokenize(state);

  rules = this.ruler2.getRules('');
  len = rules.length;

  for (i = 0; i < len; i++) {
    rules[i](state);
  }
};


ParserInline.prototype.State = __webpack_require__(61);


module.exports = ParserInline;


/***/ }),
/* 29 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Commonmark default options




module.exports = {
  options: {
    html:         true,         // Enable HTML tags in source
    xhtmlOut:     true,         // Use '/' to close single tags (<br />)
    breaks:       false,        // Convert '\n' in paragraphs into <br>
    langPrefix:   'language-',  // CSS language prefix for fenced blocks
    linkify:      false,        // autoconvert URL-like texts to links

    // Enable some language-neutral replacements + quotes beautification
    typographer:  false,

    // Double + single quotes replacement pairs, when typographer enabled,
    // and smartquotes on. Could be either a String or an Array.
    //
    // For example, you can use 'Â«Â»â€žâ€œ' for Russian, 'â€žâ€œâ€šâ€˜' for German,
    // and ['Â«\xA0', '\xA0Â»', 'â€¹\xA0', '\xA0â€º'] for French (including nbsp).
    quotes: '\u201c\u201d\u2018\u2019', /* â€œâ€â€˜â€™ */

    // Highlighter function. Should return escaped HTML,
    // or '' if the source string is not changed and should be escaped externaly.
    // If result starts with <pre... internal wrapper is skipped.
    //
    // function (/*str, lang*/) { return ''; }
    //
    highlight: null,

    maxNesting:   20            // Internal protection, recursion limit
  },

  components: {

    core: {
      rules: [
        'normalize',
        'block',
        'inline'
      ]
    },

    block: {
      rules: [
        'blockquote',
        'code',
        'fence',
        'heading',
        'hr',
        'html_block',
        'lheading',
        'list',
        'reference',
        'paragraph'
      ]
    },

    inline: {
      rules: [
        'autolink',
        'backticks',
        'emphasis',
        'entity',
        'escape',
        'html_inline',
        'image',
        'link',
        'newline',
        'text'
      ],
      rules2: [
        'balance_pairs',
        'emphasis',
        'text_collapse'
      ]
    }
  }
};


/***/ }),
/* 30 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// markdown-it default options




module.exports = {
  options: {
    html:         false,        // Enable HTML tags in source
    xhtmlOut:     false,        // Use '/' to close single tags (<br />)
    breaks:       false,        // Convert '\n' in paragraphs into <br>
    langPrefix:   'language-',  // CSS language prefix for fenced blocks
    linkify:      false,        // autoconvert URL-like texts to links

    // Enable some language-neutral replacements + quotes beautification
    typographer:  false,

    // Double + single quotes replacement pairs, when typographer enabled,
    // and smartquotes on. Could be either a String or an Array.
    //
    // For example, you can use 'Â«Â»â€žâ€œ' for Russian, 'â€žâ€œâ€šâ€˜' for German,
    // and ['Â«\xA0', '\xA0Â»', 'â€¹\xA0', '\xA0â€º'] for French (including nbsp).
    quotes: '\u201c\u201d\u2018\u2019', /* â€œâ€â€˜â€™ */

    // Highlighter function. Should return escaped HTML,
    // or '' if the source string is not changed and should be escaped externaly.
    // If result starts with <pre... internal wrapper is skipped.
    //
    // function (/*str, lang*/) { return ''; }
    //
    highlight: null,

    maxNesting:   100            // Internal protection, recursion limit
  },

  components: {

    core: {},
    block: {},
    inline: {}
  }
};


/***/ }),
/* 31 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// "Zero" preset, with nothing enabled. Useful for manual configuring of simple
// modes. For example, to parse bold/italic only.




module.exports = {
  options: {
    html:         false,        // Enable HTML tags in source
    xhtmlOut:     false,        // Use '/' to close single tags (<br />)
    breaks:       false,        // Convert '\n' in paragraphs into <br>
    langPrefix:   'language-',  // CSS language prefix for fenced blocks
    linkify:      false,        // autoconvert URL-like texts to links

    // Enable some language-neutral replacements + quotes beautification
    typographer:  false,

    // Double + single quotes replacement pairs, when typographer enabled,
    // and smartquotes on. Could be either a String or an Array.
    //
    // For example, you can use 'Â«Â»â€žâ€œ' for Russian, 'â€žâ€œâ€šâ€˜' for German,
    // and ['Â«\xA0', '\xA0Â»', 'â€¹\xA0', '\xA0â€º'] for French (including nbsp).
    quotes: '\u201c\u201d\u2018\u2019', /* â€œâ€â€˜â€™ */

    // Highlighter function. Should return escaped HTML,
    // or '' if the source string is not changed and should be escaped externaly.
    // If result starts with <pre... internal wrapper is skipped.
    //
    // function (/*str, lang*/) { return ''; }
    //
    highlight: null,

    maxNesting:   20            // Internal protection, recursion limit
  },

  components: {

    core: {
      rules: [
        'normalize',
        'block',
        'inline'
      ]
    },

    block: {
      rules: [
        'paragraph'
      ]
    },

    inline: {
      rules: [
        'text'
      ],
      rules2: [
        'balance_pairs',
        'text_collapse'
      ]
    }
  }
};


/***/ }),
/* 32 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/**
 * class Renderer
 *
 * Generates HTML from parsed token stream. Each instance has independent
 * copy of rules. Those can be rewritten with ease. Also, you can add new
 * rules if you create plugin and adds new token types.
 **/



var assign          = __webpack_require__(0).assign;
var unescapeAll     = __webpack_require__(0).unescapeAll;
var escapeHtml      = __webpack_require__(0).escapeHtml;


////////////////////////////////////////////////////////////////////////////////

var default_rules = {};


default_rules.code_inline = function (tokens, idx, options, env, slf) {
  var token = tokens[idx];

  return  '<code' + slf.renderAttrs(token) + '>' +
          escapeHtml(tokens[idx].content) +
          '</code>';
};


default_rules.code_block = function (tokens, idx, options, env, slf) {
  var token = tokens[idx];

  return  '<pre' + slf.renderAttrs(token) + '><code>' +
          escapeHtml(tokens[idx].content) +
          '</code></pre>\n';
};


default_rules.fence = function (tokens, idx, options, env, slf) {
  var token = tokens[idx],
      info = token.info ? unescapeAll(token.info).trim() : '',
      langName = '',
      highlighted, i, tmpAttrs, tmpToken;

  if (info) {
    langName = info.split(/\s+/g)[0];
  }

  if (options.highlight) {
    highlighted = options.highlight(token.content, langName) || escapeHtml(token.content);
  } else {
    highlighted = escapeHtml(token.content);
  }

  if (highlighted.indexOf('<pre') === 0) {
    return highlighted + '\n';
  }

  // If language exists, inject class gently, without mudofying original token.
  // May be, one day we will add .clone() for token and simplify this part, but
  // now we prefer to keep things local.
  if (info) {
    i        = token.attrIndex('class');
    tmpAttrs = token.attrs ? token.attrs.slice() : [];

    if (i < 0) {
      tmpAttrs.push([ 'class', options.langPrefix + langName ]);
    } else {
      tmpAttrs[i][1] += ' ' + options.langPrefix + langName;
    }

    // Fake token just to render attributes
    tmpToken = {
      attrs: tmpAttrs
    };

    return  '<pre><code' + slf.renderAttrs(tmpToken) + '>'
          + highlighted
          + '</code></pre>\n';
  }


  return  '<pre><code' + slf.renderAttrs(token) + '>'
        + highlighted
        + '</code></pre>\n';
};


default_rules.image = function (tokens, idx, options, env, slf) {
  var token = tokens[idx];

  // "alt" attr MUST be set, even if empty. Because it's mandatory and
  // should be placed on proper position for tests.
  //
  // Replace content with actual value

  token.attrs[token.attrIndex('alt')][1] =
    slf.renderInlineAsText(token.children, options, env);

  return slf.renderToken(tokens, idx, options);
};


default_rules.hardbreak = function (tokens, idx, options /*, env */) {
  return options.xhtmlOut ? '<br />\n' : '<br>\n';
};
default_rules.softbreak = function (tokens, idx, options /*, env */) {
  return options.breaks ? (options.xhtmlOut ? '<br />\n' : '<br>\n') : '\n';
};


default_rules.text = function (tokens, idx /*, options, env */) {
  return escapeHtml(tokens[idx].content);
};


default_rules.html_block = function (tokens, idx /*, options, env */) {
  return tokens[idx].content;
};
default_rules.html_inline = function (tokens, idx /*, options, env */) {
  return tokens[idx].content;
};


/**
 * new Renderer()
 *
 * Creates new [[Renderer]] instance and fill [[Renderer#rules]] with defaults.
 **/
function Renderer() {

  /**
   * Renderer#rules -> Object
   *
   * Contains render rules for tokens. Can be updated and extended.
   *
   * ##### Example
   *
   * ```javascript
   * var md = require('markdown-it')();
   *
   * md.renderer.rules.strong_open  = function () { return '<b>'; };
   * md.renderer.rules.strong_close = function () { return '</b>'; };
   *
   * var result = md.renderInline(...);
   * ```
   *
   * Each rule is called as independed static function with fixed signature:
   *
   * ```javascript
   * function my_token_render(tokens, idx, options, env, renderer) {
   *   // ...
   *   return renderedHTML;
   * }
   * ```
   *
   * See [source code](https://github.com/markdown-it/markdown-it/blob/master/lib/renderer.js)
   * for more details and examples.
   **/
  this.rules = assign({}, default_rules);
}


/**
 * Renderer.renderAttrs(token) -> String
 *
 * Render token attributes to string.
 **/
Renderer.prototype.renderAttrs = function renderAttrs(token) {
  var i, l, result;

  if (!token.attrs) { return ''; }

  result = '';

  for (i = 0, l = token.attrs.length; i < l; i++) {
    result += ' ' + escapeHtml(token.attrs[i][0]) + '="' + escapeHtml(token.attrs[i][1]) + '"';
  }

  return result;
};


/**
 * Renderer.renderToken(tokens, idx, options) -> String
 * - tokens (Array): list of tokens
 * - idx (Numbed): token index to render
 * - options (Object): params of parser instance
 *
 * Default token renderer. Can be overriden by custom function
 * in [[Renderer#rules]].
 **/
Renderer.prototype.renderToken = function renderToken(tokens, idx, options) {
  var nextToken,
      result = '',
      needLf = false,
      token = tokens[idx];

  // Tight list paragraphs
  if (token.hidden) {
    return '';
  }

  // Insert a newline between hidden paragraph and subsequent opening
  // block-level tag.
  //
  // For example, here we should insert a newline before blockquote:
  //  - a
  //    >
  //
  if (token.block && token.nesting !== -1 && idx && tokens[idx - 1].hidden) {
    result += '\n';
  }

  // Add token name, e.g. `<img`
  result += (token.nesting === -1 ? '</' : '<') + token.tag;

  // Encode attributes, e.g. `<img src="foo"`
  result += this.renderAttrs(token);

  // Add a slash for self-closing tags, e.g. `<img src="foo" /`
  if (token.nesting === 0 && options.xhtmlOut) {
    result += ' /';
  }

  // Check if we need to add a newline after this tag
  if (token.block) {
    needLf = true;

    if (token.nesting === 1) {
      if (idx + 1 < tokens.length) {
        nextToken = tokens[idx + 1];

        if (nextToken.type === 'inline' || nextToken.hidden) {
          // Block-level tag containing an inline tag.
          //
          needLf = false;

        } else if (nextToken.nesting === -1 && nextToken.tag === token.tag) {
          // Opening tag + closing tag of the same type. E.g. `<li></li>`.
          //
          needLf = false;
        }
      }
    }
  }

  result += needLf ? '>\n' : '>';

  return result;
};


/**
 * Renderer.renderInline(tokens, options, env) -> String
 * - tokens (Array): list on block tokens to renter
 * - options (Object): params of parser instance
 * - env (Object): additional data from parsed input (references, for example)
 *
 * The same as [[Renderer.render]], but for single token of `inline` type.
 **/
Renderer.prototype.renderInline = function (tokens, options, env) {
  var type,
      result = '',
      rules = this.rules;

  for (var i = 0, len = tokens.length; i < len; i++) {
    type = tokens[i].type;

    if (typeof rules[type] !== 'undefined') {
      result += rules[type](tokens, i, options, env, this);
    } else {
      result += this.renderToken(tokens, i, options);
    }
  }

  return result;
};


/** internal
 * Renderer.renderInlineAsText(tokens, options, env) -> String
 * - tokens (Array): list on block tokens to renter
 * - options (Object): params of parser instance
 * - env (Object): additional data from parsed input (references, for example)
 *
 * Special kludge for image `alt` attributes to conform CommonMark spec.
 * Don't try to use it! Spec requires to show `alt` content with stripped markup,
 * instead of simple escaping.
 **/
Renderer.prototype.renderInlineAsText = function (tokens, options, env) {
  var result = '';

  for (var i = 0, len = tokens.length; i < len; i++) {
    if (tokens[i].type === 'text') {
      result += tokens[i].content;
    } else if (tokens[i].type === 'image') {
      result += this.renderInlineAsText(tokens[i].children, options, env);
    }
  }

  return result;
};


/**
 * Renderer.render(tokens, options, env) -> String
 * - tokens (Array): list on block tokens to renter
 * - options (Object): params of parser instance
 * - env (Object): additional data from parsed input (references, for example)
 *
 * Takes token stream and generates HTML. Probably, you will never need to call
 * this method directly.
 **/
Renderer.prototype.render = function (tokens, options, env) {
  var i, len, type,
      result = '',
      rules = this.rules;

  for (i = 0, len = tokens.length; i < len; i++) {
    type = tokens[i].type;

    if (type === 'inline') {
      result += this.renderInline(tokens[i].children, options, env);
    } else if (typeof rules[type] !== 'undefined') {
      result += rules[tokens[i].type](tokens, i, options, env, this);
    } else {
      result += this.renderToken(tokens, i, options, env);
    }
  }

  return result;
};

module.exports = Renderer;


/***/ }),
/* 33 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Block quotes



var isSpace = __webpack_require__(0).isSpace;


module.exports = function blockquote(state, startLine, endLine, silent) {
  var adjustTab,
      ch,
      i,
      initial,
      isOutdented,
      l,
      lastLineEmpty,
      lines,
      nextLine,
      offset,
      oldBMarks,
      oldBSCount,
      oldIndent,
      oldParentType,
      oldSCount,
      oldTShift,
      spaceAfterMarker,
      terminate,
      terminatorRules,
      token,
      oldLineMax = state.lineMax,
      pos = state.bMarks[startLine] + state.tShift[startLine],
      max = state.eMarks[startLine];

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  // check the block quote marker
  if (state.src.charCodeAt(pos++) !== 0x3E/* > */) { return false; }

  // we know that it's going to be a valid blockquote,
  // so no point trying to find the end of it in silent mode
  if (silent) { return true; }

  // skip spaces after ">" and re-calculate offset
  initial = offset = state.sCount[startLine] + pos - (state.bMarks[startLine] + state.tShift[startLine]);

  // skip one optional space after '>'
  if (state.src.charCodeAt(pos) === 0x20 /* space */) {
    // ' >   test '
    //     ^ -- position start of line here:
    pos++;
    initial++;
    offset++;
    adjustTab = false;
    spaceAfterMarker = true;
  } else if (state.src.charCodeAt(pos) === 0x09 /* tab */) {
    spaceAfterMarker = true;

    if ((state.bsCount[startLine] + offset) % 4 === 3) {
      // '  >\t  test '
      //       ^ -- position start of line here (tab has width===1)
      pos++;
      initial++;
      offset++;
      adjustTab = false;
    } else {
      // ' >\t  test '
      //    ^ -- position start of line here + shift bsCount slightly
      //         to make extra space appear
      adjustTab = true;
    }
  } else {
    spaceAfterMarker = false;
  }

  oldBMarks = [ state.bMarks[startLine] ];
  state.bMarks[startLine] = pos;

  while (pos < max) {
    ch = state.src.charCodeAt(pos);

    if (isSpace(ch)) {
      if (ch === 0x09) {
        offset += 4 - (offset + state.bsCount[startLine] + (adjustTab ? 1 : 0)) % 4;
      } else {
        offset++;
      }
    } else {
      break;
    }

    pos++;
  }

  oldBSCount = [ state.bsCount[startLine] ];
  state.bsCount[startLine] = state.sCount[startLine] + 1 + (spaceAfterMarker ? 1 : 0);

  lastLineEmpty = pos >= max;

  oldSCount = [ state.sCount[startLine] ];
  state.sCount[startLine] = offset - initial;

  oldTShift = [ state.tShift[startLine] ];
  state.tShift[startLine] = pos - state.bMarks[startLine];

  terminatorRules = state.md.block.ruler.getRules('blockquote');

  oldParentType = state.parentType;
  state.parentType = 'blockquote';

  // Search the end of the block
  //
  // Block ends with either:
  //  1. an empty line outside:
  //     ```
  //     > test
  //
  //     ```
  //  2. an empty line inside:
  //     ```
  //     >
  //     test
  //     ```
  //  3. another tag:
  //     ```
  //     > test
  //      - - -
  //     ```
  for (nextLine = startLine + 1; nextLine < endLine; nextLine++) {
    // check if it's outdented, i.e. it's inside list item and indented
    // less than said list item:
    //
    // ```
    // 1. anything
    //    > current blockquote
    // 2. checking this line
    // ```
    isOutdented = state.sCount[nextLine] < state.blkIndent;

    pos = state.bMarks[nextLine] + state.tShift[nextLine];
    max = state.eMarks[nextLine];

    if (pos >= max) {
      // Case 1: line is not inside the blockquote, and this line is empty.
      break;
    }

    if (state.src.charCodeAt(pos++) === 0x3E/* > */ && !isOutdented) {
      // This line is inside the blockquote.

      // skip spaces after ">" and re-calculate offset
      initial = offset = state.sCount[nextLine] + pos - (state.bMarks[nextLine] + state.tShift[nextLine]);

      // skip one optional space after '>'
      if (state.src.charCodeAt(pos) === 0x20 /* space */) {
        // ' >   test '
        //     ^ -- position start of line here:
        pos++;
        initial++;
        offset++;
        adjustTab = false;
        spaceAfterMarker = true;
      } else if (state.src.charCodeAt(pos) === 0x09 /* tab */) {
        spaceAfterMarker = true;

        if ((state.bsCount[nextLine] + offset) % 4 === 3) {
          // '  >\t  test '
          //       ^ -- position start of line here (tab has width===1)
          pos++;
          initial++;
          offset++;
          adjustTab = false;
        } else {
          // ' >\t  test '
          //    ^ -- position start of line here + shift bsCount slightly
          //         to make extra space appear
          adjustTab = true;
        }
      } else {
        spaceAfterMarker = false;
      }

      oldBMarks.push(state.bMarks[nextLine]);
      state.bMarks[nextLine] = pos;

      while (pos < max) {
        ch = state.src.charCodeAt(pos);

        if (isSpace(ch)) {
          if (ch === 0x09) {
            offset += 4 - (offset + state.bsCount[nextLine] + (adjustTab ? 1 : 0)) % 4;
          } else {
            offset++;
          }
        } else {
          break;
        }

        pos++;
      }

      lastLineEmpty = pos >= max;

      oldBSCount.push(state.bsCount[nextLine]);
      state.bsCount[nextLine] = state.sCount[nextLine] + 1 + (spaceAfterMarker ? 1 : 0);

      oldSCount.push(state.sCount[nextLine]);
      state.sCount[nextLine] = offset - initial;

      oldTShift.push(state.tShift[nextLine]);
      state.tShift[nextLine] = pos - state.bMarks[nextLine];
      continue;
    }

    // Case 2: line is not inside the blockquote, and the last line was empty.
    if (lastLineEmpty) { break; }

    // Case 3: another tag found.
    terminate = false;
    for (i = 0, l = terminatorRules.length; i < l; i++) {
      if (terminatorRules[i](state, nextLine, endLine, true)) {
        terminate = true;
        break;
      }
    }

    if (terminate) {
      // Quirk to enforce "hard termination mode" for paragraphs;
      // normally if you call `tokenize(state, startLine, nextLine)`,
      // paragraphs will look below nextLine for paragraph continuation,
      // but if blockquote is terminated by another tag, they shouldn't
      state.lineMax = nextLine;

      if (state.blkIndent !== 0) {
        // state.blkIndent was non-zero, we now set it to zero,
        // so we need to re-calculate all offsets to appear as
        // if indent wasn't changed
        oldBMarks.push(state.bMarks[nextLine]);
        oldBSCount.push(state.bsCount[nextLine]);
        oldTShift.push(state.tShift[nextLine]);
        oldSCount.push(state.sCount[nextLine]);
        state.sCount[nextLine] -= state.blkIndent;
      }

      break;
    }

    if (isOutdented) break;

    oldBMarks.push(state.bMarks[nextLine]);
    oldBSCount.push(state.bsCount[nextLine]);
    oldTShift.push(state.tShift[nextLine]);
    oldSCount.push(state.sCount[nextLine]);

    // A negative indentation means that this is a paragraph continuation
    //
    state.sCount[nextLine] = -1;
  }

  oldIndent = state.blkIndent;
  state.blkIndent = 0;

  token        = state.push('blockquote_open', 'blockquote', 1);
  token.markup = '>';
  token.map    = lines = [ startLine, 0 ];

  state.md.block.tokenize(state, startLine, nextLine);

  token        = state.push('blockquote_close', 'blockquote', -1);
  token.markup = '>';

  state.lineMax = oldLineMax;
  state.parentType = oldParentType;
  lines[1] = state.line;

  // Restore original tShift; this might not be necessary since the parser
  // has already been here, but just to make sure we can do that.
  for (i = 0; i < oldTShift.length; i++) {
    state.bMarks[i + startLine] = oldBMarks[i];
    state.tShift[i + startLine] = oldTShift[i];
    state.sCount[i + startLine] = oldSCount[i];
    state.bsCount[i + startLine] = oldBSCount[i];
  }
  state.blkIndent = oldIndent;

  return true;
};


/***/ }),
/* 34 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Code block (4 spaces padded)




module.exports = function code(state, startLine, endLine/*, silent*/) {
  var nextLine, last, token;

  if (state.sCount[startLine] - state.blkIndent < 4) { return false; }

  last = nextLine = startLine + 1;

  while (nextLine < endLine) {
    if (state.isEmpty(nextLine)) {
      nextLine++;
      continue;
    }

    if (state.sCount[nextLine] - state.blkIndent >= 4) {
      nextLine++;
      last = nextLine;
      continue;
    }
    break;
  }

  state.line = last;

  token         = state.push('code_block', 'code', 0);
  token.content = state.getLines(startLine, last, 4 + state.blkIndent, true);
  token.map     = [ startLine, state.line ];

  return true;
};


/***/ }),
/* 35 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// fences (``` lang, ~~~ lang)




module.exports = function fence(state, startLine, endLine, silent) {
  var marker, len, params, nextLine, mem, token, markup,
      haveEndMarker = false,
      pos = state.bMarks[startLine] + state.tShift[startLine],
      max = state.eMarks[startLine];

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  if (pos + 3 > max) { return false; }

  marker = state.src.charCodeAt(pos);

  if (marker !== 0x7E/* ~ */ && marker !== 0x60 /* ` */) {
    return false;
  }

  // scan marker length
  mem = pos;
  pos = state.skipChars(pos, marker);

  len = pos - mem;

  if (len < 3) { return false; }

  markup = state.src.slice(mem, pos);
  params = state.src.slice(pos, max);

  if (params.indexOf(String.fromCharCode(marker)) >= 0) { return false; }

  // Since start is found, we can report success here in validation mode
  if (silent) { return true; }

  // search end of block
  nextLine = startLine;

  for (;;) {
    nextLine++;
    if (nextLine >= endLine) {
      // unclosed block should be autoclosed by end of document.
      // also block seems to be autoclosed by end of parent
      break;
    }

    pos = mem = state.bMarks[nextLine] + state.tShift[nextLine];
    max = state.eMarks[nextLine];

    if (pos < max && state.sCount[nextLine] < state.blkIndent) {
      // non-empty line with negative indent should stop the list:
      // - ```
      //  test
      break;
    }

    if (state.src.charCodeAt(pos) !== marker) { continue; }

    if (state.sCount[nextLine] - state.blkIndent >= 4) {
      // closing fence should be indented less than 4 spaces
      continue;
    }

    pos = state.skipChars(pos, marker);

    // closing code fence must be at least as long as the opening one
    if (pos - mem < len) { continue; }

    // make sure tail has spaces only
    pos = state.skipSpaces(pos);

    if (pos < max) { continue; }

    haveEndMarker = true;
    // found!
    break;
  }

  // If a fence has heading spaces, they should be removed from its inner block
  len = state.sCount[startLine];

  state.line = nextLine + (haveEndMarker ? 1 : 0);

  token         = state.push('fence', 'code', 0);
  token.info    = params;
  token.content = state.getLines(startLine + 1, nextLine, len, true);
  token.markup  = markup;
  token.map     = [ startLine, state.line ];

  return true;
};


/***/ }),
/* 36 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// heading (#, ##, ...)



var isSpace = __webpack_require__(0).isSpace;


module.exports = function heading(state, startLine, endLine, silent) {
  var ch, level, tmp, token,
      pos = state.bMarks[startLine] + state.tShift[startLine],
      max = state.eMarks[startLine];

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  ch  = state.src.charCodeAt(pos);

  if (ch !== 0x23/* # */ || pos >= max) { return false; }

  // count heading level
  level = 1;
  ch = state.src.charCodeAt(++pos);
  while (ch === 0x23/* # */ && pos < max && level <= 6) {
    level++;
    ch = state.src.charCodeAt(++pos);
  }

  if (level > 6 || (pos < max && !isSpace(ch))) { return false; }

  if (silent) { return true; }

  // Let's cut tails like '    ###  ' from the end of string

  max = state.skipSpacesBack(max, pos);
  tmp = state.skipCharsBack(max, 0x23, pos); // #
  if (tmp > pos && isSpace(state.src.charCodeAt(tmp - 1))) {
    max = tmp;
  }

  state.line = startLine + 1;

  token        = state.push('heading_open', 'h' + String(level), 1);
  token.markup = '########'.slice(0, level);
  token.map    = [ startLine, state.line ];

  token          = state.push('inline', '', 0);
  token.content  = state.src.slice(pos, max).trim();
  token.map      = [ startLine, state.line ];
  token.children = [];

  token        = state.push('heading_close', 'h' + String(level), -1);
  token.markup = '########'.slice(0, level);

  return true;
};


/***/ }),
/* 37 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Horizontal rule



var isSpace = __webpack_require__(0).isSpace;


module.exports = function hr(state, startLine, endLine, silent) {
  var marker, cnt, ch, token,
      pos = state.bMarks[startLine] + state.tShift[startLine],
      max = state.eMarks[startLine];

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  marker = state.src.charCodeAt(pos++);

  // Check hr marker
  if (marker !== 0x2A/* * */ &&
      marker !== 0x2D/* - */ &&
      marker !== 0x5F/* _ */) {
    return false;
  }

  // markers can be mixed with spaces, but there should be at least 3 of them

  cnt = 1;
  while (pos < max) {
    ch = state.src.charCodeAt(pos++);
    if (ch !== marker && !isSpace(ch)) { return false; }
    if (ch === marker) { cnt++; }
  }

  if (cnt < 3) { return false; }

  if (silent) { return true; }

  state.line = startLine + 1;

  token        = state.push('hr', 'hr', 0);
  token.map    = [ startLine, state.line ];
  token.markup = Array(cnt + 1).join(String.fromCharCode(marker));

  return true;
};


/***/ }),
/* 38 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// HTML block




var block_names = __webpack_require__(20);
var HTML_OPEN_CLOSE_TAG_RE = __webpack_require__(6).HTML_OPEN_CLOSE_TAG_RE;

// An array of opening and corresponding closing sequences for html tags,
// last argument defines whether it can terminate a paragraph or not
//
var HTML_SEQUENCES = [
  [ /^<(script|pre|style)(?=(\s|>|$))/i, /<\/(script|pre|style)>/i, true ],
  [ /^<!--/,        /-->/,   true ],
  [ /^<\?/,         /\?>/,   true ],
  [ /^<![A-Z]/,     />/,     true ],
  [ /^<!\[CDATA\[/, /\]\]>/, true ],
  [ new RegExp('^</?(' + block_names.join('|') + ')(?=(\\s|/?>|$))', 'i'), /^$/, true ],
  [ new RegExp(HTML_OPEN_CLOSE_TAG_RE.source + '\\s*$'),  /^$/, false ]
];


module.exports = function html_block(state, startLine, endLine, silent) {
  var i, nextLine, token, lineText,
      pos = state.bMarks[startLine] + state.tShift[startLine],
      max = state.eMarks[startLine];

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  if (!state.md.options.html) { return false; }

  if (state.src.charCodeAt(pos) !== 0x3C/* < */) { return false; }

  lineText = state.src.slice(pos, max);

  for (i = 0; i < HTML_SEQUENCES.length; i++) {
    if (HTML_SEQUENCES[i][0].test(lineText)) { break; }
  }

  if (i === HTML_SEQUENCES.length) { return false; }

  if (silent) {
    // true if this sequence can be a terminator, false otherwise
    return HTML_SEQUENCES[i][2];
  }

  nextLine = startLine + 1;

  // If we are here - we detected HTML block.
  // Let's roll down till block end.
  if (!HTML_SEQUENCES[i][1].test(lineText)) {
    for (; nextLine < endLine; nextLine++) {
      if (state.sCount[nextLine] < state.blkIndent) { break; }

      pos = state.bMarks[nextLine] + state.tShift[nextLine];
      max = state.eMarks[nextLine];
      lineText = state.src.slice(pos, max);

      if (HTML_SEQUENCES[i][1].test(lineText)) {
        if (lineText.length !== 0) { nextLine++; }
        break;
      }
    }
  }

  state.line = nextLine;

  token         = state.push('html_block', '', 0);
  token.map     = [ startLine, nextLine ];
  token.content = state.getLines(startLine, nextLine, state.blkIndent, true);

  return true;
};


/***/ }),
/* 39 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// lheading (---, ===)




module.exports = function lheading(state, startLine, endLine/*, silent*/) {
  var content, terminate, i, l, token, pos, max, level, marker,
      nextLine = startLine + 1, oldParentType,
      terminatorRules = state.md.block.ruler.getRules('paragraph');

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  oldParentType = state.parentType;
  state.parentType = 'paragraph'; // use paragraph to match terminatorRules

  // jump line-by-line until empty one or EOF
  for (; nextLine < endLine && !state.isEmpty(nextLine); nextLine++) {
    // this would be a code block normally, but after paragraph
    // it's considered a lazy continuation regardless of what's there
    if (state.sCount[nextLine] - state.blkIndent > 3) { continue; }

    //
    // Check for underline in setext header
    //
    if (state.sCount[nextLine] >= state.blkIndent) {
      pos = state.bMarks[nextLine] + state.tShift[nextLine];
      max = state.eMarks[nextLine];

      if (pos < max) {
        marker = state.src.charCodeAt(pos);

        if (marker === 0x2D/* - */ || marker === 0x3D/* = */) {
          pos = state.skipChars(pos, marker);
          pos = state.skipSpaces(pos);

          if (pos >= max) {
            level = (marker === 0x3D/* = */ ? 1 : 2);
            break;
          }
        }
      }
    }

    // quirk for blockquotes, this line should already be checked by that rule
    if (state.sCount[nextLine] < 0) { continue; }

    // Some tags can terminate paragraph without empty line.
    terminate = false;
    for (i = 0, l = terminatorRules.length; i < l; i++) {
      if (terminatorRules[i](state, nextLine, endLine, true)) {
        terminate = true;
        break;
      }
    }
    if (terminate) { break; }
  }

  if (!level) {
    // Didn't find valid underline
    return false;
  }

  content = state.getLines(startLine, nextLine, state.blkIndent, false).trim();

  state.line = nextLine + 1;

  token          = state.push('heading_open', 'h' + String(level), 1);
  token.markup   = String.fromCharCode(marker);
  token.map      = [ startLine, state.line ];

  token          = state.push('inline', '', 0);
  token.content  = content;
  token.map      = [ startLine, state.line - 1 ];
  token.children = [];

  token          = state.push('heading_close', 'h' + String(level), -1);
  token.markup   = String.fromCharCode(marker);

  state.parentType = oldParentType;

  return true;
};


/***/ }),
/* 40 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Lists



var isSpace = __webpack_require__(0).isSpace;


// Search `[-+*][\n ]`, returns next pos arter marker on success
// or -1 on fail.
function skipBulletListMarker(state, startLine) {
  var marker, pos, max, ch;

  pos = state.bMarks[startLine] + state.tShift[startLine];
  max = state.eMarks[startLine];

  marker = state.src.charCodeAt(pos++);
  // Check bullet
  if (marker !== 0x2A/* * */ &&
      marker !== 0x2D/* - */ &&
      marker !== 0x2B/* + */) {
    return -1;
  }

  if (pos < max) {
    ch = state.src.charCodeAt(pos);

    if (!isSpace(ch)) {
      // " -test " - is not a list item
      return -1;
    }
  }

  return pos;
}

// Search `\d+[.)][\n ]`, returns next pos arter marker on success
// or -1 on fail.
function skipOrderedListMarker(state, startLine) {
  var ch,
      start = state.bMarks[startLine] + state.tShift[startLine],
      pos = start,
      max = state.eMarks[startLine];

  // List marker should have at least 2 chars (digit + dot)
  if (pos + 1 >= max) { return -1; }

  ch = state.src.charCodeAt(pos++);

  if (ch < 0x30/* 0 */ || ch > 0x39/* 9 */) { return -1; }

  for (;;) {
    // EOL -> fail
    if (pos >= max) { return -1; }

    ch = state.src.charCodeAt(pos++);

    if (ch >= 0x30/* 0 */ && ch <= 0x39/* 9 */) {

      // List marker should have no more than 9 digits
      // (prevents integer overflow in browsers)
      if (pos - start >= 10) { return -1; }

      continue;
    }

    // found valid marker
    if (ch === 0x29/* ) */ || ch === 0x2e/* . */) {
      break;
    }

    return -1;
  }


  if (pos < max) {
    ch = state.src.charCodeAt(pos);

    if (!isSpace(ch)) {
      // " 1.test " - is not a list item
      return -1;
    }
  }
  return pos;
}

function markTightParagraphs(state, idx) {
  var i, l,
      level = state.level + 2;

  for (i = idx + 2, l = state.tokens.length - 2; i < l; i++) {
    if (state.tokens[i].level === level && state.tokens[i].type === 'paragraph_open') {
      state.tokens[i + 2].hidden = true;
      state.tokens[i].hidden = true;
      i += 2;
    }
  }
}


module.exports = function list(state, startLine, endLine, silent) {
  var ch,
      contentStart,
      i,
      indent,
      indentAfterMarker,
      initial,
      isOrdered,
      itemLines,
      l,
      listLines,
      listTokIdx,
      markerCharCode,
      markerValue,
      max,
      nextLine,
      offset,
      oldIndent,
      oldLIndent,
      oldParentType,
      oldTShift,
      oldTight,
      pos,
      posAfterMarker,
      prevEmptyEnd,
      start,
      terminate,
      terminatorRules,
      token,
      isTerminatingParagraph = false,
      tight = true;

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  // limit conditions when list can interrupt
  // a paragraph (validation mode only)
  if (silent && state.parentType === 'paragraph') {
    // Next list item should still terminate previous list item;
    //
    // This code can fail if plugins use blkIndent as well as lists,
    // but I hope the spec gets fixed long before that happens.
    //
    if (state.tShift[startLine] >= state.blkIndent) {
      isTerminatingParagraph = true;
    }
  }

  // Detect list type and position after marker
  if ((posAfterMarker = skipOrderedListMarker(state, startLine)) >= 0) {
    isOrdered = true;
    start = state.bMarks[startLine] + state.tShift[startLine];
    markerValue = Number(state.src.substr(start, posAfterMarker - start - 1));

    // If we're starting a new ordered list right after
    // a paragraph, it should start with 1.
    if (isTerminatingParagraph && markerValue !== 1) return false;

  } else if ((posAfterMarker = skipBulletListMarker(state, startLine)) >= 0) {
    isOrdered = false;

  } else {
    return false;
  }

  // If we're starting a new unordered list right after
  // a paragraph, first line should not be empty.
  if (isTerminatingParagraph) {
    if (state.skipSpaces(posAfterMarker) >= state.eMarks[startLine]) return false;
  }

  // We should terminate list on style change. Remember first one to compare.
  markerCharCode = state.src.charCodeAt(posAfterMarker - 1);

  // For validation mode we can terminate immediately
  if (silent) { return true; }

  // Start list
  listTokIdx = state.tokens.length;

  if (isOrdered) {
    token       = state.push('ordered_list_open', 'ol', 1);
    if (markerValue !== 1) {
      token.attrs = [ [ 'start', markerValue ] ];
    }

  } else {
    token       = state.push('bullet_list_open', 'ul', 1);
  }

  token.map    = listLines = [ startLine, 0 ];
  token.markup = String.fromCharCode(markerCharCode);

  //
  // Iterate list items
  //

  nextLine = startLine;
  prevEmptyEnd = false;
  terminatorRules = state.md.block.ruler.getRules('list');

  oldParentType = state.parentType;
  state.parentType = 'list';

  while (nextLine < endLine) {
    pos = posAfterMarker;
    max = state.eMarks[nextLine];

    initial = offset = state.sCount[nextLine] + posAfterMarker - (state.bMarks[startLine] + state.tShift[startLine]);

    while (pos < max) {
      ch = state.src.charCodeAt(pos);

      if (isSpace(ch)) {
        if (ch === 0x09) {
          offset += 4 - (offset + state.bsCount[nextLine]) % 4;
        } else {
          offset++;
        }
      } else {
        break;
      }

      pos++;
    }

    contentStart = pos;

    if (contentStart >= max) {
      // trimming space in "-    \n  3" case, indent is 1 here
      indentAfterMarker = 1;
    } else {
      indentAfterMarker = offset - initial;
    }

    // If we have more than 4 spaces, the indent is 1
    // (the rest is just indented code block)
    if (indentAfterMarker > 4) { indentAfterMarker = 1; }

    // "  -  test"
    //  ^^^^^ - calculating total length of this thing
    indent = initial + indentAfterMarker;

    // Run subparser & write tokens
    token        = state.push('list_item_open', 'li', 1);
    token.markup = String.fromCharCode(markerCharCode);
    token.map    = itemLines = [ startLine, 0 ];

    oldIndent = state.blkIndent;
    oldTight = state.tight;
    oldTShift = state.tShift[startLine];
    oldLIndent = state.sCount[startLine];
    state.blkIndent = indent;
    state.tight = true;
    state.tShift[startLine] = contentStart - state.bMarks[startLine];
    state.sCount[startLine] = offset;

    if (contentStart >= max && state.isEmpty(startLine + 1)) {
      // workaround for this case
      // (list item is empty, list terminates before "foo"):
      // ~~~~~~~~
      //   -
      //
      //     foo
      // ~~~~~~~~
      state.line = Math.min(state.line + 2, endLine);
    } else {
      state.md.block.tokenize(state, startLine, endLine, true);
    }

    // If any of list item is tight, mark list as tight
    if (!state.tight || prevEmptyEnd) {
      tight = false;
    }
    // Item become loose if finish with empty line,
    // but we should filter last element, because it means list finish
    prevEmptyEnd = (state.line - startLine) > 1 && state.isEmpty(state.line - 1);

    state.blkIndent = oldIndent;
    state.tShift[startLine] = oldTShift;
    state.sCount[startLine] = oldLIndent;
    state.tight = oldTight;

    token        = state.push('list_item_close', 'li', -1);
    token.markup = String.fromCharCode(markerCharCode);

    nextLine = startLine = state.line;
    itemLines[1] = nextLine;
    contentStart = state.bMarks[startLine];

    if (nextLine >= endLine) { break; }

    //
    // Try to check if list is terminated or continued.
    //
    if (state.sCount[nextLine] < state.blkIndent) { break; }

    // fail if terminating block found
    terminate = false;
    for (i = 0, l = terminatorRules.length; i < l; i++) {
      if (terminatorRules[i](state, nextLine, endLine, true)) {
        terminate = true;
        break;
      }
    }
    if (terminate) { break; }

    // fail if list has another type
    if (isOrdered) {
      posAfterMarker = skipOrderedListMarker(state, nextLine);
      if (posAfterMarker < 0) { break; }
    } else {
      posAfterMarker = skipBulletListMarker(state, nextLine);
      if (posAfterMarker < 0) { break; }
    }

    if (markerCharCode !== state.src.charCodeAt(posAfterMarker - 1)) { break; }
  }

  // Finilize list
  if (isOrdered) {
    token = state.push('ordered_list_close', 'ol', -1);
  } else {
    token = state.push('bullet_list_close', 'ul', -1);
  }
  token.markup = String.fromCharCode(markerCharCode);

  listLines[1] = nextLine;
  state.line = nextLine;

  state.parentType = oldParentType;

  // mark paragraphs tight if needed
  if (tight) {
    markTightParagraphs(state, listTokIdx);
  }

  return true;
};


/***/ }),
/* 41 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Paragraph




module.exports = function paragraph(state, startLine/*, endLine*/) {
  var content, terminate, i, l, token, oldParentType,
      nextLine = startLine + 1,
      terminatorRules = state.md.block.ruler.getRules('paragraph'),
      endLine = state.lineMax;

  oldParentType = state.parentType;
  state.parentType = 'paragraph';

  // jump line-by-line until empty one or EOF
  for (; nextLine < endLine && !state.isEmpty(nextLine); nextLine++) {
    // this would be a code block normally, but after paragraph
    // it's considered a lazy continuation regardless of what's there
    if (state.sCount[nextLine] - state.blkIndent > 3) { continue; }

    // quirk for blockquotes, this line should already be checked by that rule
    if (state.sCount[nextLine] < 0) { continue; }

    // Some tags can terminate paragraph without empty line.
    terminate = false;
    for (i = 0, l = terminatorRules.length; i < l; i++) {
      if (terminatorRules[i](state, nextLine, endLine, true)) {
        terminate = true;
        break;
      }
    }
    if (terminate) { break; }
  }

  content = state.getLines(startLine, nextLine, state.blkIndent, false).trim();

  state.line = nextLine;

  token          = state.push('paragraph_open', 'p', 1);
  token.map      = [ startLine, state.line ];

  token          = state.push('inline', '', 0);
  token.content  = content;
  token.map      = [ startLine, state.line ];
  token.children = [];

  token          = state.push('paragraph_close', 'p', -1);

  state.parentType = oldParentType;

  return true;
};


/***/ }),
/* 42 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";



var normalizeReference   = __webpack_require__(0).normalizeReference;
var isSpace              = __webpack_require__(0).isSpace;


module.exports = function reference(state, startLine, _endLine, silent) {
  var ch,
      destEndPos,
      destEndLineNo,
      endLine,
      href,
      i,
      l,
      label,
      labelEnd,
      oldParentType,
      res,
      start,
      str,
      terminate,
      terminatorRules,
      title,
      lines = 0,
      pos = state.bMarks[startLine] + state.tShift[startLine],
      max = state.eMarks[startLine],
      nextLine = startLine + 1;

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }

  if (state.src.charCodeAt(pos) !== 0x5B/* [ */) { return false; }

  // Simple check to quickly interrupt scan on [link](url) at the start of line.
  // Can be useful on practice: https://github.com/markdown-it/markdown-it/issues/54
  while (++pos < max) {
    if (state.src.charCodeAt(pos) === 0x5D /* ] */ &&
        state.src.charCodeAt(pos - 1) !== 0x5C/* \ */) {
      if (pos + 1 === max) { return false; }
      if (state.src.charCodeAt(pos + 1) !== 0x3A/* : */) { return false; }
      break;
    }
  }

  endLine = state.lineMax;

  // jump line-by-line until empty one or EOF
  terminatorRules = state.md.block.ruler.getRules('reference');

  oldParentType = state.parentType;
  state.parentType = 'reference';

  for (; nextLine < endLine && !state.isEmpty(nextLine); nextLine++) {
    // this would be a code block normally, but after paragraph
    // it's considered a lazy continuation regardless of what's there
    if (state.sCount[nextLine] - state.blkIndent > 3) { continue; }

    // quirk for blockquotes, this line should already be checked by that rule
    if (state.sCount[nextLine] < 0) { continue; }

    // Some tags can terminate paragraph without empty line.
    terminate = false;
    for (i = 0, l = terminatorRules.length; i < l; i++) {
      if (terminatorRules[i](state, nextLine, endLine, true)) {
        terminate = true;
        break;
      }
    }
    if (terminate) { break; }
  }

  str = state.getLines(startLine, nextLine, state.blkIndent, false).trim();
  max = str.length;

  for (pos = 1; pos < max; pos++) {
    ch = str.charCodeAt(pos);
    if (ch === 0x5B /* [ */) {
      return false;
    } else if (ch === 0x5D /* ] */) {
      labelEnd = pos;
      break;
    } else if (ch === 0x0A /* \n */) {
      lines++;
    } else if (ch === 0x5C /* \ */) {
      pos++;
      if (pos < max && str.charCodeAt(pos) === 0x0A) {
        lines++;
      }
    }
  }

  if (labelEnd < 0 || str.charCodeAt(labelEnd + 1) !== 0x3A/* : */) { return false; }

  // [label]:   destination   'title'
  //         ^^^ skip optional whitespace here
  for (pos = labelEnd + 2; pos < max; pos++) {
    ch = str.charCodeAt(pos);
    if (ch === 0x0A) {
      lines++;
    } else if (isSpace(ch)) {
      /*eslint no-empty:0*/
    } else {
      break;
    }
  }

  // [label]:   destination   'title'
  //            ^^^^^^^^^^^ parse this
  res = state.md.helpers.parseLinkDestination(str, pos, max);
  if (!res.ok) { return false; }

  href = state.md.normalizeLink(res.str);
  if (!state.md.validateLink(href)) { return false; }

  pos = res.pos;
  lines += res.lines;

  // save cursor state, we could require to rollback later
  destEndPos = pos;
  destEndLineNo = lines;

  // [label]:   destination   'title'
  //                       ^^^ skipping those spaces
  start = pos;
  for (; pos < max; pos++) {
    ch = str.charCodeAt(pos);
    if (ch === 0x0A) {
      lines++;
    } else if (isSpace(ch)) {
      /*eslint no-empty:0*/
    } else {
      break;
    }
  }

  // [label]:   destination   'title'
  //                          ^^^^^^^ parse this
  res = state.md.helpers.parseLinkTitle(str, pos, max);
  if (pos < max && start !== pos && res.ok) {
    title = res.str;
    pos = res.pos;
    lines += res.lines;
  } else {
    title = '';
    pos = destEndPos;
    lines = destEndLineNo;
  }

  // skip trailing spaces until the rest of the line
  while (pos < max) {
    ch = str.charCodeAt(pos);
    if (!isSpace(ch)) { break; }
    pos++;
  }

  if (pos < max && str.charCodeAt(pos) !== 0x0A) {
    if (title) {
      // garbage at the end of the line after title,
      // but it could still be a valid reference if we roll back
      title = '';
      pos = destEndPos;
      lines = destEndLineNo;
      while (pos < max) {
        ch = str.charCodeAt(pos);
        if (!isSpace(ch)) { break; }
        pos++;
      }
    }
  }

  if (pos < max && str.charCodeAt(pos) !== 0x0A) {
    // garbage at the end of the line
    return false;
  }

  label = normalizeReference(str.slice(1, labelEnd));
  if (!label) {
    // CommonMark 0.20 disallows empty labels
    return false;
  }

  // Reference can not terminate anything. This check is for safety only.
  /*istanbul ignore if*/
  if (silent) { return true; }

  if (typeof state.env.references === 'undefined') {
    state.env.references = {};
  }
  if (typeof state.env.references[label] === 'undefined') {
    state.env.references[label] = { title: title, href: href };
  }

  state.parentType = oldParentType;

  state.line = startLine + lines + 1;
  return true;
};


/***/ }),
/* 43 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Parser state class



var Token = __webpack_require__(2);
var isSpace = __webpack_require__(0).isSpace;


function StateBlock(src, md, env, tokens) {
  var ch, s, start, pos, len, indent, offset, indent_found;

  this.src = src;

  // link to parser instance
  this.md     = md;

  this.env = env;

  //
  // Internal state vartiables
  //

  this.tokens = tokens;

  this.bMarks = [];  // line begin offsets for fast jumps
  this.eMarks = [];  // line end offsets for fast jumps
  this.tShift = [];  // offsets of the first non-space characters (tabs not expanded)
  this.sCount = [];  // indents for each line (tabs expanded)

  // An amount of virtual spaces (tabs expanded) between beginning
  // of each line (bMarks) and real beginning of that line.
  //
  // It exists only as a hack because blockquotes override bMarks
  // losing information in the process.
  //
  // It's used only when expanding tabs, you can think about it as
  // an initial tab length, e.g. bsCount=21 applied to string `\t123`
  // means first tab should be expanded to 4-21%4 === 3 spaces.
  //
  this.bsCount = [];

  // block parser variables
  this.blkIndent  = 0; // required block content indent
                       // (for example, if we are in list)
  this.line       = 0; // line index in src
  this.lineMax    = 0; // lines count
  this.tight      = false;  // loose/tight mode for lists
  this.ddIndent   = -1; // indent of the current dd block (-1 if there isn't any)

  // can be 'blockquote', 'list', 'root', 'paragraph' or 'reference'
  // used in lists to determine if they interrupt a paragraph
  this.parentType = 'root';

  this.level = 0;

  // renderer
  this.result = '';

  // Create caches
  // Generate markers.
  s = this.src;
  indent_found = false;

  for (start = pos = indent = offset = 0, len = s.length; pos < len; pos++) {
    ch = s.charCodeAt(pos);

    if (!indent_found) {
      if (isSpace(ch)) {
        indent++;

        if (ch === 0x09) {
          offset += 4 - offset % 4;
        } else {
          offset++;
        }
        continue;
      } else {
        indent_found = true;
      }
    }

    if (ch === 0x0A || pos === len - 1) {
      if (ch !== 0x0A) { pos++; }
      this.bMarks.push(start);
      this.eMarks.push(pos);
      this.tShift.push(indent);
      this.sCount.push(offset);
      this.bsCount.push(0);

      indent_found = false;
      indent = 0;
      offset = 0;
      start = pos + 1;
    }
  }

  // Push fake entry to simplify cache bounds checks
  this.bMarks.push(s.length);
  this.eMarks.push(s.length);
  this.tShift.push(0);
  this.sCount.push(0);
  this.bsCount.push(0);

  this.lineMax = this.bMarks.length - 1; // don't count last fake line
}

// Push new token to "stream".
//
StateBlock.prototype.push = function (type, tag, nesting) {
  var token = new Token(type, tag, nesting);
  token.block = true;

  if (nesting < 0) { this.level--; }
  token.level = this.level;
  if (nesting > 0) { this.level++; }

  this.tokens.push(token);
  return token;
};

StateBlock.prototype.isEmpty = function isEmpty(line) {
  return this.bMarks[line] + this.tShift[line] >= this.eMarks[line];
};

StateBlock.prototype.skipEmptyLines = function skipEmptyLines(from) {
  for (var max = this.lineMax; from < max; from++) {
    if (this.bMarks[from] + this.tShift[from] < this.eMarks[from]) {
      break;
    }
  }
  return from;
};

// Skip spaces from given position.
StateBlock.prototype.skipSpaces = function skipSpaces(pos) {
  var ch;

  for (var max = this.src.length; pos < max; pos++) {
    ch = this.src.charCodeAt(pos);
    if (!isSpace(ch)) { break; }
  }
  return pos;
};

// Skip spaces from given position in reverse.
StateBlock.prototype.skipSpacesBack = function skipSpacesBack(pos, min) {
  if (pos <= min) { return pos; }

  while (pos > min) {
    if (!isSpace(this.src.charCodeAt(--pos))) { return pos + 1; }
  }
  return pos;
};

// Skip char codes from given position
StateBlock.prototype.skipChars = function skipChars(pos, code) {
  for (var max = this.src.length; pos < max; pos++) {
    if (this.src.charCodeAt(pos) !== code) { break; }
  }
  return pos;
};

// Skip char codes reverse from given position - 1
StateBlock.prototype.skipCharsBack = function skipCharsBack(pos, code, min) {
  if (pos <= min) { return pos; }

  while (pos > min) {
    if (code !== this.src.charCodeAt(--pos)) { return pos + 1; }
  }
  return pos;
};

// cut lines range from source.
StateBlock.prototype.getLines = function getLines(begin, end, indent, keepLastLF) {
  var i, lineIndent, ch, first, last, queue, lineStart,
      line = begin;

  if (begin >= end) {
    return '';
  }

  queue = new Array(end - begin);

  for (i = 0; line < end; line++, i++) {
    lineIndent = 0;
    lineStart = first = this.bMarks[line];

    if (line + 1 < end || keepLastLF) {
      // No need for bounds check because we have fake entry on tail.
      last = this.eMarks[line] + 1;
    } else {
      last = this.eMarks[line];
    }

    while (first < last && lineIndent < indent) {
      ch = this.src.charCodeAt(first);

      if (isSpace(ch)) {
        if (ch === 0x09) {
          lineIndent += 4 - (lineIndent + this.bsCount[line]) % 4;
        } else {
          lineIndent++;
        }
      } else if (first - lineStart < this.tShift[line]) {
        // patched tShift masked characters to look like spaces (blockquotes, list markers)
        lineIndent++;
      } else {
        break;
      }

      first++;
    }

    if (lineIndent > indent) {
      // partially expanding tabs in code blocks, e.g '\t\tfoobar'
      // with indent=2 becomes '  \tfoobar'
      queue[i] = new Array(lineIndent - indent + 1).join(' ') + this.src.slice(first, last);
    } else {
      queue[i] = this.src.slice(first, last);
    }
  }

  return queue.join('');
};

// re-export Token class to use in block rules
StateBlock.prototype.Token = Token;


module.exports = StateBlock;


/***/ }),
/* 44 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// GFM table, non-standard



var isSpace = __webpack_require__(0).isSpace;


function getLine(state, line) {
  var pos = state.bMarks[line] + state.blkIndent,
      max = state.eMarks[line];

  return state.src.substr(pos, max - pos);
}

function escapedSplit(str) {
  var result = [],
      pos = 0,
      max = str.length,
      ch,
      escapes = 0,
      lastPos = 0,
      backTicked = false,
      lastBackTick = 0;

  ch  = str.charCodeAt(pos);

  while (pos < max) {
    if (ch === 0x60/* ` */) {
      if (backTicked) {
        // make \` close code sequence, but not open it;
        // the reason is: `\` is correct code block
        backTicked = false;
        lastBackTick = pos;
      } else if (escapes % 2 === 0) {
        backTicked = true;
        lastBackTick = pos;
      }
    } else if (ch === 0x7c/* | */ && (escapes % 2 === 0) && !backTicked) {
      result.push(str.substring(lastPos, pos));
      lastPos = pos + 1;
    }

    if (ch === 0x5c/* \ */) {
      escapes++;
    } else {
      escapes = 0;
    }

    pos++;

    // If there was an un-closed backtick, go back to just after
    // the last backtick, but as if it was a normal character
    if (pos === max && backTicked) {
      backTicked = false;
      pos = lastBackTick + 1;
    }

    ch = str.charCodeAt(pos);
  }

  result.push(str.substring(lastPos));

  return result;
}


module.exports = function table(state, startLine, endLine, silent) {
  var ch, lineText, pos, i, nextLine, columns, columnCount, token,
      aligns, t, tableLines, tbodyLines;

  // should have at least two lines
  if (startLine + 2 > endLine) { return false; }

  nextLine = startLine + 1;

  if (state.sCount[nextLine] < state.blkIndent) { return false; }

  // if it's indented more than 3 spaces, it should be a code block
  if (state.sCount[nextLine] - state.blkIndent >= 4) { return false; }

  // first character of the second line should be '|', '-', ':',
  // and no other characters are allowed but spaces;
  // basically, this is the equivalent of /^[-:|][-:|\s]*$/ regexp

  pos = state.bMarks[nextLine] + state.tShift[nextLine];
  if (pos >= state.eMarks[nextLine]) { return false; }

  ch = state.src.charCodeAt(pos++);
  if (ch !== 0x7C/* | */ && ch !== 0x2D/* - */ && ch !== 0x3A/* : */) { return false; }

  while (pos < state.eMarks[nextLine]) {
    ch = state.src.charCodeAt(pos);

    if (ch !== 0x7C/* | */ && ch !== 0x2D/* - */ && ch !== 0x3A/* : */ && !isSpace(ch)) { return false; }

    pos++;
  }

  lineText = getLine(state, startLine + 1);

  columns = lineText.split('|');
  aligns = [];
  for (i = 0; i < columns.length; i++) {
    t = columns[i].trim();
    if (!t) {
      // allow empty columns before and after table, but not in between columns;
      // e.g. allow ` |---| `, disallow ` ---||--- `
      if (i === 0 || i === columns.length - 1) {
        continue;
      } else {
        return false;
      }
    }

    if (!/^:?-+:?$/.test(t)) { return false; }
    if (t.charCodeAt(t.length - 1) === 0x3A/* : */) {
      aligns.push(t.charCodeAt(0) === 0x3A/* : */ ? 'center' : 'right');
    } else if (t.charCodeAt(0) === 0x3A/* : */) {
      aligns.push('left');
    } else {
      aligns.push('');
    }
  }

  lineText = getLine(state, startLine).trim();
  if (lineText.indexOf('|') === -1) { return false; }
  if (state.sCount[startLine] - state.blkIndent >= 4) { return false; }
  columns = escapedSplit(lineText.replace(/^\||\|$/g, ''));

  // header row will define an amount of columns in the entire table,
  // and align row shouldn't be smaller than that (the rest of the rows can)
  columnCount = columns.length;
  if (columnCount > aligns.length) { return false; }

  if (silent) { return true; }

  token     = state.push('table_open', 'table', 1);
  token.map = tableLines = [ startLine, 0 ];

  token     = state.push('thead_open', 'thead', 1);
  token.map = [ startLine, startLine + 1 ];

  token     = state.push('tr_open', 'tr', 1);
  token.map = [ startLine, startLine + 1 ];

  for (i = 0; i < columns.length; i++) {
    token          = state.push('th_open', 'th', 1);
    token.map      = [ startLine, startLine + 1 ];
    if (aligns[i]) {
      token.attrs  = [ [ 'style', 'text-align:' + aligns[i] ] ];
    }

    token          = state.push('inline', '', 0);
    token.content  = columns[i].trim();
    token.map      = [ startLine, startLine + 1 ];
    token.children = [];

    token          = state.push('th_close', 'th', -1);
  }

  token     = state.push('tr_close', 'tr', -1);
  token     = state.push('thead_close', 'thead', -1);

  token     = state.push('tbody_open', 'tbody', 1);
  token.map = tbodyLines = [ startLine + 2, 0 ];

  for (nextLine = startLine + 2; nextLine < endLine; nextLine++) {
    if (state.sCount[nextLine] < state.blkIndent) { break; }

    lineText = getLine(state, nextLine).trim();
    if (lineText.indexOf('|') === -1) { break; }
    if (state.sCount[nextLine] - state.blkIndent >= 4) { break; }
    columns = escapedSplit(lineText.replace(/^\||\|$/g, ''));

    token = state.push('tr_open', 'tr', 1);
    for (i = 0; i < columnCount; i++) {
      token          = state.push('td_open', 'td', 1);
      if (aligns[i]) {
        token.attrs  = [ [ 'style', 'text-align:' + aligns[i] ] ];
      }

      token          = state.push('inline', '', 0);
      token.content  = columns[i] ? columns[i].trim() : '';
      token.children = [];

      token          = state.push('td_close', 'td', -1);
    }
    token = state.push('tr_close', 'tr', -1);
  }
  token = state.push('tbody_close', 'tbody', -1);
  token = state.push('table_close', 'table', -1);

  tableLines[1] = tbodyLines[1] = nextLine;
  state.line = nextLine;
  return true;
};


/***/ }),
/* 45 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";



module.exports = function block(state) {
  var token;

  if (state.inlineMode) {
    token          = new state.Token('inline', '', 0);
    token.content  = state.src;
    token.map      = [ 0, 1 ];
    token.children = [];
    state.tokens.push(token);
  } else {
    state.md.block.parse(state.src, state.md, state.env, state.tokens);
  }
};


/***/ }),
/* 46 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function inline(state) {
  var tokens = state.tokens, tok, i, l;

  // Parse inlines
  for (i = 0, l = tokens.length; i < l; i++) {
    tok = tokens[i];
    if (tok.type === 'inline') {
      state.md.inline.parse(tok.content, state.md, state.env, tok.children);
    }
  }
};


/***/ }),
/* 47 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Replace link-like texts with link nodes.
//
// Currently restricted by `md.validateLink()` to http/https/ftp
//



var arrayReplaceAt = __webpack_require__(0).arrayReplaceAt;


function isLinkOpen(str) {
  return /^<a[>\s]/i.test(str);
}
function isLinkClose(str) {
  return /^<\/a\s*>/i.test(str);
}


module.exports = function linkify(state) {
  var i, j, l, tokens, token, currentToken, nodes, ln, text, pos, lastPos,
      level, htmlLinkLevel, url, fullUrl, urlText,
      blockTokens = state.tokens,
      links;

  if (!state.md.options.linkify) { return; }

  for (j = 0, l = blockTokens.length; j < l; j++) {
    if (blockTokens[j].type !== 'inline' ||
        !state.md.linkify.pretest(blockTokens[j].content)) {
      continue;
    }

    tokens = blockTokens[j].children;

    htmlLinkLevel = 0;

    // We scan from the end, to keep position when new tags added.
    // Use reversed logic in links start/end match
    for (i = tokens.length - 1; i >= 0; i--) {
      currentToken = tokens[i];

      // Skip content of markdown links
      if (currentToken.type === 'link_close') {
        i--;
        while (tokens[i].level !== currentToken.level && tokens[i].type !== 'link_open') {
          i--;
        }
        continue;
      }

      // Skip content of html tag links
      if (currentToken.type === 'html_inline') {
        if (isLinkOpen(currentToken.content) && htmlLinkLevel > 0) {
          htmlLinkLevel--;
        }
        if (isLinkClose(currentToken.content)) {
          htmlLinkLevel++;
        }
      }
      if (htmlLinkLevel > 0) { continue; }

      if (currentToken.type === 'text' && state.md.linkify.test(currentToken.content)) {

        text = currentToken.content;
        links = state.md.linkify.match(text);

        // Now split string to nodes
        nodes = [];
        level = currentToken.level;
        lastPos = 0;

        for (ln = 0; ln < links.length; ln++) {

          url = links[ln].url;
          fullUrl = state.md.normalizeLink(url);
          if (!state.md.validateLink(fullUrl)) { continue; }

          urlText = links[ln].text;

          // Linkifier might send raw hostnames like "example.com", where url
          // starts with domain name. So we prepend http:// in those cases,
          // and remove it afterwards.
          //
          if (!links[ln].schema) {
            urlText = state.md.normalizeLinkText('http://' + urlText).replace(/^http:\/\//, '');
          } else if (links[ln].schema === 'mailto:' && !/^mailto:/i.test(urlText)) {
            urlText = state.md.normalizeLinkText('mailto:' + urlText).replace(/^mailto:/, '');
          } else {
            urlText = state.md.normalizeLinkText(urlText);
          }

          pos = links[ln].index;

          if (pos > lastPos) {
            token         = new state.Token('text', '', 0);
            token.content = text.slice(lastPos, pos);
            token.level   = level;
            nodes.push(token);
          }

          token         = new state.Token('link_open', 'a', 1);
          token.attrs   = [ [ 'href', fullUrl ] ];
          token.level   = level++;
          token.markup  = 'linkify';
          token.info    = 'auto';
          nodes.push(token);

          token         = new state.Token('text', '', 0);
          token.content = urlText;
          token.level   = level;
          nodes.push(token);

          token         = new state.Token('link_close', 'a', -1);
          token.level   = --level;
          token.markup  = 'linkify';
          token.info    = 'auto';
          nodes.push(token);

          lastPos = links[ln].lastIndex;
        }
        if (lastPos < text.length) {
          token         = new state.Token('text', '', 0);
          token.content = text.slice(lastPos);
          token.level   = level;
          nodes.push(token);
        }

        // replace current node
        blockTokens[j].children = tokens = arrayReplaceAt(tokens, i, nodes);
      }
    }
  }
};


/***/ }),
/* 48 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Normalize input string




var NEWLINES_RE  = /\r[\n\u0085]?|[\u2424\u2028\u0085]/g;
var NULL_RE      = /\u0000/g;


module.exports = function inline(state) {
  var str;

  // Normalize newlines
  str = state.src.replace(NEWLINES_RE, '\n');

  // Replace NULL characters
  str = str.replace(NULL_RE, '\uFFFD');

  state.src = str;
};


/***/ }),
/* 49 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Simple typographyc replacements
//
// (c) (C) â†’ Â©
// (tm) (TM) â†’ â„¢
// (r) (R) â†’ Â®
// +- â†’ Â±
// (p) (P) -> Â§
// ... â†’ â€¦ (also ?.... â†’ ?.., !.... â†’ !..)
// ???????? â†’ ???, !!!!! â†’ !!!, `,,` â†’ `,`
// -- â†’ &ndash;, --- â†’ &mdash;
//


// TODO:
// - fractionals 1/2, 1/4, 3/4 -> Â½, Â¼, Â¾
// - miltiplication 2 x 4 -> 2 Ã— 4

var RARE_RE = /\+-|\.\.|\?\?\?\?|!!!!|,,|--/;

// Workaround for phantomjs - need regex without /g flag,
// or root check will fail every second time
var SCOPED_ABBR_TEST_RE = /\((c|tm|r|p)\)/i;

var SCOPED_ABBR_RE = /\((c|tm|r|p)\)/ig;
var SCOPED_ABBR = {
  c: 'Â©',
  r: 'Â®',
  p: 'Â§',
  tm: 'â„¢'
};

function replaceFn(match, name) {
  return SCOPED_ABBR[name.toLowerCase()];
}

function replace_scoped(inlineTokens) {
  var i, token, inside_autolink = 0;

  for (i = inlineTokens.length - 1; i >= 0; i--) {
    token = inlineTokens[i];

    if (token.type === 'text' && !inside_autolink) {
      token.content = token.content.replace(SCOPED_ABBR_RE, replaceFn);
    }

    if (token.type === 'link_open' && token.info === 'auto') {
      inside_autolink--;
    }

    if (token.type === 'link_close' && token.info === 'auto') {
      inside_autolink++;
    }
  }
}

function replace_rare(inlineTokens) {
  var i, token, inside_autolink = 0;

  for (i = inlineTokens.length - 1; i >= 0; i--) {
    token = inlineTokens[i];

    if (token.type === 'text' && !inside_autolink) {
      if (RARE_RE.test(token.content)) {
        token.content = token.content
                    .replace(/\+-/g, 'Â±')
                    // .., ..., ....... -> â€¦
                    // but ?..... & !..... -> ?.. & !..
                    .replace(/\.{2,}/g, 'â€¦').replace(/([?!])â€¦/g, '$1..')
                    .replace(/([?!]){4,}/g, '$1$1$1').replace(/,{2,}/g, ',')
                    // em-dash
                    .replace(/(^|[^-])---([^-]|$)/mg, '$1\u2014$2')
                    // en-dash
                    .replace(/(^|\s)--(\s|$)/mg, '$1\u2013$2')
                    .replace(/(^|[^-\s])--([^-\s]|$)/mg, '$1\u2013$2');
      }
    }

    if (token.type === 'link_open' && token.info === 'auto') {
      inside_autolink--;
    }

    if (token.type === 'link_close' && token.info === 'auto') {
      inside_autolink++;
    }
  }
}


module.exports = function replace(state) {
  var blkIdx;

  if (!state.md.options.typographer) { return; }

  for (blkIdx = state.tokens.length - 1; blkIdx >= 0; blkIdx--) {

    if (state.tokens[blkIdx].type !== 'inline') { continue; }

    if (SCOPED_ABBR_TEST_RE.test(state.tokens[blkIdx].content)) {
      replace_scoped(state.tokens[blkIdx].children);
    }

    if (RARE_RE.test(state.tokens[blkIdx].content)) {
      replace_rare(state.tokens[blkIdx].children);
    }

  }
};


/***/ }),
/* 50 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Convert straight quotation marks to typographic ones
//



var isWhiteSpace   = __webpack_require__(0).isWhiteSpace;
var isPunctChar    = __webpack_require__(0).isPunctChar;
var isMdAsciiPunct = __webpack_require__(0).isMdAsciiPunct;

var QUOTE_TEST_RE = /['"]/;
var QUOTE_RE = /['"]/g;
var APOSTROPHE = '\u2019'; /* â€™ */


function replaceAt(str, index, ch) {
  return str.substr(0, index) + ch + str.substr(index + 1);
}

function process_inlines(tokens, state) {
  var i, token, text, t, pos, max, thisLevel, item, lastChar, nextChar,
      isLastPunctChar, isNextPunctChar, isLastWhiteSpace, isNextWhiteSpace,
      canOpen, canClose, j, isSingle, stack, openQuote, closeQuote;

  stack = [];

  for (i = 0; i < tokens.length; i++) {
    token = tokens[i];

    thisLevel = tokens[i].level;

    for (j = stack.length - 1; j >= 0; j--) {
      if (stack[j].level <= thisLevel) { break; }
    }
    stack.length = j + 1;

    if (token.type !== 'text') { continue; }

    text = token.content;
    pos = 0;
    max = text.length;

    /*eslint no-labels:0,block-scoped-var:0*/
    OUTER:
    while (pos < max) {
      QUOTE_RE.lastIndex = pos;
      t = QUOTE_RE.exec(text);
      if (!t) { break; }

      canOpen = canClose = true;
      pos = t.index + 1;
      isSingle = (t[0] === "'");

      // Find previous character,
      // default to space if it's the beginning of the line
      //
      lastChar = 0x20;

      if (t.index - 1 >= 0) {
        lastChar = text.charCodeAt(t.index - 1);
      } else {
        for (j = i - 1; j >= 0; j--) {
          if (tokens[j].type !== 'text') { continue; }

          lastChar = tokens[j].content.charCodeAt(tokens[j].content.length - 1);
          break;
        }
      }

      // Find next character,
      // default to space if it's the end of the line
      //
      nextChar = 0x20;

      if (pos < max) {
        nextChar = text.charCodeAt(pos);
      } else {
        for (j = i + 1; j < tokens.length; j++) {
          if (tokens[j].type !== 'text') { continue; }

          nextChar = tokens[j].content.charCodeAt(0);
          break;
        }
      }

      isLastPunctChar = isMdAsciiPunct(lastChar) || isPunctChar(String.fromCharCode(lastChar));
      isNextPunctChar = isMdAsciiPunct(nextChar) || isPunctChar(String.fromCharCode(nextChar));

      isLastWhiteSpace = isWhiteSpace(lastChar);
      isNextWhiteSpace = isWhiteSpace(nextChar);

      if (isNextWhiteSpace) {
        canOpen = false;
      } else if (isNextPunctChar) {
        if (!(isLastWhiteSpace || isLastPunctChar)) {
          canOpen = false;
        }
      }

      if (isLastWhiteSpace) {
        canClose = false;
      } else if (isLastPunctChar) {
        if (!(isNextWhiteSpace || isNextPunctChar)) {
          canClose = false;
        }
      }

      if (nextChar === 0x22 /* " */ && t[0] === '"') {
        if (lastChar >= 0x30 /* 0 */ && lastChar <= 0x39 /* 9 */) {
          // special case: 1"" - count first quote as an inch
          canClose = canOpen = false;
        }
      }

      if (canOpen && canClose) {
        // treat this as the middle of the word
        canOpen = false;
        canClose = isNextPunctChar;
      }

      if (!canOpen && !canClose) {
        // middle of word
        if (isSingle) {
          token.content = replaceAt(token.content, t.index, APOSTROPHE);
        }
        continue;
      }

      if (canClose) {
        // this could be a closing quote, rewind the stack to get a match
        for (j = stack.length - 1; j >= 0; j--) {
          item = stack[j];
          if (stack[j].level < thisLevel) { break; }
          if (item.single === isSingle && stack[j].level === thisLevel) {
            item = stack[j];

            if (isSingle) {
              openQuote = state.md.options.quotes[2];
              closeQuote = state.md.options.quotes[3];
            } else {
              openQuote = state.md.options.quotes[0];
              closeQuote = state.md.options.quotes[1];
            }

            // replace token.content *before* tokens[item.token].content,
            // because, if they are pointing at the same token, replaceAt
            // could mess up indices when quote length != 1
            token.content = replaceAt(token.content, t.index, closeQuote);
            tokens[item.token].content = replaceAt(
              tokens[item.token].content, item.pos, openQuote);

            pos += closeQuote.length - 1;
            if (item.token === i) { pos += openQuote.length - 1; }

            text = token.content;
            max = text.length;

            stack.length = j;
            continue OUTER;
          }
        }
      }

      if (canOpen) {
        stack.push({
          token: i,
          pos: t.index,
          single: isSingle,
          level: thisLevel
        });
      } else if (canClose && isSingle) {
        token.content = replaceAt(token.content, t.index, APOSTROPHE);
      }
    }
  }
}


module.exports = function smartquotes(state) {
  /*eslint max-depth:0*/
  var blkIdx;

  if (!state.md.options.typographer) { return; }

  for (blkIdx = state.tokens.length - 1; blkIdx >= 0; blkIdx--) {

    if (state.tokens[blkIdx].type !== 'inline' ||
        !QUOTE_TEST_RE.test(state.tokens[blkIdx].content)) {
      continue;
    }

    process_inlines(state.tokens[blkIdx].children, state);
  }
};


/***/ }),
/* 51 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Core state object
//


var Token = __webpack_require__(2);


function StateCore(src, md, env) {
  this.src = src;
  this.env = env;
  this.tokens = [];
  this.inlineMode = false;
  this.md = md; // link to parser instance
}

// re-export Token class to use in core rules
StateCore.prototype.Token = Token;


module.exports = StateCore;


/***/ }),
/* 52 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Process autolinks '<protocol:...>'




/*eslint max-len:0*/
var EMAIL_RE    = /^<([a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*)>/;
var AUTOLINK_RE = /^<([a-zA-Z][a-zA-Z0-9+.\-]{1,31}):([^<>\x00-\x20]*)>/;


module.exports = function autolink(state, silent) {
  var tail, linkMatch, emailMatch, url, fullUrl, token,
      pos = state.pos;

  if (state.src.charCodeAt(pos) !== 0x3C/* < */) { return false; }

  tail = state.src.slice(pos);

  if (tail.indexOf('>') < 0) { return false; }

  if (AUTOLINK_RE.test(tail)) {
    linkMatch = tail.match(AUTOLINK_RE);

    url = linkMatch[0].slice(1, -1);
    fullUrl = state.md.normalizeLink(url);
    if (!state.md.validateLink(fullUrl)) { return false; }

    if (!silent) {
      token         = state.push('link_open', 'a', 1);
      token.attrs   = [ [ 'href', fullUrl ] ];
      token.markup  = 'autolink';
      token.info    = 'auto';

      token         = state.push('text', '', 0);
      token.content = state.md.normalizeLinkText(url);

      token         = state.push('link_close', 'a', -1);
      token.markup  = 'autolink';
      token.info    = 'auto';
    }

    state.pos += linkMatch[0].length;
    return true;
  }

  if (EMAIL_RE.test(tail)) {
    emailMatch = tail.match(EMAIL_RE);

    url = emailMatch[0].slice(1, -1);
    fullUrl = state.md.normalizeLink('mailto:' + url);
    if (!state.md.validateLink(fullUrl)) { return false; }

    if (!silent) {
      token         = state.push('link_open', 'a', 1);
      token.attrs   = [ [ 'href', fullUrl ] ];
      token.markup  = 'autolink';
      token.info    = 'auto';

      token         = state.push('text', '', 0);
      token.content = state.md.normalizeLinkText(url);

      token         = state.push('link_close', 'a', -1);
      token.markup  = 'autolink';
      token.info    = 'auto';
    }

    state.pos += emailMatch[0].length;
    return true;
  }

  return false;
};


/***/ }),
/* 53 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Parse backticks



module.exports = function backtick(state, silent) {
  var start, max, marker, matchStart, matchEnd, token,
      pos = state.pos,
      ch = state.src.charCodeAt(pos);

  if (ch !== 0x60/* ` */) { return false; }

  start = pos;
  pos++;
  max = state.posMax;

  while (pos < max && state.src.charCodeAt(pos) === 0x60/* ` */) { pos++; }

  marker = state.src.slice(start, pos);

  matchStart = matchEnd = pos;

  while ((matchStart = state.src.indexOf('`', matchEnd)) !== -1) {
    matchEnd = matchStart + 1;

    while (matchEnd < max && state.src.charCodeAt(matchEnd) === 0x60/* ` */) { matchEnd++; }

    if (matchEnd - matchStart === marker.length) {
      if (!silent) {
        token         = state.push('code_inline', 'code', 0);
        token.markup  = marker;
        token.content = state.src.slice(pos, matchStart)
                                 .replace(/[ \n]+/g, ' ')
                                 .trim();
      }
      state.pos = matchEnd;
      return true;
    }
  }

  if (!silent) { state.pending += marker; }
  state.pos += marker.length;
  return true;
};


/***/ }),
/* 54 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// For each opening emphasis-like marker find a matching closing one
//



module.exports = function link_pairs(state) {
  var i, j, lastDelim, currDelim,
      delimiters = state.delimiters,
      max = state.delimiters.length;

  for (i = 0; i < max; i++) {
    lastDelim = delimiters[i];

    if (!lastDelim.close) { continue; }

    j = i - lastDelim.jump - 1;

    while (j >= 0) {
      currDelim = delimiters[j];

      if (currDelim.open &&
          currDelim.marker === lastDelim.marker &&
          currDelim.end < 0 &&
          currDelim.level === lastDelim.level) {

        // typeofs are for backward compatibility with plugins
        var odd_match = (currDelim.close || lastDelim.open) &&
                        typeof currDelim.length !== 'undefined' &&
                        typeof lastDelim.length !== 'undefined' &&
                        (currDelim.length + lastDelim.length) % 3 === 0;

        if (!odd_match) {
          lastDelim.jump = i - j;
          lastDelim.open = false;
          currDelim.end  = i;
          currDelim.jump = 0;
          break;
        }
      }

      j -= currDelim.jump + 1;
    }
  }
};


/***/ }),
/* 55 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Process html entity - &#123;, &#xAF;, &quot;, ...



var entities          = __webpack_require__(5);
var has               = __webpack_require__(0).has;
var isValidEntityCode = __webpack_require__(0).isValidEntityCode;
var fromCodePoint     = __webpack_require__(0).fromCodePoint;


var DIGITAL_RE = /^&#((?:x[a-f0-9]{1,8}|[0-9]{1,8}));/i;
var NAMED_RE   = /^&([a-z][a-z0-9]{1,31});/i;


module.exports = function entity(state, silent) {
  var ch, code, match, pos = state.pos, max = state.posMax;

  if (state.src.charCodeAt(pos) !== 0x26/* & */) { return false; }

  if (pos + 1 < max) {
    ch = state.src.charCodeAt(pos + 1);

    if (ch === 0x23 /* # */) {
      match = state.src.slice(pos).match(DIGITAL_RE);
      if (match) {
        if (!silent) {
          code = match[1][0].toLowerCase() === 'x' ? parseInt(match[1].slice(1), 16) : parseInt(match[1], 10);
          state.pending += isValidEntityCode(code) ? fromCodePoint(code) : fromCodePoint(0xFFFD);
        }
        state.pos += match[0].length;
        return true;
      }
    } else {
      match = state.src.slice(pos).match(NAMED_RE);
      if (match) {
        if (has(entities, match[1])) {
          if (!silent) { state.pending += entities[match[1]]; }
          state.pos += match[0].length;
          return true;
        }
      }
    }
  }

  if (!silent) { state.pending += '&'; }
  state.pos++;
  return true;
};


/***/ }),
/* 56 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Proceess escaped chars and hardbreaks



var isSpace = __webpack_require__(0).isSpace;

var ESCAPED = [];

for (var i = 0; i < 256; i++) { ESCAPED.push(0); }

'\\!"#$%&\'()*+,./:;<=>?@[]^_`{|}~-'
  .split('').forEach(function (ch) { ESCAPED[ch.charCodeAt(0)] = 1; });


module.exports = function escape(state, silent) {
  var ch, pos = state.pos, max = state.posMax;

  if (state.src.charCodeAt(pos) !== 0x5C/* \ */) { return false; }

  pos++;

  if (pos < max) {
    ch = state.src.charCodeAt(pos);

    if (ch < 256 && ESCAPED[ch] !== 0) {
      if (!silent) { state.pending += state.src[pos]; }
      state.pos += 2;
      return true;
    }

    if (ch === 0x0A) {
      if (!silent) {
        state.push('hardbreak', 'br', 0);
      }

      pos++;
      // skip leading whitespaces from next line
      while (pos < max) {
        ch = state.src.charCodeAt(pos);
        if (!isSpace(ch)) { break; }
        pos++;
      }

      state.pos = pos;
      return true;
    }
  }

  if (!silent) { state.pending += '\\'; }
  state.pos++;
  return true;
};


/***/ }),
/* 57 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Process html tags




var HTML_TAG_RE = __webpack_require__(6).HTML_TAG_RE;


function isLetter(ch) {
  /*eslint no-bitwise:0*/
  var lc = ch | 0x20; // to lower case
  return (lc >= 0x61/* a */) && (lc <= 0x7a/* z */);
}


module.exports = function html_inline(state, silent) {
  var ch, match, max, token,
      pos = state.pos;

  if (!state.md.options.html) { return false; }

  // Check start
  max = state.posMax;
  if (state.src.charCodeAt(pos) !== 0x3C/* < */ ||
      pos + 2 >= max) {
    return false;
  }

  // Quick fail on second char
  ch = state.src.charCodeAt(pos + 1);
  if (ch !== 0x21/* ! */ &&
      ch !== 0x3F/* ? */ &&
      ch !== 0x2F/* / */ &&
      !isLetter(ch)) {
    return false;
  }

  match = state.src.slice(pos).match(HTML_TAG_RE);
  if (!match) { return false; }

  if (!silent) {
    token         = state.push('html_inline', '', 0);
    token.content = state.src.slice(pos, pos + match[0].length);
  }
  state.pos += match[0].length;
  return true;
};


/***/ }),
/* 58 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Process ![image](<src> "title")



var normalizeReference   = __webpack_require__(0).normalizeReference;
var isSpace              = __webpack_require__(0).isSpace;


module.exports = function image(state, silent) {
  var attrs,
      code,
      content,
      label,
      labelEnd,
      labelStart,
      pos,
      ref,
      res,
      title,
      token,
      tokens,
      start,
      href = '',
      oldPos = state.pos,
      max = state.posMax;

  if (state.src.charCodeAt(state.pos) !== 0x21/* ! */) { return false; }
  if (state.src.charCodeAt(state.pos + 1) !== 0x5B/* [ */) { return false; }

  labelStart = state.pos + 2;
  labelEnd = state.md.helpers.parseLinkLabel(state, state.pos + 1, false);

  // parser failed to find ']', so it's not a valid link
  if (labelEnd < 0) { return false; }

  pos = labelEnd + 1;
  if (pos < max && state.src.charCodeAt(pos) === 0x28/* ( */) {
    //
    // Inline link
    //

    // [link](  <href>  "title"  )
    //        ^^ skipping these spaces
    pos++;
    for (; pos < max; pos++) {
      code = state.src.charCodeAt(pos);
      if (!isSpace(code) && code !== 0x0A) { break; }
    }
    if (pos >= max) { return false; }

    // [link](  <href>  "title"  )
    //          ^^^^^^ parsing link destination
    start = pos;
    res = state.md.helpers.parseLinkDestination(state.src, pos, state.posMax);
    if (res.ok) {
      href = state.md.normalizeLink(res.str);
      if (state.md.validateLink(href)) {
        pos = res.pos;
      } else {
        href = '';
      }
    }

    // [link](  <href>  "title"  )
    //                ^^ skipping these spaces
    start = pos;
    for (; pos < max; pos++) {
      code = state.src.charCodeAt(pos);
      if (!isSpace(code) && code !== 0x0A) { break; }
    }

    // [link](  <href>  "title"  )
    //                  ^^^^^^^ parsing link title
    res = state.md.helpers.parseLinkTitle(state.src, pos, state.posMax);
    if (pos < max && start !== pos && res.ok) {
      title = res.str;
      pos = res.pos;

      // [link](  <href>  "title"  )
      //                         ^^ skipping these spaces
      for (; pos < max; pos++) {
        code = state.src.charCodeAt(pos);
        if (!isSpace(code) && code !== 0x0A) { break; }
      }
    } else {
      title = '';
    }

    if (pos >= max || state.src.charCodeAt(pos) !== 0x29/* ) */) {
      state.pos = oldPos;
      return false;
    }
    pos++;
  } else {
    //
    // Link reference
    //
    if (typeof state.env.references === 'undefined') { return false; }

    if (pos < max && state.src.charCodeAt(pos) === 0x5B/* [ */) {
      start = pos + 1;
      pos = state.md.helpers.parseLinkLabel(state, pos);
      if (pos >= 0) {
        label = state.src.slice(start, pos++);
      } else {
        pos = labelEnd + 1;
      }
    } else {
      pos = labelEnd + 1;
    }

    // covers label === '' and label === undefined
    // (collapsed reference link and shortcut reference link respectively)
    if (!label) { label = state.src.slice(labelStart, labelEnd); }

    ref = state.env.references[normalizeReference(label)];
    if (!ref) {
      state.pos = oldPos;
      return false;
    }
    href = ref.href;
    title = ref.title;
  }

  //
  // We found the end of the link, and know for a fact it's a valid link;
  // so all that's left to do is to call tokenizer.
  //
  if (!silent) {
    content = state.src.slice(labelStart, labelEnd);

    state.md.inline.parse(
      content,
      state.md,
      state.env,
      tokens = []
    );

    token          = state.push('image', 'img', 0);
    token.attrs    = attrs = [ [ 'src', href ], [ 'alt', '' ] ];
    token.children = tokens;
    token.content  = content;

    if (title) {
      attrs.push([ 'title', title ]);
    }
  }

  state.pos = pos;
  state.posMax = max;
  return true;
};


/***/ }),
/* 59 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Process [link](<to> "stuff")



var normalizeReference   = __webpack_require__(0).normalizeReference;
var isSpace              = __webpack_require__(0).isSpace;


module.exports = function link(state, silent) {
  var attrs,
      code,
      label,
      labelEnd,
      labelStart,
      pos,
      res,
      ref,
      title,
      token,
      href = '',
      oldPos = state.pos,
      max = state.posMax,
      start = state.pos,
      parseReference = true;

  if (state.src.charCodeAt(state.pos) !== 0x5B/* [ */) { return false; }

  labelStart = state.pos + 1;
  labelEnd = state.md.helpers.parseLinkLabel(state, state.pos, true);

  // parser failed to find ']', so it's not a valid link
  if (labelEnd < 0) { return false; }

  pos = labelEnd + 1;
  if (pos < max && state.src.charCodeAt(pos) === 0x28/* ( */) {
    //
    // Inline link
    //

    // might have found a valid shortcut link, disable reference parsing
    parseReference = false;

    // [link](  <href>  "title"  )
    //        ^^ skipping these spaces
    pos++;
    for (; pos < max; pos++) {
      code = state.src.charCodeAt(pos);
      if (!isSpace(code) && code !== 0x0A) { break; }
    }
    if (pos >= max) { return false; }

    // [link](  <href>  "title"  )
    //          ^^^^^^ parsing link destination
    start = pos;
    res = state.md.helpers.parseLinkDestination(state.src, pos, state.posMax);
    if (res.ok) {
      href = state.md.normalizeLink(res.str);
      if (state.md.validateLink(href)) {
        pos = res.pos;
      } else {
        href = '';
      }
    }

    // [link](  <href>  "title"  )
    //                ^^ skipping these spaces
    start = pos;
    for (; pos < max; pos++) {
      code = state.src.charCodeAt(pos);
      if (!isSpace(code) && code !== 0x0A) { break; }
    }

    // [link](  <href>  "title"  )
    //                  ^^^^^^^ parsing link title
    res = state.md.helpers.parseLinkTitle(state.src, pos, state.posMax);
    if (pos < max && start !== pos && res.ok) {
      title = res.str;
      pos = res.pos;

      // [link](  <href>  "title"  )
      //                         ^^ skipping these spaces
      for (; pos < max; pos++) {
        code = state.src.charCodeAt(pos);
        if (!isSpace(code) && code !== 0x0A) { break; }
      }
    } else {
      title = '';
    }

    if (pos >= max || state.src.charCodeAt(pos) !== 0x29/* ) */) {
      // parsing a valid shortcut link failed, fallback to reference
      parseReference = true;
    }
    pos++;
  }

  if (parseReference) {
    //
    // Link reference
    //
    if (typeof state.env.references === 'undefined') { return false; }

    if (pos < max && state.src.charCodeAt(pos) === 0x5B/* [ */) {
      start = pos + 1;
      pos = state.md.helpers.parseLinkLabel(state, pos);
      if (pos >= 0) {
        label = state.src.slice(start, pos++);
      } else {
        pos = labelEnd + 1;
      }
    } else {
      pos = labelEnd + 1;
    }

    // covers label === '' and label === undefined
    // (collapsed reference link and shortcut reference link respectively)
    if (!label) { label = state.src.slice(labelStart, labelEnd); }

    ref = state.env.references[normalizeReference(label)];
    if (!ref) {
      state.pos = oldPos;
      return false;
    }
    href = ref.href;
    title = ref.title;
  }

  //
  // We found the end of the link, and know for a fact it's a valid link;
  // so all that's left to do is to call tokenizer.
  //
  if (!silent) {
    state.pos = labelStart;
    state.posMax = labelEnd;

    token        = state.push('link_open', 'a', 1);
    token.attrs  = attrs = [ [ 'href', href ] ];
    if (title) {
      attrs.push([ 'title', title ]);
    }

    state.md.inline.tokenize(state);

    token        = state.push('link_close', 'a', -1);
  }

  state.pos = pos;
  state.posMax = max;
  return true;
};


/***/ }),
/* 60 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Proceess '\n'



var isSpace = __webpack_require__(0).isSpace;


module.exports = function newline(state, silent) {
  var pmax, max, pos = state.pos;

  if (state.src.charCodeAt(pos) !== 0x0A/* \n */) { return false; }

  pmax = state.pending.length - 1;
  max = state.posMax;

  // '  \n' -> hardbreak
  // Lookup in pending chars is bad practice! Don't copy to other rules!
  // Pending string is stored in concat mode, indexed lookups will cause
  // convertion to flat mode.
  if (!silent) {
    if (pmax >= 0 && state.pending.charCodeAt(pmax) === 0x20) {
      if (pmax >= 1 && state.pending.charCodeAt(pmax - 1) === 0x20) {
        state.pending = state.pending.replace(/ +$/, '');
        state.push('hardbreak', 'br', 0);
      } else {
        state.pending = state.pending.slice(0, -1);
        state.push('softbreak', 'br', 0);
      }

    } else {
      state.push('softbreak', 'br', 0);
    }
  }

  pos++;

  // skip heading spaces for next line
  while (pos < max && isSpace(state.src.charCodeAt(pos))) { pos++; }

  state.pos = pos;
  return true;
};


/***/ }),
/* 61 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Inline parser state




var Token          = __webpack_require__(2);
var isWhiteSpace   = __webpack_require__(0).isWhiteSpace;
var isPunctChar    = __webpack_require__(0).isPunctChar;
var isMdAsciiPunct = __webpack_require__(0).isMdAsciiPunct;


function StateInline(src, md, env, outTokens) {
  this.src = src;
  this.env = env;
  this.md = md;
  this.tokens = outTokens;

  this.pos = 0;
  this.posMax = this.src.length;
  this.level = 0;
  this.pending = '';
  this.pendingLevel = 0;

  this.cache = {};        // Stores { start: end } pairs. Useful for backtrack
                          // optimization of pairs parse (emphasis, strikes).

  this.delimiters = [];   // Emphasis-like delimiters
}


// Flush pending text
//
StateInline.prototype.pushPending = function () {
  var token = new Token('text', '', 0);
  token.content = this.pending;
  token.level = this.pendingLevel;
  this.tokens.push(token);
  this.pending = '';
  return token;
};


// Push new token to "stream".
// If pending text exists - flush it as text token
//
StateInline.prototype.push = function (type, tag, nesting) {
  if (this.pending) {
    this.pushPending();
  }

  var token = new Token(type, tag, nesting);

  if (nesting < 0) { this.level--; }
  token.level = this.level;
  if (nesting > 0) { this.level++; }

  this.pendingLevel = this.level;
  this.tokens.push(token);
  return token;
};


// Scan a sequence of emphasis-like markers, and determine whether
// it can start an emphasis sequence or end an emphasis sequence.
//
//  - start - position to scan from (it should point at a valid marker);
//  - canSplitWord - determine if these markers can be found inside a word
//
StateInline.prototype.scanDelims = function (start, canSplitWord) {
  var pos = start, lastChar, nextChar, count, can_open, can_close,
      isLastWhiteSpace, isLastPunctChar,
      isNextWhiteSpace, isNextPunctChar,
      left_flanking = true,
      right_flanking = true,
      max = this.posMax,
      marker = this.src.charCodeAt(start);

  // treat beginning of the line as a whitespace
  lastChar = start > 0 ? this.src.charCodeAt(start - 1) : 0x20;

  while (pos < max && this.src.charCodeAt(pos) === marker) { pos++; }

  count = pos - start;

  // treat end of the line as a whitespace
  nextChar = pos < max ? this.src.charCodeAt(pos) : 0x20;

  isLastPunctChar = isMdAsciiPunct(lastChar) || isPunctChar(String.fromCharCode(lastChar));
  isNextPunctChar = isMdAsciiPunct(nextChar) || isPunctChar(String.fromCharCode(nextChar));

  isLastWhiteSpace = isWhiteSpace(lastChar);
  isNextWhiteSpace = isWhiteSpace(nextChar);

  if (isNextWhiteSpace) {
    left_flanking = false;
  } else if (isNextPunctChar) {
    if (!(isLastWhiteSpace || isLastPunctChar)) {
      left_flanking = false;
    }
  }

  if (isLastWhiteSpace) {
    right_flanking = false;
  } else if (isLastPunctChar) {
    if (!(isNextWhiteSpace || isNextPunctChar)) {
      right_flanking = false;
    }
  }

  if (!canSplitWord) {
    can_open  = left_flanking  && (!right_flanking || isLastPunctChar);
    can_close = right_flanking && (!left_flanking  || isNextPunctChar);
  } else {
    can_open  = left_flanking;
    can_close = right_flanking;
  }

  return {
    can_open:  can_open,
    can_close: can_close,
    length:    count
  };
};


// re-export Token class to use in block rules
StateInline.prototype.Token = Token;


module.exports = StateInline;


/***/ }),
/* 62 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Skip text characters for text token, place those to pending buffer
// and increment current pos




// Rule to skip pure text
// '{}$%@~+=:' reserved for extentions

// !, ", #, $, %, &, ', (, ), *, +, ,, -, ., /, :, ;, <, =, >, ?, @, [, \, ], ^, _, `, {, |, }, or ~

// !!!! Don't confuse with "Markdown ASCII Punctuation" chars
// http://spec.commonmark.org/0.15/#ascii-punctuation-character
function isTerminatorChar(ch) {
  switch (ch) {
    case 0x0A/* \n */:
    case 0x21/* ! */:
    case 0x23/* # */:
    case 0x24/* $ */:
    case 0x25/* % */:
    case 0x26/* & */:
    case 0x2A/* * */:
    case 0x2B/* + */:
    case 0x2D/* - */:
    case 0x3A/* : */:
    case 0x3C/* < */:
    case 0x3D/* = */:
    case 0x3E/* > */:
    case 0x40/* @ */:
    case 0x5B/* [ */:
    case 0x5C/* \ */:
    case 0x5D/* ] */:
    case 0x5E/* ^ */:
    case 0x5F/* _ */:
    case 0x60/* ` */:
    case 0x7B/* { */:
    case 0x7D/* } */:
    case 0x7E/* ~ */:
      return true;
    default:
      return false;
  }
}

module.exports = function text(state, silent) {
  var pos = state.pos;

  while (pos < state.posMax && !isTerminatorChar(state.src.charCodeAt(pos))) {
    pos++;
  }

  if (pos === state.pos) { return false; }

  if (!silent) { state.pending += state.src.slice(state.pos, pos); }

  state.pos = pos;

  return true;
};

// Alternative implementation, for memory.
//
// It costs 10% of performance, but allows extend terminators list, if place it
// to `ParcerInline` property. Probably, will switch to it sometime, such
// flexibility required.

/*
var TERMINATOR_RE = /[\n!#$%&*+\-:<=>@[\\\]^_`{}~]/;

module.exports = function text(state, silent) {
  var pos = state.pos,
      idx = state.src.slice(pos).search(TERMINATOR_RE);

  // first char is terminator -> empty text
  if (idx === 0) { return false; }

  // no terminator -> text till end of string
  if (idx < 0) {
    if (!silent) { state.pending += state.src.slice(pos); }
    state.pos = state.src.length;
    return true;
  }

  if (!silent) { state.pending += state.src.slice(pos, pos + idx); }

  state.pos += idx;

  return true;
};*/


/***/ }),
/* 63 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Merge adjacent text nodes into one, and re-calculate all token levels
//



module.exports = function text_collapse(state) {
  var curr, last,
      level = 0,
      tokens = state.tokens,
      max = state.tokens.length;

  for (curr = last = 0; curr < max; curr++) {
    // re-calculate levels
    level += tokens[curr].nesting;
    tokens[curr].level = level;

    if (tokens[curr].type === 'text' &&
        curr + 1 < max &&
        tokens[curr + 1].type === 'text') {

      // collapse two adjacent text nodes
      tokens[curr + 1].content = tokens[curr].content + tokens[curr + 1].content;
    } else {
      if (curr !== last) { tokens[last] = tokens[curr]; }

      last++;
    }
  }

  if (curr !== last) {
    tokens.length = last;
  }
};


/***/ }),
/* 64 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";




/* eslint-disable no-bitwise */

var decodeCache = {};

function getDecodeCache(exclude) {
  var i, ch, cache = decodeCache[exclude];
  if (cache) { return cache; }

  cache = decodeCache[exclude] = [];

  for (i = 0; i < 128; i++) {
    ch = String.fromCharCode(i);
    cache.push(ch);
  }

  for (i = 0; i < exclude.length; i++) {
    ch = exclude.charCodeAt(i);
    cache[ch] = '%' + ('0' + ch.toString(16).toUpperCase()).slice(-2);
  }

  return cache;
}


// Decode percent-encoded string.
//
function decode(string, exclude) {
  var cache;

  if (typeof exclude !== 'string') {
    exclude = decode.defaultChars;
  }

  cache = getDecodeCache(exclude);

  return string.replace(/(%[a-f0-9]{2})+/gi, function(seq) {
    var i, l, b1, b2, b3, b4, chr,
        result = '';

    for (i = 0, l = seq.length; i < l; i += 3) {
      b1 = parseInt(seq.slice(i + 1, i + 3), 16);

      if (b1 < 0x80) {
        result += cache[b1];
        continue;
      }

      if ((b1 & 0xE0) === 0xC0 && (i + 3 < l)) {
        // 110xxxxx 10xxxxxx
        b2 = parseInt(seq.slice(i + 4, i + 6), 16);

        if ((b2 & 0xC0) === 0x80) {
          chr = ((b1 << 6) & 0x7C0) | (b2 & 0x3F);

          if (chr < 0x80) {
            result += '\ufffd\ufffd';
          } else {
            result += String.fromCharCode(chr);
          }

          i += 3;
          continue;
        }
      }

      if ((b1 & 0xF0) === 0xE0 && (i + 6 < l)) {
        // 1110xxxx 10xxxxxx 10xxxxxx
        b2 = parseInt(seq.slice(i + 4, i + 6), 16);
        b3 = parseInt(seq.slice(i + 7, i + 9), 16);

        if ((b2 & 0xC0) === 0x80 && (b3 & 0xC0) === 0x80) {
          chr = ((b1 << 12) & 0xF000) | ((b2 << 6) & 0xFC0) | (b3 & 0x3F);

          if (chr < 0x800 || (chr >= 0xD800 && chr <= 0xDFFF)) {
            result += '\ufffd\ufffd\ufffd';
          } else {
            result += String.fromCharCode(chr);
          }

          i += 6;
          continue;
        }
      }

      if ((b1 & 0xF8) === 0xF0 && (i + 9 < l)) {
        // 111110xx 10xxxxxx 10xxxxxx 10xxxxxx
        b2 = parseInt(seq.slice(i + 4, i + 6), 16);
        b3 = parseInt(seq.slice(i + 7, i + 9), 16);
        b4 = parseInt(seq.slice(i + 10, i + 12), 16);

        if ((b2 & 0xC0) === 0x80 && (b3 & 0xC0) === 0x80 && (b4 & 0xC0) === 0x80) {
          chr = ((b1 << 18) & 0x1C0000) | ((b2 << 12) & 0x3F000) | ((b3 << 6) & 0xFC0) | (b4 & 0x3F);

          if (chr < 0x10000 || chr > 0x10FFFF) {
            result += '\ufffd\ufffd\ufffd\ufffd';
          } else {
            chr -= 0x10000;
            result += String.fromCharCode(0xD800 + (chr >> 10), 0xDC00 + (chr & 0x3FF));
          }

          i += 9;
          continue;
        }
      }

      result += '\ufffd';
    }

    return result;
  });
}


decode.defaultChars   = ';/?:@&=+$,#';
decode.componentChars = '';


module.exports = decode;


/***/ }),
/* 65 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";




var encodeCache = {};


// Create a lookup array where anything but characters in `chars` string
// and alphanumeric chars is percent-encoded.
//
function getEncodeCache(exclude) {
  var i, ch, cache = encodeCache[exclude];
  if (cache) { return cache; }

  cache = encodeCache[exclude] = [];

  for (i = 0; i < 128; i++) {
    ch = String.fromCharCode(i);

    if (/^[0-9a-z]$/i.test(ch)) {
      // always allow unencoded alphanumeric characters
      cache.push(ch);
    } else {
      cache.push('%' + ('0' + i.toString(16).toUpperCase()).slice(-2));
    }
  }

  for (i = 0; i < exclude.length; i++) {
    cache[exclude.charCodeAt(i)] = exclude[i];
  }

  return cache;
}


// Encode unsafe characters with percent-encoding, skipping already
// encoded sequences.
//
//  - string       - string to encode
//  - exclude      - list of characters to ignore (in addition to a-zA-Z0-9)
//  - keepEscaped  - don't encode '%' in a correct escape sequence (default: true)
//
function encode(string, exclude, keepEscaped) {
  var i, l, code, nextCode, cache,
      result = '';

  if (typeof exclude !== 'string') {
    // encode(string, keepEscaped)
    keepEscaped  = exclude;
    exclude = encode.defaultChars;
  }

  if (typeof keepEscaped === 'undefined') {
    keepEscaped = true;
  }

  cache = getEncodeCache(exclude);

  for (i = 0, l = string.length; i < l; i++) {
    code = string.charCodeAt(i);

    if (keepEscaped && code === 0x25 /* % */ && i + 2 < l) {
      if (/^[0-9a-f]{2}$/i.test(string.slice(i + 1, i + 3))) {
        result += string.slice(i, i + 3);
        i += 2;
        continue;
      }
    }

    if (code < 128) {
      result += cache[code];
      continue;
    }

    if (code >= 0xD800 && code <= 0xDFFF) {
      if (code >= 0xD800 && code <= 0xDBFF && i + 1 < l) {
        nextCode = string.charCodeAt(i + 1);
        if (nextCode >= 0xDC00 && nextCode <= 0xDFFF) {
          result += encodeURIComponent(string[i] + string[i + 1]);
          i++;
          continue;
        }
      }
      result += '%EF%BF%BD';
      continue;
    }

    result += encodeURIComponent(string[i]);
  }

  return result;
}

encode.defaultChars   = ";/?:@&=+$,-_.!~*'()#";
encode.componentChars = "-_.!~*'()";


module.exports = encode;


/***/ }),
/* 66 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";




module.exports = function format(url) {
  var result = '';

  result += url.protocol || '';
  result += url.slashes ? '//' : '';
  result += url.auth ? url.auth + '@' : '';

  if (url.hostname && url.hostname.indexOf(':') !== -1) {
    // ipv6 address
    result += '[' + url.hostname + ']';
  } else {
    result += url.hostname || '';
  }

  result += url.port ? ':' + url.port : '';
  result += url.pathname || '';
  result += url.search || '';
  result += url.hash || '';

  return result;
};


/***/ }),
/* 67 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
// Copyright Joyent, Inc. and other Node contributors.
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to permit
// persons to whom the Software is furnished to do so, subject to the
// following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
// OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
// NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
// DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
// USE OR OTHER DEALINGS IN THE SOFTWARE.



//
// Changes from joyent/node:
//
// 1. No leading slash in paths,
//    e.g. in `url.parse('http://foo?bar')` pathname is ``, not `/`
//
// 2. Backslashes are not replaced with slashes,
//    so `http:\\example.org\` is treated like a relative path
//
// 3. Trailing colon is treated like a part of the path,
//    i.e. in `http://example.org:foo` pathname is `:foo`
//
// 4. Nothing is URL-encoded in the resulting object,
//    (in joyent/node some chars in auth and paths are encoded)
//
// 5. `url.parse()` does not have `parseQueryString` argument
//
// 6. Removed extraneous result properties: `host`, `path`, `query`, etc.,
//    which can be constructed using other parts of the url.
//


function Url() {
  this.protocol = null;
  this.slashes = null;
  this.auth = null;
  this.port = null;
  this.hostname = null;
  this.hash = null;
  this.search = null;
  this.pathname = null;
}

// Reference: RFC 3986, RFC 1808, RFC 2396

// define these here so at least they only have to be
// compiled once on the first module load.
var protocolPattern = /^([a-z0-9.+-]+:)/i,
    portPattern = /:[0-9]*$/,

    // Special case for a simple path URL
    simplePathPattern = /^(\/\/?(?!\/)[^\?\s]*)(\?[^\s]*)?$/,

    // RFC 2396: characters reserved for delimiting URLs.
    // We actually just auto-escape these.
    delims = [ '<', '>', '"', '`', ' ', '\r', '\n', '\t' ],

    // RFC 2396: characters not allowed for various reasons.
    unwise = [ '{', '}', '|', '\\', '^', '`' ].concat(delims),

    // Allowed by RFCs, but cause of XSS attacks.  Always escape these.
    autoEscape = [ '\'' ].concat(unwise),
    // Characters that are never ever allowed in a hostname.
    // Note that any invalid chars are also handled, but these
    // are the ones that are *expected* to be seen, so we fast-path
    // them.
    nonHostChars = [ '%', '/', '?', ';', '#' ].concat(autoEscape),
    hostEndingChars = [ '/', '?', '#' ],
    hostnameMaxLen = 255,
    hostnamePartPattern = /^[+a-z0-9A-Z_-]{0,63}$/,
    hostnamePartStart = /^([+a-z0-9A-Z_-]{0,63})(.*)$/,
    // protocols that can allow "unsafe" and "unwise" chars.
    /* eslint-disable no-script-url */
    // protocols that never have a hostname.
    hostlessProtocol = {
      'javascript': true,
      'javascript:': true
    },
    // protocols that always contain a // bit.
    slashedProtocol = {
      'http': true,
      'https': true,
      'ftp': true,
      'gopher': true,
      'file': true,
      'http:': true,
      'https:': true,
      'ftp:': true,
      'gopher:': true,
      'file:': true
    };
    /* eslint-enable no-script-url */

function urlParse(url, slashesDenoteHost) {
  if (url && url instanceof Url) { return url; }

  var u = new Url();
  u.parse(url, slashesDenoteHost);
  return u;
}

Url.prototype.parse = function(url, slashesDenoteHost) {
  var i, l, lowerProto, hec, slashes,
      rest = url;

  // trim before proceeding.
  // This is to support parse stuff like "  http://foo.com  \n"
  rest = rest.trim();

  if (!slashesDenoteHost && url.split('#').length === 1) {
    // Try fast path regexp
    var simplePath = simplePathPattern.exec(rest);
    if (simplePath) {
      this.pathname = simplePath[1];
      if (simplePath[2]) {
        this.search = simplePath[2];
      }
      return this;
    }
  }

  var proto = protocolPattern.exec(rest);
  if (proto) {
    proto = proto[0];
    lowerProto = proto.toLowerCase();
    this.protocol = proto;
    rest = rest.substr(proto.length);
  }

  // figure out if it's got a host
  // user@server is *always* interpreted as a hostname, and url
  // resolution will treat //foo/bar as host=foo,path=bar because that's
  // how the browser resolves relative URLs.
  if (slashesDenoteHost || proto || rest.match(/^\/\/[^@\/]+@[^@\/]+/)) {
    slashes = rest.substr(0, 2) === '//';
    if (slashes && !(proto && hostlessProtocol[proto])) {
      rest = rest.substr(2);
      this.slashes = true;
    }
  }

  if (!hostlessProtocol[proto] &&
      (slashes || (proto && !slashedProtocol[proto]))) {

    // there's a hostname.
    // the first instance of /, ?, ;, or # ends the host.
    //
    // If there is an @ in the hostname, then non-host chars *are* allowed
    // to the left of the last @ sign, unless some host-ending character
    // comes *before* the @-sign.
    // URLs are obnoxious.
    //
    // ex:
    // http://a@b@c/ => user:a@b host:c
    // http://a@b?@c => user:a host:c path:/?@c

    // v0.12 TODO(isaacs): This is not quite how Chrome does things.
    // Review our test case against browsers more comprehensively.

    // find the first instance of any hostEndingChars
    var hostEnd = -1;
    for (i = 0; i < hostEndingChars.length; i++) {
      hec = rest.indexOf(hostEndingChars[i]);
      if (hec !== -1 && (hostEnd === -1 || hec < hostEnd)) {
        hostEnd = hec;
      }
    }

    // at this point, either we have an explicit point where the
    // auth portion cannot go past, or the last @ char is the decider.
    var auth, atSign;
    if (hostEnd === -1) {
      // atSign can be anywhere.
      atSign = rest.lastIndexOf('@');
    } else {
      // atSign must be in auth portion.
      // http://a@b/c@d => host:b auth:a path:/c@d
      atSign = rest.lastIndexOf('@', hostEnd);
    }

    // Now we have a portion which is definitely the auth.
    // Pull that off.
    if (atSign !== -1) {
      auth = rest.slice(0, atSign);
      rest = rest.slice(atSign + 1);
      this.auth = auth;
    }

    // the host is the remaining to the left of the first non-host char
    hostEnd = -1;
    for (i = 0; i < nonHostChars.length; i++) {
      hec = rest.indexOf(nonHostChars[i]);
      if (hec !== -1 && (hostEnd === -1 || hec < hostEnd)) {
        hostEnd = hec;
      }
    }
    // if we still have not hit it, then the entire thing is a host.
    if (hostEnd === -1) {
      hostEnd = rest.length;
    }

    if (rest[hostEnd - 1] === ':') { hostEnd--; }
    var host = rest.slice(0, hostEnd);
    rest = rest.slice(hostEnd);

    // pull out port.
    this.parseHost(host);

    // we've indicated that there is a hostname,
    // so even if it's empty, it has to be present.
    this.hostname = this.hostname || '';

    // if hostname begins with [ and ends with ]
    // assume that it's an IPv6 address.
    var ipv6Hostname = this.hostname[0] === '[' &&
        this.hostname[this.hostname.length - 1] === ']';

    // validate a little.
    if (!ipv6Hostname) {
      var hostparts = this.hostname.split(/\./);
      for (i = 0, l = hostparts.length; i < l; i++) {
        var part = hostparts[i];
        if (!part) { continue; }
        if (!part.match(hostnamePartPattern)) {
          var newpart = '';
          for (var j = 0, k = part.length; j < k; j++) {
            if (part.charCodeAt(j) > 127) {
              // we replace non-ASCII char with a temporary placeholder
              // we need this to make sure size of hostname is not
              // broken by replacing non-ASCII by nothing
              newpart += 'x';
            } else {
              newpart += part[j];
            }
          }
          // we test again with ASCII char only
          if (!newpart.match(hostnamePartPattern)) {
            var validParts = hostparts.slice(0, i);
            var notHost = hostparts.slice(i + 1);
            var bit = part.match(hostnamePartStart);
            if (bit) {
              validParts.push(bit[1]);
              notHost.unshift(bit[2]);
            }
            if (notHost.length) {
              rest = notHost.join('.') + rest;
            }
            this.hostname = validParts.join('.');
            break;
          }
        }
      }
    }

    if (this.hostname.length > hostnameMaxLen) {
      this.hostname = '';
    }

    // strip [ and ] from the hostname
    // the host field still retains them, though
    if (ipv6Hostname) {
      this.hostname = this.hostname.substr(1, this.hostname.length - 2);
    }
  }

  // chop off from the tail first.
  var hash = rest.indexOf('#');
  if (hash !== -1) {
    // got a fragment string.
    this.hash = rest.substr(hash);
    rest = rest.slice(0, hash);
  }
  var qm = rest.indexOf('?');
  if (qm !== -1) {
    this.search = rest.substr(qm);
    rest = rest.slice(0, qm);
  }
  if (rest) { this.pathname = rest; }
  if (slashedProtocol[lowerProto] &&
      this.hostname && !this.pathname) {
    this.pathname = '';
  }

  return this;
};

Url.prototype.parseHost = function(host) {
  var port = portPattern.exec(host);
  if (port) {
    port = port[0];
    if (port !== ':') {
      this.port = port.substr(1);
    }
    host = host.substr(0, host.length - port.length);
  }
  if (host) { this.hostname = host; }
};

module.exports = urlParse;


/***/ }),
/* 68 */
/***/ (function(module, exports, __webpack_require__) {

/* WEBPACK VAR INJECTION */(function(module, global) {var __WEBPACK_AMD_DEFINE_RESULT__;/*! https://mths.be/punycode v1.4.1 by @mathias */
;(function(root) {

	/** Detect free variables */
	var freeExports = typeof exports == 'object' && exports &&
		!exports.nodeType && exports;
	var freeModule = typeof module == 'object' && module &&
		!module.nodeType && module;
	var freeGlobal = typeof global == 'object' && global;
	if (
		freeGlobal.global === freeGlobal ||
		freeGlobal.window === freeGlobal ||
		freeGlobal.self === freeGlobal
	) {
		root = freeGlobal;
	}

	/**
	 * The `punycode` object.
	 * @name punycode
	 * @type Object
	 */
	var punycode,

	/** Highest positive signed 32-bit float value */
	maxInt = 2147483647, // aka. 0x7FFFFFFF or 2^31-1

	/** Bootstring parameters */
	base = 36,
	tMin = 1,
	tMax = 26,
	skew = 38,
	damp = 700,
	initialBias = 72,
	initialN = 128, // 0x80
	delimiter = '-', // '\x2D'

	/** Regular expressions */
	regexPunycode = /^xn--/,
	regexNonASCII = /[^\x20-\x7E]/, // unprintable ASCII chars + non-ASCII chars
	regexSeparators = /[\x2E\u3002\uFF0E\uFF61]/g, // RFC 3490 separators

	/** Error messages */
	errors = {
		'overflow': 'Overflow: input needs wider integers to process',
		'not-basic': 'Illegal input >= 0x80 (not a basic code point)',
		'invalid-input': 'Invalid input'
	},

	/** Convenience shortcuts */
	baseMinusTMin = base - tMin,
	floor = Math.floor,
	stringFromCharCode = String.fromCharCode,

	/** Temporary variable */
	key;

	/*--------------------------------------------------------------------------*/

	/**
	 * A generic error utility function.
	 * @private
	 * @param {String} type The error type.
	 * @returns {Error} Throws a `RangeError` with the applicable error message.
	 */
	function error(type) {
		throw new RangeError(errors[type]);
	}

	/**
	 * A generic `Array#map` utility function.
	 * @private
	 * @param {Array} array The array to iterate over.
	 * @param {Function} callback The function that gets called for every array
	 * item.
	 * @returns {Array} A new array of values returned by the callback function.
	 */
	function map(array, fn) {
		var length = array.length;
		var result = [];
		while (length--) {
			result[length] = fn(array[length]);
		}
		return result;
	}

	/**
	 * A simple `Array#map`-like wrapper to work with domain name strings or email
	 * addresses.
	 * @private
	 * @param {String} domain The domain name or email address.
	 * @param {Function} callback The function that gets called for every
	 * character.
	 * @returns {Array} A new string of characters returned by the callback
	 * function.
	 */
	function mapDomain(string, fn) {
		var parts = string.split('@');
		var result = '';
		if (parts.length > 1) {
			// In email addresses, only the domain name should be punycoded. Leave
			// the local part (i.e. everything up to `@`) intact.
			result = parts[0] + '@';
			string = parts[1];
		}
		// Avoid `split(regex)` for IE8 compatibility. See #17.
		string = string.replace(regexSeparators, '\x2E');
		var labels = string.split('.');
		var encoded = map(labels, fn).join('.');
		return result + encoded;
	}

	/**
	 * Creates an array containing the numeric code points of each Unicode
	 * character in the string. While JavaScript uses UCS-2 internally,
	 * this function will convert a pair of surrogate halves (each of which
	 * UCS-2 exposes as separate characters) into a single code point,
	 * matching UTF-16.
	 * @see `punycode.ucs2.encode`
	 * @see <https://mathiasbynens.be/notes/javascript-encoding>
	 * @memberOf punycode.ucs2
	 * @name decode
	 * @param {String} string The Unicode input string (UCS-2).
	 * @returns {Array} The new array of code points.
	 */
	function ucs2decode(string) {
		var output = [],
		    counter = 0,
		    length = string.length,
		    value,
		    extra;
		while (counter < length) {
			value = string.charCodeAt(counter++);
			if (value >= 0xD800 && value <= 0xDBFF && counter < length) {
				// high surrogate, and there is a next character
				extra = string.charCodeAt(counter++);
				if ((extra & 0xFC00) == 0xDC00) { // low surrogate
					output.push(((value & 0x3FF) << 10) + (extra & 0x3FF) + 0x10000);
				} else {
					// unmatched surrogate; only append this code unit, in case the next
					// code unit is the high surrogate of a surrogate pair
					output.push(value);
					counter--;
				}
			} else {
				output.push(value);
			}
		}
		return output;
	}

	/**
	 * Creates a string based on an array of numeric code points.
	 * @see `punycode.ucs2.decode`
	 * @memberOf punycode.ucs2
	 * @name encode
	 * @param {Array} codePoints The array of numeric code points.
	 * @returns {String} The new Unicode string (UCS-2).
	 */
	function ucs2encode(array) {
		return map(array, function(value) {
			var output = '';
			if (value > 0xFFFF) {
				value -= 0x10000;
				output += stringFromCharCode(value >>> 10 & 0x3FF | 0xD800);
				value = 0xDC00 | value & 0x3FF;
			}
			output += stringFromCharCode(value);
			return output;
		}).join('');
	}

	/**
	 * Converts a basic code point into a digit/integer.
	 * @see `digitToBasic()`
	 * @private
	 * @param {Number} codePoint The basic numeric code point value.
	 * @returns {Number} The numeric value of a basic code point (for use in
	 * representing integers) in the range `0` to `base - 1`, or `base` if
	 * the code point does not represent a value.
	 */
	function basicToDigit(codePoint) {
		if (codePoint - 48 < 10) {
			return codePoint - 22;
		}
		if (codePoint - 65 < 26) {
			return codePoint - 65;
		}
		if (codePoint - 97 < 26) {
			return codePoint - 97;
		}
		return base;
	}

	/**
	 * Converts a digit/integer into a basic code point.
	 * @see `basicToDigit()`
	 * @private
	 * @param {Number} digit The numeric value of a basic code point.
	 * @returns {Number} The basic code point whose value (when used for
	 * representing integers) is `digit`, which needs to be in the range
	 * `0` to `base - 1`. If `flag` is non-zero, the uppercase form is
	 * used; else, the lowercase form is used. The behavior is undefined
	 * if `flag` is non-zero and `digit` has no uppercase form.
	 */
	function digitToBasic(digit, flag) {
		//  0..25 map to ASCII a..z or A..Z
		// 26..35 map to ASCII 0..9
		return digit + 22 + 75 * (digit < 26) - ((flag != 0) << 5);
	}

	/**
	 * Bias adaptation function as per section 3.4 of RFC 3492.
	 * https://tools.ietf.org/html/rfc3492#section-3.4
	 * @private
	 */
	function adapt(delta, numPoints, firstTime) {
		var k = 0;
		delta = firstTime ? floor(delta / damp) : delta >> 1;
		delta += floor(delta / numPoints);
		for (/* no initialization */; delta > baseMinusTMin * tMax >> 1; k += base) {
			delta = floor(delta / baseMinusTMin);
		}
		return floor(k + (baseMinusTMin + 1) * delta / (delta + skew));
	}

	/**
	 * Converts a Punycode string of ASCII-only symbols to a string of Unicode
	 * symbols.
	 * @memberOf punycode
	 * @param {String} input The Punycode string of ASCII-only symbols.
	 * @returns {String} The resulting string of Unicode symbols.
	 */
	function decode(input) {
		// Don't use UCS-2
		var output = [],
		    inputLength = input.length,
		    out,
		    i = 0,
		    n = initialN,
		    bias = initialBias,
		    basic,
		    j,
		    index,
		    oldi,
		    w,
		    k,
		    digit,
		    t,
		    /** Cached calculation results */
		    baseMinusT;

		// Handle the basic code points: let `basic` be the number of input code
		// points before the last delimiter, or `0` if there is none, then copy
		// the first basic code points to the output.

		basic = input.lastIndexOf(delimiter);
		if (basic < 0) {
			basic = 0;
		}

		for (j = 0; j < basic; ++j) {
			// if it's not a basic code point
			if (input.charCodeAt(j) >= 0x80) {
				error('not-basic');
			}
			output.push(input.charCodeAt(j));
		}

		// Main decoding loop: start just after the last delimiter if any basic code
		// points were copied; start at the beginning otherwise.

		for (index = basic > 0 ? basic + 1 : 0; index < inputLength; /* no final expression */) {

			// `index` is the index of the next character to be consumed.
			// Decode a generalized variable-length integer into `delta`,
			// which gets added to `i`. The overflow checking is easier
			// if we increase `i` as we go, then subtract off its starting
			// value at the end to obtain `delta`.
			for (oldi = i, w = 1, k = base; /* no condition */; k += base) {

				if (index >= inputLength) {
					error('invalid-input');
				}

				digit = basicToDigit(input.charCodeAt(index++));

				if (digit >= base || digit > floor((maxInt - i) / w)) {
					error('overflow');
				}

				i += digit * w;
				t = k <= bias ? tMin : (k >= bias + tMax ? tMax : k - bias);

				if (digit < t) {
					break;
				}

				baseMinusT = base - t;
				if (w > floor(maxInt / baseMinusT)) {
					error('overflow');
				}

				w *= baseMinusT;

			}

			out = output.length + 1;
			bias = adapt(i - oldi, out, oldi == 0);

			// `i` was supposed to wrap around from `out` to `0`,
			// incrementing `n` each time, so we'll fix that now:
			if (floor(i / out) > maxInt - n) {
				error('overflow');
			}

			n += floor(i / out);
			i %= out;

			// Insert `n` at position `i` of the output
			output.splice(i++, 0, n);

		}

		return ucs2encode(output);
	}

	/**
	 * Converts a string of Unicode symbols (e.g. a domain name label) to a
	 * Punycode string of ASCII-only symbols.
	 * @memberOf punycode
	 * @param {String} input The string of Unicode symbols.
	 * @returns {String} The resulting Punycode string of ASCII-only symbols.
	 */
	function encode(input) {
		var n,
		    delta,
		    handledCPCount,
		    basicLength,
		    bias,
		    j,
		    m,
		    q,
		    k,
		    t,
		    currentValue,
		    output = [],
		    /** `inputLength` will hold the number of code points in `input`. */
		    inputLength,
		    /** Cached calculation results */
		    handledCPCountPlusOne,
		    baseMinusT,
		    qMinusT;

		// Convert the input in UCS-2 to Unicode
		input = ucs2decode(input);

		// Cache the length
		inputLength = input.length;

		// Initialize the state
		n = initialN;
		delta = 0;
		bias = initialBias;

		// Handle the basic code points
		for (j = 0; j < inputLength; ++j) {
			currentValue = input[j];
			if (currentValue < 0x80) {
				output.push(stringFromCharCode(currentValue));
			}
		}

		handledCPCount = basicLength = output.length;

		// `handledCPCount` is the number of code points that have been handled;
		// `basicLength` is the number of basic code points.

		// Finish the basic string - if it is not empty - with a delimiter
		if (basicLength) {
			output.push(delimiter);
		}

		// Main encoding loop:
		while (handledCPCount < inputLength) {

			// All non-basic code points < n have been handled already. Find the next
			// larger one:
			for (m = maxInt, j = 0; j < inputLength; ++j) {
				currentValue = input[j];
				if (currentValue >= n && currentValue < m) {
					m = currentValue;
				}
			}

			// Increase `delta` enough to advance the decoder's <n,i> state to <m,0>,
			// but guard against overflow
			handledCPCountPlusOne = handledCPCount + 1;
			if (m - n > floor((maxInt - delta) / handledCPCountPlusOne)) {
				error('overflow');
			}

			delta += (m - n) * handledCPCountPlusOne;
			n = m;

			for (j = 0; j < inputLength; ++j) {
				currentValue = input[j];

				if (currentValue < n && ++delta > maxInt) {
					error('overflow');
				}

				if (currentValue == n) {
					// Represent delta as a generalized variable-length integer
					for (q = delta, k = base; /* no condition */; k += base) {
						t = k <= bias ? tMin : (k >= bias + tMax ? tMax : k - bias);
						if (q < t) {
							break;
						}
						qMinusT = q - t;
						baseMinusT = base - t;
						output.push(
							stringFromCharCode(digitToBasic(t + qMinusT % baseMinusT, 0))
						);
						q = floor(qMinusT / baseMinusT);
					}

					output.push(stringFromCharCode(digitToBasic(q, 0)));
					bias = adapt(delta, handledCPCountPlusOne, handledCPCount == basicLength);
					delta = 0;
					++handledCPCount;
				}
			}

			++delta;
			++n;

		}
		return output.join('');
	}

	/**
	 * Converts a Punycode string representing a domain name or an email address
	 * to Unicode. Only the Punycoded parts of the input will be converted, i.e.
	 * it doesn't matter if you call it on a string that has already been
	 * converted to Unicode.
	 * @memberOf punycode
	 * @param {String} input The Punycoded domain name or email address to
	 * convert to Unicode.
	 * @returns {String} The Unicode representation of the given Punycode
	 * string.
	 */
	function toUnicode(input) {
		return mapDomain(input, function(string) {
			return regexPunycode.test(string)
				? decode(string.slice(4).toLowerCase())
				: string;
		});
	}

	/**
	 * Converts a Unicode string representing a domain name or an email address to
	 * Punycode. Only the non-ASCII parts of the domain name will be converted,
	 * i.e. it doesn't matter if you call it with a domain that's already in
	 * ASCII.
	 * @memberOf punycode
	 * @param {String} input The domain name or email address to convert, as a
	 * Unicode string.
	 * @returns {String} The Punycode representation of the given domain name or
	 * email address.
	 */
	function toASCII(input) {
		return mapDomain(input, function(string) {
			return regexNonASCII.test(string)
				? 'xn--' + encode(string)
				: string;
		});
	}

	/*--------------------------------------------------------------------------*/

	/** Define the public API */
	punycode = {
		/**
		 * A string representing the current Punycode.js version number.
		 * @memberOf punycode
		 * @type String
		 */
		'version': '1.4.1',
		/**
		 * An object of methods to convert from JavaScript's internal character
		 * representation (UCS-2) to Unicode code points, and back.
		 * @see <https://mathiasbynens.be/notes/javascript-encoding>
		 * @memberOf punycode
		 * @type Object
		 */
		'ucs2': {
			'decode': ucs2decode,
			'encode': ucs2encode
		},
		'decode': decode,
		'encode': encode,
		'toASCII': toASCII,
		'toUnicode': toUnicode
	};

	/** Expose `punycode` */
	// Some AMD build optimizers, like r.js, check for specific condition patterns
	// like the following:
	if (
		true
	) {
		!(__WEBPACK_AMD_DEFINE_RESULT__ = function() {
			return punycode;
		}.call(exports, __webpack_require__, exports, module),
				__WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));
	} else if (freeExports && freeModule) {
		if (module.exports == freeExports) {
			// in Node.js, io.js, or RingoJS v0.8.0+
			freeModule.exports = punycode;
		} else {
			// in Narwhal or RingoJS v0.7.0-
			for (key in punycode) {
				punycode.hasOwnProperty(key) && (freeExports[key] = punycode[key]);
			}
		}
	} else {
		// in Rhino or a web browser
		root.punycode = punycode;
	}

}(this));

/* WEBPACK VAR INJECTION */}.call(exports, __webpack_require__(74)(module), __webpack_require__(73)))

/***/ }),
/* 69 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

function __export(m) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}
Object.defineProperty(exports, "__esModule", { value: true });
__export(__webpack_require__(14));
__export(__webpack_require__(4));
__export(__webpack_require__(15));


/***/ }),
/* 70 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var AbstractTextFormatter = (function () {
    function AbstractTextFormatter(regularExpression) {
        this._regularExpression = regularExpression;
    }
    AbstractTextFormatter.prototype.format = function (input) {
        var matches;
        var result = input;
        while ((matches = this._regularExpression.exec(input)) != null) {
            result = result.replace(matches[0], this.internalFormat(matches));
        }
        ;
        return result;
    };
    return AbstractTextFormatter;
}());
var DateFormatter = (function (_super) {
    __extends(DateFormatter, _super);
    function DateFormatter() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    DateFormatter.prototype.internalFormat = function (matches) {
        var date = new Date(Date.parse(matches[1]));
        if (matches[2] != undefined) {
            var longOrShort = matches[2].toLowerCase();
            return date.toLocaleDateString(undefined, { day: "numeric", weekday: longOrShort, month: longOrShort, year: "numeric" });
        }
        else {
            return date.toLocaleDateString();
        }
    };
    return DateFormatter;
}(AbstractTextFormatter));
var TimeFormatter = (function (_super) {
    __extends(TimeFormatter, _super);
    function TimeFormatter() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TimeFormatter.prototype.internalFormat = function (matches) {
        var date = new Date(Date.parse(matches[1]));
        return date.toLocaleTimeString();
    };
    return TimeFormatter;
}(AbstractTextFormatter));
function formatText(text) {
    var formatters = [
        new DateFormatter(/\{{2}date\((\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:z|Z|-\d{1,2}))(?:,(long|short))?\)\}{2}/gi),
        new TimeFormatter(/\{{2}time\((\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:z|Z|-\d{1,2}))\)\}{2}/gi)
    ];
    var result = text;
    for (var i = 0; i < formatters.length; i++) {
        result = formatters[i].format(result);
    }
    return result;
}
exports.formatText = formatText;


/***/ }),
/* 71 */
/***/ (function(module, exports) {

module.exports=/[\xAD\u0600-\u0605\u061C\u06DD\u070F\u08E2\u180E\u200B-\u200F\u202A-\u202E\u2060-\u2064\u2066-\u206F\uFEFF\uFFF9-\uFFFB]|\uD804\uDCBD|\uD82F[\uDCA0-\uDCA3]|\uD834[\uDD73-\uDD7A]|\uDB40[\uDC01\uDC20-\uDC7F]/

/***/ }),
/* 72 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


exports.Any = __webpack_require__(13);
exports.Cc  = __webpack_require__(11);
exports.Cf  = __webpack_require__(71);
exports.P   = __webpack_require__(3);
exports.Z   = __webpack_require__(12);


/***/ }),
/* 73 */
/***/ (function(module, exports) {

var g;

// This works in non-strict mode
g = (function() {
	return this;
})();

try {
	// This works if eval is allowed (see CSP)
	g = g || Function("return this")() || (1,eval)("this");
} catch(e) {
	// This works if the window reference is available
	if(typeof window === "object")
		g = window;
}

// g can still be undefined, but nothing to do about it...
// We return undefined, instead of nothing here, so it's
// easier to handle this case. if(!global) { ...}

module.exports = g;


/***/ }),
/* 74 */
/***/ (function(module, exports) {

module.exports = function(module) {
	if(!module.webpackPolyfill) {
		module.deprecate = function() {};
		module.paths = [];
		// module.parent = undefined by default
		if(!module.children) module.children = [];
		Object.defineProperty(module, "loaded", {
			enumerable: true,
			get: function() {
				return module.l;
			}
		});
		Object.defineProperty(module, "id", {
			enumerable: true,
			get: function() {
				return module.i;
			}
		});
		module.webpackPolyfill = 1;
	}
	return module;
};


/***/ })
/******/ ]);
