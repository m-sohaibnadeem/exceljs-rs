// The parts of dayjs 1.11 (MIT) that ExcelJS's CSV code uses, with the
// customParseFormat and utc plugins and the default 'en' locale:
//   dayjs(input, format, true)  -> parseStrict(input, format)
//   dayjs(date).format(fmt)     -> format(date, fmt, false)
//   dayjs.utc(date).format(fmt) -> format(date, fmt, true)
// Ported from dayjs/esm/index.js and plugin/{customParseFormat,utc}/index.js.
'use strict';

const INVALID_DATE_STRING = 'Invalid Date';
// Returned by a token parser where dayjs throws.
const INVALID = Symbol('invalid');
const FORMAT_DEFAULT = 'YYYY-MM-DDTHH:mm:ssZ';
const UTC_FORMAT_DEFAULT = 'YYYY-MM-DDTHH:mm:ss[Z]';
const REGEX_FORMAT = /\[([^\]]+)]|YYYY|YY|M{1,4}|D{1,2}|d{1,4}|H{1,2}|h{1,2}|a|A|m{1,2}|s{1,2}|Z{1,2}|SSS/g;

// locale/en.js
const en = {
  weekdays: 'Sunday_Monday_Tuesday_Wednesday_Thursday_Friday_Saturday'.split('_'),
  months: 'January_February_March_April_May_June_July_August_September_October_November_December'.split('_'),
  ordinal(n) {
    const s = ['th', 'st', 'nd', 'rd'];
    const v = n % 100;
    return `[${n}${s[(v - 20) % 10] || s[v] || s[0]}]`;
  },
};

const padStart = (string, length, pad) => {
  const s = String(string);
  if (!s || s.length >= length) return string;
  return `${Array(length + 1 - s.length).join(pad)}${string}`;
};

// `$d.toString() !== 'Invalid Date'`, without formatting the date.
function isValid(d) {
  return !Number.isNaN(d.getTime());
}

// --------------------------------------------------------------------------
// format (index.js format, utc plugin)

function format(date, formatStr, utc) {
  const $d = new Date(date);
  if (!isValid($d)) return INVALID_DATE_STRING;
  const str = formatStr || (utc ? UTC_FORMAT_DEFAULT : '') || FORMAT_DEFAULT;
  const $y = utc ? $d.getUTCFullYear() : $d.getFullYear();
  const $M = utc ? $d.getUTCMonth() : $d.getMonth();
  const $D = utc ? $d.getUTCDate() : $d.getDate();
  const $W = utc ? $d.getUTCDay() : $d.getDay();
  const $H = utc ? $d.getUTCHours() : $d.getHours();
  const $m = utc ? $d.getUTCMinutes() : $d.getMinutes();
  const $s = utc ? $d.getUTCSeconds() : $d.getSeconds();
  const $ms = utc ? $d.getUTCMilliseconds() : $d.getMilliseconds();
  const utcOffset = utc ? 0 : -Math.round($d.getTimezoneOffset() / 15) * 15;
  const negMinutes = -utcOffset;
  const minutes = Math.abs(negMinutes);
  const zoneStr = `${negMinutes <= 0 ? '+' : '-'}${padStart(Math.floor(minutes / 60), 2, '0')}:${padStart(minutes % 60, 2, '0')}`;
  const {weekdays, months} = en;
  const getShort = (arr, index, full, length) => (arr && arr[index]) || full[index].slice(0, length);
  const get$H = num => padStart($H % 12 || 12, num, '0');
  const meridiemFunc = (hour, minute, isLowercase) => {
    const m = hour < 12 ? 'AM' : 'PM';
    return isLowercase ? m.toLowerCase() : m;
  };
  const matches = match => {
    switch (match) {
      case 'YY':
        return String($y).slice(-2);
      case 'YYYY':
        return padStart($y, 4, '0');
      case 'M':
        return $M + 1;
      case 'MM':
        return padStart($M + 1, 2, '0');
      case 'MMM':
        return getShort(undefined, $M, months, 3);
      case 'MMMM':
        return getShort(months, $M);
      case 'D':
        return $D;
      case 'DD':
        return padStart($D, 2, '0');
      case 'd':
        return String($W);
      case 'dd':
        return getShort(undefined, $W, weekdays, 2);
      case 'ddd':
        return getShort(undefined, $W, weekdays, 3);
      case 'dddd':
        return weekdays[$W];
      case 'H':
        return String($H);
      case 'HH':
        return padStart($H, 2, '0');
      case 'h':
        return get$H(1);
      case 'hh':
        return get$H(2);
      case 'a':
        return meridiemFunc($H, $m, true);
      case 'A':
        return meridiemFunc($H, $m, false);
      case 'm':
        return String($m);
      case 'mm':
        return padStart($m, 2, '0');
      case 's':
        return String($s);
      case 'ss':
        return padStart($s, 2, '0');
      case 'SSS':
        return padStart($ms, 3, '0');
      case 'Z':
        return zoneStr;
      default:
        break;
    }
    return null;
  };
  return str.replace(REGEX_FORMAT, (match, $1) => $1 || matches(match) || zoneStr.replace(':', ''));
}

// --------------------------------------------------------------------------
// customParseFormat

const formattingTokens = /(\[[^[]*\])|([-_:/.,()\s]+)|(A|a|Q|YYYY|YY?|ww?|MM?M?M?|Do|DD?|hh?|HH?|mm?|ss?|S{1,3}|z|ZZ?)/g;
const match1 = /\d/;
const match2 = /\d\d/;
const match3 = /\d{3}/;
const match4 = /\d{4}/;
const match1to2 = /\d\d?/;
const matchSigned = /[+-]?\d+/;
const matchOffset = /[+-]\d\d:?(\d\d)?|Z/;
const matchWord = /\d*[^-_:/,()\s\d]+/;

const parseTwoDigitYear = input => {
  input = +input;
  return input + (input > 68 ? 1900 : 2000);
};

function offsetFromString(string) {
  if (!string) return 0;
  if (string === 'Z') return 0;
  const parts = string.match(/([+-]|\d\d)/g);
  const minutes = +(parts[1] * 60) + (+parts[2] || 0);
  // eslint-disable-next-line no-nested-ternary
  return minutes === 0 ? 0 : parts[0] === '+' ? -minutes : minutes;
}

const addInput = property =>
  function(input) {
    this[property] = +input;
  };

const zoneExpressions = [
  matchOffset,
  function(input) {
    const zone = this.zone || (this.zone = {});
    zone.offset = offsetFromString(input);
  },
];

const meridiemMatch = (input, isLowerCase) => input === (isLowerCase ? 'pm' : 'PM');

const expressions = {
  A: [
    matchWord,
    function(input) {
      this.afternoon = meridiemMatch(input, false);
    },
  ],
  a: [
    matchWord,
    function(input) {
      this.afternoon = meridiemMatch(input, true);
    },
  ],
  Q: [
    match1,
    function(input) {
      this.month = (input - 1) * 3 + 1;
    },
  ],
  S: [
    match1,
    function(input) {
      this.milliseconds = +input * 100;
    },
  ],
  SS: [
    match2,
    function(input) {
      this.milliseconds = +input * 10;
    },
  ],
  SSS: [
    match3,
    function(input) {
      this.milliseconds = +input;
    },
  ],
  s: [match1to2, addInput('seconds')],
  ss: [match1to2, addInput('seconds')],
  m: [match1to2, addInput('minutes')],
  mm: [match1to2, addInput('minutes')],
  H: [match1to2, addInput('hours')],
  h: [match1to2, addInput('hours')],
  HH: [match1to2, addInput('hours')],
  hh: [match1to2, addInput('hours')],
  D: [match1to2, addInput('day')],
  DD: [match2, addInput('day')],
  Do: [
    matchWord,
    function(input) {
      const [day] = input.match(/\d+/);
      this.day = day;
      for (let i = 1; i <= 31; i += 1) {
        if (en.ordinal(i).replace(/\[|\]/g, '') === input) {
          this.day = i;
        }
      }
    },
  ],
  w: [match1to2, addInput('week')],
  ww: [match2, addInput('week')],
  M: [match1to2, addInput('month')],
  MM: [match2, addInput('month')],
  MMM: [
    matchWord,
    function(input) {
      const matchIndex = en.months.map(_ => _.slice(0, 3)).indexOf(input) + 1;
      if (matchIndex < 1) {
        return INVALID; // dayjs: throw new Error()
      }
      this.month = matchIndex % 12 || matchIndex;
    },
  ],
  MMMM: [
    matchWord,
    function(input) {
      const matchIndex = en.months.indexOf(input) + 1;
      if (matchIndex < 1) {
        return INVALID; // dayjs: throw new Error()
      }
      this.month = matchIndex % 12 || matchIndex;
    },
  ],
  Y: [matchSigned, addInput('year')],
  YY: [
    match2,
    function(input) {
      this.year = parseTwoDigitYear(input);
    },
  ],
  YYYY: [match4, addInput('year')],
  Z: zoneExpressions,
  ZZ: zoneExpressions,
};

function correctHours(time) {
  const {afternoon} = time;
  if (afternoon !== undefined) {
    const {hours} = time;
    if (afternoon) {
      if (hours < 12) {
        time.hours += 12;
      }
    } else if (hours === 12) {
      time.hours = 0;
    }
    delete time.afternoon;
  }
}

// localizedFormat's u(format, undefined): the 'en' locale has no `formats`,
// so a localized token (L, LT, ...) throws, which makes the date invalid.
function expandLocalized(formatStr) {
  return formatStr.replace(/(\[[^\]]+])|(LTS?|l{1,4}|L{1,4})/g, (_, a) => {
    if (a) return a;
    throw new TypeError("Cannot read properties of undefined (reading 'L')");
  });
}

function makeParser(formatStr) {
  formatStr = expandLocalized(formatStr);
  const array = formatStr.match(formattingTokens);
  const {length} = array;
  for (let i = 0; i < length; i += 1) {
    const token = array[i];
    const parseTo = expressions[token];
    const regex = parseTo && parseTo[0];
    const parser = parseTo && parseTo[1];
    if (parser) {
      array[i] = {regex, parser};
    } else {
      array[i] = token.replace(/^\[|\]$/g, '');
    }
  }
  return function(input) {
    const time = {};
    for (let i = 0, start = 0; i < length; i += 1) {
      const token = array[i];
      if (typeof token === 'string') {
        start += token.length;
      } else {
        const part = input.slice(start);
        const match = token.regex.exec(part);
        // dayjs reads match[0] of null here: a TypeError its caller turns
        // into an invalid date. Failing without the throw is much faster.
        if (match === null) return null;
        const value = match[0];
        if (token.parser.call(time, value) === INVALID) return null;
        input = input.replace(value, '');
      }
    }
    correctHours(time);
    return time;
  };
}

// makeParser is pure for the fixed 'en' locale: compile each format once.
const parsers = new Map();
function cachedParser(formatStr) {
  let parser = parsers.get(formatStr);
  if (parser === undefined) {
    parser = makeParser(formatStr);
    parsers.set(formatStr, parser);
  }
  return parser;
}

function parseFormattedInput(input, formatStr) {
  try {
    if (['x', 'X'].indexOf(formatStr) > -1) return new Date((formatStr === 'X' ? 1000 : 1) * input);
    const time = cachedParser(formatStr)(input);
    if (time === null) return new Date('');
    const {year, month, day, hours, minutes, seconds, milliseconds, zone, week} = time;
    const now = new Date();
    const d = day || (!year && !month ? now.getDate() : 1);
    const y = year || now.getFullYear();
    let M = 0;
    if (!(year && !month)) {
      M = month > 0 ? month - 1 : now.getMonth();
    }
    const h = hours || 0;
    const m = minutes || 0;
    const s = seconds || 0;
    const ms = milliseconds || 0;
    if (zone) {
      return new Date(Date.UTC(y, M, d, h, m, s, ms + zone.offset * 60 * 1000));
    }
    if (week) {
      // dayjs(newDate).week(...) needs the weekOfYear plugin ExcelJS does not load.
      throw new TypeError('week is not a function');
    }
    return new Date(y, M, d, h, m, s, ms);
  } catch (e) {
    return new Date('');
  }
}

// dayjs(input, format, true): a Date, possibly invalid.
function parseStrict(input, formatStr) {
  const d = parseFormattedInput(input, formatStr);
  // eslint-disable-next-line eqeqeq
  if (input != format(d, formatStr, false)) {
    return new Date('');
  }
  return d;
}

module.exports = {format, parseStrict, isValid};
