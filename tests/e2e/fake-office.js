// Tiruan minimal Office.js (Excel & Word) untuk uji asap di browser.
// Memeriksa hal yang sering salah: dimensi array saat menulis, nama tabel, alamat.
(function () {
  const host = new URLSearchParams(location.search).get('fakeHost') || 'Excel';
  const log = (window.__officeLog = []);

  function colNum(l) {
    let n = 0;
    for (const ch of l.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function colLet(n) {
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function parse(addr) {
    const a = String(addr).split('!').pop().replace(/\$/g, '');
    const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!m) throw new Error('InvalidArgument: alamat ' + addr);
    const c1 = colNum(m[1]);
    const r1 = +m[2];
    const c2 = m[3] ? colNum(m[3]) : c1;
    const r2 = m[4] ? +m[4] : r1;
    return { r: Math.min(r1, r2) - 1, c: Math.min(c1, c2) - 1, rows: Math.abs(r2 - r1) + 1, cols: Math.abs(c2 - c1) + 1 };
  }
  // Objek serba-boleh: properti apa pun bisa dibaca/ditulis, method apa pun bisa dipanggil.
  const nested = () => {
    const store = {};
    const fn = function () {
      return nested();
    };
    return new Proxy(fn, {
      get(t, k) {
        if (k in store) return store[k];
        if (k === Symbol.toPrimitive) return () => 0;
        store[k] = nested();
        return store[k];
      },
      set(t, k, v) {
        store[k] = v;
        return true;
      },
      apply() {
        return nested();
      },
    });
  };

  class Range {
    constructor(sheet, r, c, rows, cols) {
      Object.assign(this, { sheet, r, c, rowCount: rows, columnCount: cols });
      this.format = nested();
      this.format.autofitColumns = () => {};
      this.conditionalFormats = { add: (type) => (log.push(['cf', type, this.address]), nested()) };
      this.dataValidation = { clear() {}, rule: null };
      this.sort = { apply: (fields) => log.push(['sort', this.address, JSON.stringify(fields)]) };
    }
    get rowIndex() {
      return this.r;
    }
    get columnIndex() {
      return this.c;
    }
    get address() {
      return `${this.sheet.name}!${colLet(this.c + 1)}${this.r + 1}:${colLet(this.c + this.columnCount)}${this.r + this.rowCount}`;
    }
    load() {
      return this;
    }
    _grid(fn) {
      return Array.from({ length: this.rowCount }, (_, i) => Array.from({ length: this.columnCount }, (_, j) => fn(this.sheet.cell(this.r + i, this.c + j))));
    }
    get values() {
      return this._grid((x) => (x.f && String(x.f).startsWith('=') ? 0 : x.f === undefined ? '' : x.f));
    }
    set values(v) {
      this._set(v, 'f');
    }
    get formulas() {
      return this._grid((x) => (x.f === undefined ? '' : x.f));
    }
    set formulas(v) {
      this._set(v, 'f');
    }
    get numberFormat() {
      return this._grid((x) => x.nf || 'General');
    }
    set numberFormat(v) {
      this._set(v, 'nf');
    }
    _set(v, key) {
      if (!Array.isArray(v) || v.length !== this.rowCount || v.some((r) => !Array.isArray(r) || r.length !== this.columnCount)) {
        throw new Error(`InvalidArgument: dimensi ${key} ${Array.isArray(v) ? v.length + 'x' + (v[0] || []).length : typeof v} tidak cocok dengan ${this.rowCount}x${this.columnCount} di ${this.address}`);
      }
      v.forEach((row, i) => row.forEach((val, j) => (this.sheet.cell(this.r + i, this.c + j)[key] = val)));
      if (key === 'f') log.push(['write', this.address]);
    }
    getCell(i, j) {
      return new Range(this.sheet, this.r + i, this.c + j, 1, 1);
    }
    getResizedRange(dr, dc) {
      if (this.rowCount + dr < 1 || this.columnCount + dc < 1) throw new Error('InvalidArgument: resize');
      return new Range(this.sheet, this.r, this.c, this.rowCount + dr, this.columnCount + dc);
    }
    getOffsetRange(dr, dc) {
      return new Range(this.sheet, this.r + dr, this.c + dc, this.rowCount, this.columnCount);
    }
    getRow(i) {
      return new Range(this.sheet, this.r + i, this.c, 1, this.columnCount);
    }
    get columns() {
      return { load() {}, items: Array.from({ length: this.columnCount }, (_, j) => new Range(this.sheet, this.r, this.c + j, this.rowCount, 1)) };
    }
    merge() {
      log.push(['merge', this.address]);
    }
    select() {}
    clear() {}
    insert() {}
    delete() {
      log.push(['delete', this.address]);
    }
    getTables() {
      return { load() {}, items: [] };
    }
    removeDuplicates() {
      return { load() {}, removed: 0, uniqueRemaining: this.rowCount };
    }
  }

  class Sheet {
    constructor(wb, name) {
      this.wb = wb;
      this.name = name;
      this.cells = new Map();
      this.visibility = 'Visible';
      this.tables = { add: (addr, h) => wb._addTable(this, addr, h) };
      this.charts = {
        add: (type, range) => {
          log.push(['chart', type, range.address]);
          return { title: {}, dataLabels: {}, setPosition() {} };
        },
      };
      this.freezePanes = { freezeRows: (n) => log.push(['freeze', n]), freezeColumns() {}, freezeAt() {}, unfreeze() {} };
      this.autoFilter = { apply: () => log.push(['filter']) };
    }
    cell(r, c) {
      const k = r + ',' + c;
      if (!this.cells.has(k)) this.cells.set(k, {});
      return this.cells.get(k);
    }
    load() {
      return this;
    }
    activate() {
      this.wb.active = this;
    }
    getRange(addr) {
      const p = parse(addr);
      return new Range(this, p.r, p.c, p.rows, p.cols);
    }
    _used() {
      let r1 = Infinity,
        c1 = Infinity,
        r2 = -1,
        c2 = -1;
      for (const [k, v] of this.cells) {
        if (v.f === undefined || v.f === '') continue;
        const [r, c] = k.split(',').map(Number);
        r1 = Math.min(r1, r);
        c1 = Math.min(c1, c);
        r2 = Math.max(r2, r);
        c2 = Math.max(c2, c);
      }
      return r2 < 0 ? null : new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    getUsedRange() {
      return this._used() || new Range(this, 0, 0, 1, 1);
    }
    getUsedRangeOrNullObject() {
      const u = this._used();
      if (!u) return { isNullObject: true, load() {} };
      u.isNullObject = false;
      return u;
    }
  }

  class Workbook {
    constructor() {
      this.sheets = [new Sheet(this, 'Sheet1')];
      this.active = this.sheets[0];
      this.tableNames = [];
      const wb = this;
      this.worksheets = {
        load() {},
        get items() {
          return wb.sheets;
        },
        add(name) {
          if (wb.sheets.some((s) => s.name.toLowerCase() === String(name).toLowerCase())) throw new Error('ItemAlreadyExists: ' + name);
          const s = new Sheet(wb, name);
          wb.sheets.push(s);
          log.push(['addSheet', name]);
          return s;
        },
        getItem(name) {
          const s = wb.sheets.find((x) => x.name === name);
          if (!s) throw new Error('ItemNotFound: ' + name);
          return s;
        },
        getItemOrNullObject(name) {
          const s = wb.sheets.find((x) => x.name === name);
          return s ? Object.assign(s, { isNullObject: false }) : { isNullObject: true };
        },
        getActiveWorksheet: () => this.active,
      };
      this.tables = {
        load() {},
        get items() {
          return wb.tableNames.map((n) => ({ name: n, getRange: () => ({ load() {}, address: '' }) }));
        },
      };
      this.names = { load() {}, items: [], add() {} };
      this.selection = null;
    }
    _addTable(sheet, addr, hasHeaders) {
      const range = typeof addr === 'string' ? sheet.getRange(addr) : addr;
      log.push(['table', range.address, hasHeaders]);
      const wb = this;
      const t = {
        _name: 'Table' + (wb.tableNames.length + 1),
        get name() {
          return this._name;
        },
        set name(v) {
          if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(v)) throw new Error('InvalidArgument: nama tabel ' + v);
          if (wb.tableNames.includes(v)) throw new Error('nama tabel dipakai: ' + v);
          wb.tableNames = wb.tableNames.filter((n) => n !== this._name);
          this._name = v;
          wb.tableNames.push(v);
        },
        style: '',
        getRange: () => range,
      };
      wb.tableNames.push(t._name);
      return t;
    }
    getSelectedRange() {
      return this.selection || this.active.getRange('A1');
    }
  }

  const workbook = new Workbook();
  window.__workbook = workbook;

  // ---------------- Word ----------------
  const doc = { html: [], text: '', ops: [] };
  window.__doc = doc;
  const wrange = (kind) => ({
    kind,
    text: '',
    load() {
      return this;
    },
    insertHtml(html, loc) {
      doc.html.push(html);
      doc.ops.push(['insertHtml', kind, loc]);
      return wrange('inserted');
    },
    insertText(t, loc) {
      doc.ops.push(['insertText', t, loc]);
      return wrange('inserted');
    },
    insertBreak(t) {
      doc.ops.push(['break', t]);
    },
    insertOoxml(x) {
      doc.ops.push(['ooxml', x.length]);
    },
    insertParagraph(t) {
      doc.ops.push(['paragraph', t]);
      return { font: {} };
    },
    select() {},
    clear() {},
    getRange() {
      return wrange('whole');
    },
    font: {},
    paragraphs: {
      load() {},
      items: [],
      getFirst: () => wrange('para'),
    },
    search(t) {
      return { load() {}, items: doc.text.includes(t) ? [wrange('found')] : [] };
    },
  });
  const body = wrange('body');
  const context = {
    sync: () => Promise.resolve(),
    workbook,
    document: {
      body,
      getSelection: () => wrange('selection'),
      sections: { load() {}, items: [{ getHeader: () => wrange('header'), getFooter: () => wrange('footer') }] },
      settings: {},
    },
  };

  window.Excel = { run: (fn) => fn(context) };
  window.Word = { run: (fn) => fn(context) };
  const settings = {};
  window.Office = {
    HostType: { Excel: 'Excel', Word: 'Word' },
    onReady(cb) {
      setTimeout(() => cb({ host }), 0);
      return Promise.resolve({ host });
    },
    context: { document: { settings: { get: (k) => settings[k], set: (k, v) => (settings[k] = v), saveAsync: (cb) => cb && cb() } } },
  };
})();
