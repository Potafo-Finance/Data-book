(() => {
  const DataBook = window.DataBook || (window.DataBook = {});
  DataBook.orderSummary = {
    init() {
      const metricNames = [
        'Customer Receivable',
        'Coupon Spend',
        'Revenue from Customer',
        'Vendor Sale',
        'GST Collected'
      ];
      const fileInput = document.getElementById('order-file');
      const dropzone = document.getElementById('dropzone');
      const message = document.getElementById('upload-message');
      const selectedFile = document.getElementById('selected-file');
      const sheetWorkspace = document.getElementById('sheet-workspace');
      const sheetMeta = document.getElementById('sheet-meta');
      const metricsGrid = document.getElementById('metrics-grid');
      const journalDateInput = document.getElementById('journal-date');
      const journalDateCell = document.getElementById('journal-date-cell');
      const journalDebitTotal = document.getElementById('journal-debit-total');
      const journalCreditTotal = document.getElementById('journal-credit-total');
      const journalStatus = document.getElementById('journal-status');
      const journalOrderCount = document.getElementById('journal-order-count');
      const differenceRow = document.getElementById('difference-row');
      const differenceParticulars = document.getElementById('difference-particulars');
      const differenceSide = document.getElementById('difference-side');
      const differenceDebit = document.getElementById('difference-debit');
      const differenceCredit = document.getElementById('difference-credit');
      const dateSheetSelect = document.getElementById('date-sheet');
      const dateHeaderRowInput = document.getElementById('date-header-row');
      const dateColumnSelect = document.getElementById('date-column');
      const dateSourceStatus = document.getElementById('date-source-status');
      const extensions = new Set(['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods']);
      const legacyStorageKey = 'data-book.order-summary.choices.v1';
      const numberFormat = new Intl.NumberFormat(undefined, {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
      let workbook = null;
      let textRows = null;
      let sheetNames = [];
      let sheetCache = new Map();
      let configs = [];
      let currentFileKey = null;
      let currentFile = null;
      let choiceDirty = false;
      let choiceVersion = 0;
      let pendingSave = null;
      let saveTimer = null;
      let saveChain = Promise.resolve();
      const uploadedKeys = new Set();
      let journalDate = '';
      let journalDateManual = false;
      let dateSource = { sheetName: '', headerRow: 1, column: null };
      let metricValues = Array(metricNames.length).fill(null);
      let loadId = 0;
      let sheetJsPromise = null;

      function defaultState() {
        return { headerRow: 1, terms: [{ sign: 1, columns: [] }] };
      }

      function sanitizeState(state) {
        const rawTerms = Array.isArray(state?.terms) ? state.terms : [];
        const terms = rawTerms.slice(0, 100).map(term => {
          const selected = Array.isArray(term?.columns) ? term.columns : [term?.column];
          return {
            sign: term?.sign === -1 ? -1 : 1,
            columns: [...new Set(selected.filter(column => Number.isInteger(column) && column >= 0))]
          };
        });
        return {
          headerRow: Number.isInteger(state?.headerRow) && state.headerRow > 0 ? state.headerRow : 1,
          terms: terms.length ? terms : defaultState().terms
        };
      }

      function readLegacyChoices(fileName) {
        try {
          const saved = JSON.parse(window.localStorage.getItem(legacyStorageKey) || '{}');
          return saved && typeof saved === 'object' && !Array.isArray(saved) ? saved['file:' + fileName] : null;
        } catch {
          return null;
        }
      }

      function restoreConfigs(entry) {
        const saved = Array.isArray(entry) ? entry : entry?.metrics;
        return metricNames.map((_, index) => {
          const stored = Array.isArray(saved) ? saved[index] : null;
          const savedBySheet = Object.create(null);
          if (stored?.savedBySheet && typeof stored.savedBySheet === 'object') {
            Object.entries(stored.savedBySheet).forEach(([name, state]) => {
              if (sheetNames.includes(name)) savedBySheet[name] = sanitizeState(state);
            });
          }
          const sheetName = sheetNames.includes(stored?.sheetName) ? stored.sheetName : sheetNames[0];
          const state = savedBySheet[sheetName] || defaultState();
          savedBySheet[sheetName] = sanitizeState(state);
          return { sheetName, headerRow: state.headerRow, terms: state.terms.map(term => ({ sign: term.sign, columns: [...term.columns] })), savedBySheet };
        });
      }

      function rememberCurrentSheet(config) {
        config.savedBySheet[config.sheetName] = {
          headerRow: config.headerRow,
          terms: config.terms.map(term => ({ sign: term.sign, columns: [...term.columns] }))
        };
      }

      function snapshotChoices() {
        configs.forEach(rememberCurrentSheet);
        return {
          journalDate,
          journalDateManual,
          dateSource: { ...dateSource },
          metrics: configs.map(config => ({
            sheetName: config.sheetName,
            savedBySheet: structuredClone(config.savedBySheet)
          }))
        };
      }

      function queuePendingSave() {
        const save = pendingSave;
        pendingSave = null;
        if (!save) return;
        saveChain = saveChain.catch(() => {}).then(async () => {
          if (DataBook.data.user()?.id !== save.userId) return;
          const uploadKey = save.userId + '/' + save.key;
          const record = await DataBook.data.saveRecord({
            module: 'order-summary', key: save.key, title: save.file.name,
            payload: { choices: save.choices },
            file: uploadedKeys.has(uploadKey) ? undefined : save.file
          });
          if (record?.file_path) uploadedKeys.add(uploadKey);
          if (save.loadId === loadId && save.version === choiceVersion && currentFileKey === save.key) {
            choiceDirty = false;
            message.textContent = 'File and choices saved in Data.';
            message.classList.remove('error');
          }
        }).catch(error => {
          if (save.loadId === loadId && save.version === choiceVersion && currentFileKey === save.key) {
            message.textContent = 'Could not save to Data: ' + (error.message || 'Please try again.');
            message.classList.add('error');
          }
        });
      }

      function scheduleSave(delay = 350) {
        if (!currentFileKey || !currentFile) return;
        const user = DataBook.data.user();
        if (!user) {
          message.textContent = 'File loaded locally. Sign in from Data to save it and its choices.';
          message.classList.remove('error');
          return;
        }
        pendingSave = {
          key: currentFileKey, file: currentFile, userId: user.id,
          choices: snapshotChoices(), loadId, version: choiceVersion
        };
        clearTimeout(saveTimer);
        saveTimer = setTimeout(queuePendingSave, delay);
      }

      function flushPendingSave() {
        if (!pendingSave) return;
        clearTimeout(saveTimer);
        queuePendingSave();
      }

      function saveConfigs() {
        if (!currentFileKey) return;
        choiceDirty = true;
        choiceVersion += 1;
        scheduleSave();
      }

      function make(tag, className, text) {
        const element = document.createElement(tag);
        if (className) element.className = className;
        if (text !== undefined) element.textContent = text;
        return element;
      }

      function excelColumn(index) {
        let number = index + 1;
        let label = '';
        while (number > 0) {
          number -= 1;
          label = String.fromCharCode(65 + number % 26) + label;
          number = Math.floor(number / 26);
        }
        return label;
      }

      function formatSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / 1048576).toFixed(1) + ' MB';
      }

      async function fileContentKey(bytes) {
        if (!window.crypto?.subtle) {
          let first = 2166136261;
          let second = 5381;
          for (const byte of new Uint8Array(bytes)) {
            first = Math.imul(first ^ byte, 16777619);
            second = Math.imul(second, 33) ^ byte;
          }
          return 'checksum:' + bytes.byteLength + ':' + (first >>> 0).toString(16) + ':' + (second >>> 0).toString(16);
        }
        const digest = await window.crypto.subtle.digest('SHA-256', bytes);
        return 'sha256:' + [...new Uint8Array(digest)].map(value => value.toString(16).padStart(2, '0')).join('');
      }

      function dateFromParts(year, month, day) {
        const iso = String(year).padStart(4, '0') + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
        const parsed = new Date(iso + 'T00:00:00Z');
        return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === iso ? iso : '';
      }

      function normalizeDate(value) {
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
          return dateFromParts(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
        }
        if (typeof value === 'number' && Number.isFinite(value) && value >= 20000 && value < 80000) {
          return normalizeDate(new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000));
        }
        const text = String(value || '').trim();
        const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (iso) return dateFromParts(iso[1], iso[2], iso[3]);
        const dayFirst = text.match(/^(\d{2})[/-](\d{2})[/-](\d{4})/);
        if (dayFirst) return dateFromParts(dayFirst[3], dayFirst[2], dayFirst[1]);
        if (/^\d{5}$/.test(text)) return normalizeDate(Number(text));
        return '';
      }

      function inferJournalDate(fileName) {
        const range = fileName.match(/(\d{4}-\d{2}-\d{2})-to-(\d{4}-\d{2}-\d{2})/);
        return range && range[1] === range[2] ? normalizeDate(range[1]) : '';
      }

      function renderJournal() {
        journalDateInput.disabled = !currentFileKey;
        journalDateCell.textContent = journalDate ? journalDate.slice(8, 10) + '/' + journalDate.slice(5, 7) + '/' + journalDate.slice(0, 4) : '—';
        const orderCount = countOrdersForDate();
        journalOrderCount.textContent = orderCount === null ? '—' : String(orderCount);
        metricValues.forEach((value, index) => {
          document.getElementById('summary-total-' + index).textContent = value === null ? '—' : numberFormat.format(value);
        });
        const debitReady = [0, 1].every(index => metricValues[index] !== null);
        const creditReady = [2, 3, 4].every(index => metricValues[index] !== null);
        const cents = indices => indices.reduce((sum, index) => sum + Math.round(metricValues[index] * 100), 0);
        const debitCents = debitReady ? cents([0, 1]) : null;
        const creditCents = creditReady ? cents([2, 3, 4]) : null;
        let balancedDebitCents = debitCents;
        let balancedCreditCents = creditCents;
        differenceRow.hidden = true;
        differenceRow.classList.remove('is-debit');
        differenceParticulars.textContent = '';
        differenceSide.textContent = '';
        differenceDebit.textContent = '';
        differenceCredit.textContent = '';
        journalStatus.classList.remove('is-balanced', 'is-unbalanced');
        if (debitReady && creditReady) {
          const difference = debitCents - creditCents;
          if (difference === 0) {
            journalStatus.textContent = 'Journal entry is balanced.';
            journalStatus.classList.add('is-balanced');
          } else if (difference > 0) {
            differenceRow.hidden = false;
            differenceParticulars.textContent = 'Difference in OS';
            differenceSide.textContent = 'Cr';
            differenceCredit.textContent = numberFormat.format(difference / 100);
            balancedCreditCents += difference;
            journalStatus.textContent = 'Journal balanced with Difference in OS Cr of ' + numberFormat.format(difference / 100) + '.';
            journalStatus.classList.add('is-unbalanced');
          } else {
            differenceRow.hidden = false;
            differenceRow.classList.add('is-debit');
            differenceParticulars.textContent = 'Difference in OS';
            differenceSide.textContent = 'Dr';
            differenceDebit.textContent = numberFormat.format(Math.abs(difference) / 100);
            balancedDebitCents += Math.abs(difference);
            journalStatus.textContent = 'Journal balanced with Difference in OS Dr of ' + numberFormat.format(Math.abs(difference) / 100) + '.';
            journalStatus.classList.add('is-unbalanced');
          }
        } else {
          journalStatus.textContent = !currentFileKey ? 'Upload a file to create the journal entry.' : !journalDate ? 'Choose an Entry date to calculate the journal.' : 'Complete all five calculations to check the journal balance.';
        }
        journalDebitTotal.textContent = debitReady ? numberFormat.format(balancedDebitCents / 100) : '—';
        journalCreditTotal.textContent = creditReady ? numberFormat.format(balancedCreditCents / 100) : '—';
      }

      function ensureSheetJS() {
        if (window.XLSX) return Promise.resolve(window.XLSX);
        if (!sheetJsPromise) {
          sheetJsPromise = new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
            script.onload = () => window.XLSX ? resolve(window.XLSX) : reject(new Error('Excel support did not load.'));
            script.onerror = () => reject(new Error('Excel support could not load. Connect to the internet and try again.'));
            document.head.append(script);
          }).catch(error => {
            sheetJsPromise = null;
            throw error;
          });
        }
        return sheetJsPromise;
      }

      function parseDelimited(text, delimiter) {
        const result = [];
        let row = [];
        let field = '';
        let quoted = false;
        text = text.replace(/^\uFEFF/, '');
        for (let i = 0; i < text.length; i += 1) {
          const character = text[i];
          if (character === '"') {
            if (quoted && text[i + 1] === '"') {
              field += '"';
              i += 1;
            } else if (quoted) quoted = false;
            else if (field === '') quoted = true;
            else field += '"';
          } else if (character === delimiter && !quoted) {
            row.push(field);
            field = '';
          } else if ((character === '\n' || character === '\r') && !quoted) {
            if (character === '\r' && text[i + 1] === '\n') i += 1;
            row.push(field);
            result.push(row);
            row = [];
            field = '';
          } else {
            field += character;
          }
        }
        if (field !== '' || row.length) {
          row.push(field);
          result.push(row);
        }
        if (quoted) throw new Error('This CSV or TSV file has an unclosed quote.');
        return result;
      }

      function hasValues(row) {
        return row && row.some(value => value !== null && value !== undefined && String(value).trim() !== '');
      }

      function numericValue(value) {
        if (value === null || value === undefined || String(value).trim() === '') return { number: 0, invalid: false };
        if (typeof value === 'number') return { number: Number.isFinite(value) ? value : 0, invalid: !Number.isFinite(value) };
        if (value instanceof Date) return { number: 0, invalid: true };
        let text = String(value).trim();
        const negative = text.startsWith('(') && text.endsWith(')');
        if (negative) text = text.slice(1, -1);
        text = text.replace(/[,\s₹$£€]/g, '');
        if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)) return { number: 0, invalid: true };
        const number = Number(text);
        return { number: negative ? -number : number, invalid: !Number.isFinite(number) };
      }

      function getSheetRows(name) {
        if (sheetCache.has(name)) return sheetCache.get(name);
        let rows = textRows || [];
        if (workbook) {
          const sheet = workbook.Sheets[name];
          if (sheet && sheet['!ref']) {
            const range = XLSX.utils.decode_range(sheet['!ref']);
            range.s.r = 0;
            range.s.c = 0;
            rows = XLSX.utils.sheet_to_json(sheet, {
              header: 1, defval: null, raw: true, blankrows: true,
              range: XLSX.utils.encode_range(range)
            });
          } else {
            rows = [];
          }
        }
        sheetCache.set(name, rows);
        return rows;
      }

      function getDateContext() {
        const rows = getSheetRows(dateSource.sheetName);
        dateSource.headerRow = Math.min(
          Math.max(Math.floor(Number(dateSource.headerRow)) || 1, 1),
          Math.max(rows.length, 1)
        );
        let count = 0;
        rows.forEach(row => { count = Math.max(count, row.length); });
        const header = rows[dateSource.headerRow - 1] || [];
        const columns = Array.from({ length: count }, (_, index) => {
          const letter = excelColumn(index);
          const heading = header[index] === null || header[index] === undefined ? '' : String(header[index]).trim();
          return { index, letter, heading, label: heading ? letter + ' — ' + heading.slice(0, 60) : 'Column ' + letter };
        });
        if (dateSource.column !== null && (!Number.isInteger(dateSource.column) || dateSource.column < 0 || dateSource.column >= columns.length)) dateSource.column = null;
        return { rows, columns, dataRows: rows.slice(dateSource.headerRow).filter(hasValues) };
      }

      function countOrdersForDate() {
        if (!currentFileKey || !journalDate || !sheetNames.includes(dateSource.sheetName) || dateSource.column === null) return null;
        const context = getDateContext();
        if (!context.columns[dateSource.column]) return null;
        const matchingRows = context.dataRows.filter(row => normalizeDate(row[dateSource.column]) === journalDate);
        const headings = context.columns.map(column => column.heading.toLowerCase().replace(/[._-]+/g, ' ').replace(/\s+/g, ' ').trim());
        const orderColumns = ['order number', 'order no', 'order num', 'order id']
          .map(name => headings.indexOf(name)).filter(index => index >= 0);
        const completeColumn = orderColumns.find(index => matchingRows.every(row => String(row[index] ?? '').trim() !== ''));
        if (completeColumn === undefined) return matchingRows.length;
        return new Set(matchingRows.map(row => String(row[completeColumn]).trim())).size;
      }

      function detectDateColumn(context) {
        const matches = context.columns.filter(column => /\bdate\b|created at/i.test(column.heading));
        return matches.length === 1 ? matches[0].index : null;
      }

      function renderDateSource() {
        if (!sheetNames.length) {
          dateSheetSelect.replaceChildren();
          dateColumnSelect.replaceChildren();
          dateSourceStatus.textContent = '';
          return '';
        }
        dateSheetSelect.replaceChildren();
        sheetNames.forEach(name => {
          const option = make('option', '', name);
          option.value = name;
          dateSheetSelect.append(option);
        });
        dateSheetSelect.value = dateSource.sheetName;
        const context = getDateContext();
        dateHeaderRowInput.max = String(Math.max(context.rows.length, 1));
        dateHeaderRowInput.value = String(dateSource.headerRow);
        dateColumnSelect.replaceChildren();
        const placeholder = make('option', '', 'Choose date column');
        placeholder.value = '';
        dateColumnSelect.append(placeholder);
        context.columns.forEach(column => {
          const option = make('option', '', column.label);
          option.value = String(column.index);
          dateColumnSelect.append(option);
        });
        dateColumnSelect.value = dateSource.column === null ? '' : String(dateSource.column);
        if (dateSource.column === null) {
          dateSourceStatus.textContent = 'Choose the column containing the journal date.';
          return '';
        }
        const parsedDates = context.dataRows.map(row => normalizeDate(row[dateSource.column])).filter(Boolean);
        const dates = new Set(parsedDates);
        if (!dates.size) {
          dateSourceStatus.textContent = 'No valid dates found in column ' + context.columns[dateSource.column].letter + '.';
          return '';
        }
        if (dates.size > 1) {
          dateSourceStatus.textContent = dates.size + ' different dates found. Choose the entry date in the journal above.';
          return '';
        }
        const found = [...dates][0];
        dateSourceStatus.textContent = 'Found ' + found.slice(8, 10) + '/' + found.slice(5, 7) + '/' + found.slice(0, 4) + ' in ' + parsedDates.length + ' rows.';
        return found;
      }

      function restoreDateSource(entry) {
        const stored = !Array.isArray(entry) ? entry?.dateSource : null;
        const validSheet = sheetNames.includes(stored?.sheetName);
        dateSource = {
          sheetName: validSheet ? stored.sheetName : sheetNames[0],
          headerRow: Number.isInteger(stored?.headerRow) && stored.headerRow > 0 ? stored.headerRow : 1,
          column: Number.isInteger(stored?.column) && stored.column >= 0 ? stored.column : null
        };
        const context = getDateContext();
        if (!validSheet || !stored || !Object.prototype.hasOwnProperty.call(stored, 'column')) dateSource.column = detectDateColumn(context);
        return renderDateSource();
      }

      function applyChoices(entry, fileName) {
        const foundDate = restoreDateSource(entry);
        const savedDate = normalizeDate(!Array.isArray(entry) && entry?.journalDate);
        const storedSource = !Array.isArray(entry) && entry?.dateSource;
        const explicitSource = storedSource && Object.prototype.hasOwnProperty.call(storedSource, 'column');
        journalDateManual = Boolean(!Array.isArray(entry) && (entry?.journalDateManual === true || (!storedSource && savedDate)));
        journalDate = journalDateManual ? savedDate : dateSource.column !== null ? foundDate : explicitSource ? '' : inferJournalDate(fileName);
        journalDateInput.value = journalDate;
        configs = restoreConfigs(entry);
        renderMetrics();
      }

      function applyDateSourceChange() {
        journalDate = renderDateSource();
        journalDateManual = false;
        journalDateInput.value = journalDate;
        if (configs.length) renderMetrics();
        else renderJournal();
        saveConfigs();
      }

      function getContext(config) {
        const rows = getSheetRows(config.sheetName);
        config.headerRow = Math.min(
          Math.max(Math.floor(Number(config.headerRow)) || 1, 1),
          Math.max(rows.length, 1)
        );
        let count = 0;
        rows.forEach(row => { count = Math.max(count, row.length); });
        const header = rows[config.headerRow - 1] || [];
        const columns = Array.from({ length: count }, (_, index) => {
          const letter = excelColumn(index);
          const heading = header[index] === null || header[index] === undefined ? '' : String(header[index]).trim();
          return { index, letter, heading, label: heading ? letter + ' — ' + heading.slice(0, 60) : 'Column ' + letter };
        });
        config.terms.forEach(term => {
          term.columns = term.columns.filter(column => Number.isInteger(column) && column >= 0 && column < columns.length);
        });
        const allDataRows = rows.slice(config.headerRow).filter(hasValues);
        const dateColumn = config.sheetName === dateSource.sheetName ? dateSource.column : detectDateColumn({ columns });
        let dateIssue = '';
        if (!journalDate) dateIssue = 'Choose an entry date to calculate.';
        else if (dateColumn === null) dateIssue = config.sheetName === dateSource.sheetName ? 'Choose the Date column above.' : 'No unique Date column found in this sheet.';
        else if (!allDataRows.some(row => normalizeDate(row[dateColumn]))) dateIssue = 'No valid dates found in this sheet.';
        const dataRows = dateIssue ? [] : allDataRows.filter(row => normalizeDate(row[dateColumn]) === journalDate);
        return { rows, columns, allDataRows, dataRows, dateIssue };
      }

      function calculate(terms, dataRows) {
        let total = 0;
        let ignored = 0;
        dataRows.forEach(row => {
          terms.forEach(term => {
            term.columns.forEach(column => {
              const value = numericValue(row[column]);
              total += term.sign * value.number;
              if (value.invalid) ignored += 1;
            });
          });
        });
        return { total, ignored };
      }

      function updateMetricResult(card, terms, context, metricIndex) {
        const total = card.querySelector('.metric-total');
        const preview = card.querySelector('.formula-preview');
        const note = card.querySelector('.calculation-note');
        const activeTerms = terms.filter(term => term.columns.length > 0);
        if (!activeTerms.length) {
          total.textContent = '—';
          metricValues[metricIndex] = null;
          preview.textContent = 'Choose columns to calculate.';
          note.textContent = context.dateIssue;
          renderJournal();
          return;
        }
        const signedColumns = activeTerms.flatMap(term => term.columns.map(column => ({
          sign: term.sign,
          letter: context.columns[column].letter
        })));
        const expression = signedColumns.map((entry, index) => {
          if (index === 0) return entry.sign < 0 ? '−' + entry.letter : entry.letter;
          return (entry.sign < 0 ? ' − ' : ' + ') + entry.letter;
        }).join('');
        preview.textContent = expression;
        if (context.dateIssue) {
          total.textContent = '—';
          metricValues[metricIndex] = null;
          note.textContent = context.dateIssue;
          renderJournal();
          return;
        }
        const result = calculate(activeTerms, context.dataRows);
        total.textContent = numberFormat.format(result.total);
        metricValues[metricIndex] = result.total;
        note.textContent = context.dataRows.length === 0 ? 'No rows match the entry date.' : result.ignored ? result.ignored + ' nonnumeric values ignored' : '';
        renderJournal();
      }

      function columnSelectionText(term, context) {
        const letters = term.columns.map(index => context.columns[index]?.letter).filter(Boolean);
        if (!letters.length) return 'Choose columns';
        if (letters.length <= 3) return letters.join(', ');
        return letters.slice(0, 2).join(', ') + ' +' + (letters.length - 2) + ' more';
      }

      function renderMetrics() {
        metricsGrid.replaceChildren();
        metricNames.forEach((name, metricIndex) => {
          const config = configs[metricIndex];
          const context = getContext(config);
          const terms = config.terms;
          const card = make('article', 'card metric-card');
          const heading = make('div', 'metric-heading');
          heading.append(make('span', 'metric-number', String(metricIndex + 1).padStart(2, '0')));
          heading.append(make('h3', '', name));
          card.append(heading);

          const source = make('div', 'sheet-controls metric-source');
          const sheetLabel = make('label', '', 'Sheet');
          const sheetSelect = make('select');
          sheetSelect.setAttribute('aria-label', name + ' sheet');
          sheetNames.forEach(sheetName => {
            const option = make('option', '', sheetName);
            option.value = sheetName;
            sheetSelect.append(option);
          });
          sheetSelect.value = config.sheetName;
          sheetSelect.addEventListener('change', () => {
            rememberCurrentSheet(config);
            config.sheetName = sheetSelect.value;
            const savedState = sanitizeState(config.savedBySheet[config.sheetName]);
            config.headerRow = savedState.headerRow;
            config.terms = savedState.terms;
            renderMetrics();
            saveConfigs();
            metricsGrid.children[metricIndex].querySelector('.metric-source select').focus();
          });
          sheetLabel.append(sheetSelect);

          const headerLabel = make('label', '', 'Header row');
          const headerInput = make('input');
          headerInput.type = 'number';
          headerInput.min = '1';
          headerInput.max = String(Math.max(context.rows.length, 1));
          headerInput.step = '1';
          headerInput.value = String(config.headerRow);
          headerInput.setAttribute('aria-label', name + ' header row');
          headerInput.addEventListener('change', () => {
            config.headerRow = headerInput.value;
            renderMetrics();
            saveConfigs();
            metricsGrid.children[metricIndex].querySelector('.metric-source input').focus();
          });
          headerLabel.append(headerInput);
          source.append(sheetLabel, headerLabel);
          card.append(source);
          const dateLabel = journalDate ? journalDate.slice(8, 10) + '/' + journalDate.slice(5, 7) + '/' + journalDate.slice(0, 4) : 'no date';
          const rowSummary = context.dateIssue ? context.dateIssue : context.dataRows.length + ' of ' + context.allDataRows.length + ' rows on ' + dateLabel;
          card.append(make('p', 'metric-data-meta', rowSummary + ' · ' + context.columns.length + ' columns'));
          card.append(make('div', 'metric-total', '—'));
          card.append(make('p', 'metric-caption', 'Calculated total'));
          card.append(make('div', 'formula-label', 'COLUMN FORMULA'));
          const list = make('div', 'term-list');

          terms.forEach((term, termIndex) => {
            const termRow = make('div', 'term-row');
            const signButton = make('button', 'operator-toggle', term.sign < 0 ? '−' : '+');
            signButton.type = 'button';
            signButton.classList.toggle('negative', term.sign < 0);
            signButton.setAttribute('aria-pressed', String(term.sign < 0));
            signButton.setAttribute('aria-label', name + ', term ' + (termIndex + 1) + ': ' + (term.sign < 0 ? 'subtract' : 'add') + '. Click to change.');
            signButton.addEventListener('click', () => {
              term.sign *= -1;
              signButton.textContent = term.sign < 0 ? '−' : '+';
              signButton.classList.toggle('negative', term.sign < 0);
              signButton.setAttribute('aria-pressed', String(term.sign < 0));
              signButton.setAttribute('aria-label', name + ', term ' + (termIndex + 1) + ': ' + (term.sign < 0 ? 'subtract' : 'add') + '. Click to change.');
              updateMetricResult(card, terms, context, metricIndex);
              saveConfigs();
            });

            const columnPicker = make('details', 'column-picker');
            const columnSummary = make('summary', '', columnSelectionText(term, context));
            columnSummary.setAttribute('aria-label', name + ', term ' + (termIndex + 1) + ': ' + columnSelectionText(term, context));
            const columnOptions = make('div', 'column-options');
            const optionList = make('div', 'column-option-list');
            const optionFooter = make('div', 'column-picker-footer');
            const selectedCount = make('span', 'selection-count', term.columns.length + ' selected');
            selectedCount.setAttribute('aria-live', 'polite');
            const saveButton = make('button', 'save-columns', 'Save columns');
            saveButton.type = 'button';
            saveButton.setAttribute('aria-label', 'Save columns for ' + name + ', term ' + (termIndex + 1));
            let draftColumns = [...term.columns];
            const checkboxes = [];
            context.columns.forEach(column => {
              const optionLabel = make('label', 'column-option');
              const checkbox = make('input');
              checkbox.type = 'checkbox';
              checkbox.value = String(column.index);
              checkbox.checked = term.columns.includes(column.index);
              checkboxes.push(checkbox);
              checkbox.addEventListener('change', () => {
                if (checkbox.checked) draftColumns = [...new Set([...draftColumns, column.index])].sort((a, b) => a - b);
                else draftColumns = draftColumns.filter(index => index !== column.index);
                selectedCount.textContent = draftColumns.length + ' selected';
              });
              optionLabel.append(checkbox, make('span', '', column.label));
              optionList.append(optionLabel);
            });
            if (!context.columns.length) optionList.append(make('p', 'empty-columns', 'No columns in this sheet.'));
            columnPicker.addEventListener('toggle', () => {
              draftColumns = [...term.columns];
              checkboxes.forEach(checkbox => { checkbox.checked = draftColumns.includes(Number(checkbox.value)); });
              selectedCount.textContent = draftColumns.length + ' selected';
            });
            saveButton.addEventListener('click', () => {
              term.columns = [...draftColumns];
              columnSummary.textContent = columnSelectionText(term, context);
              columnSummary.setAttribute('aria-label', name + ', term ' + (termIndex + 1) + ': ' + columnSelectionText(term, context));
              updateMetricResult(card, terms, context, metricIndex);
              saveConfigs();
              columnPicker.open = false;
              columnSummary.focus();
            });
            optionFooter.append(selectedCount, saveButton);
            columnOptions.append(optionList, optionFooter);
            columnPicker.append(columnSummary, columnOptions);

            const remove = make('button', 'term-remove', '×');
            remove.type = 'button';
            remove.setAttribute('aria-label', 'Remove term ' + (termIndex + 1) + ' from ' + name);
            remove.addEventListener('click', () => {
              if (terms.length === 1) terms[0] = { sign: 1, columns: [] };
              else terms.splice(termIndex, 1);
              renderMetrics();
              saveConfigs();
              metricsGrid.children[metricIndex].querySelector('.text-button').focus();
            });
            termRow.append(signButton, columnPicker, remove);
            list.append(termRow);
          });
          card.append(list);

          const actions = make('div', 'formula-actions');
          const add = make('button', 'text-button', '+ Add selection');
          add.type = 'button';
          add.addEventListener('click', () => {
            terms.push({ sign: 1, columns: [] });
            renderMetrics();
            saveConfigs();
            const pickers = metricsGrid.children[metricIndex].querySelectorAll('.column-picker');
            pickers[pickers.length - 1].querySelector('summary').focus();
          });
          actions.append(add);
          card.append(actions);
          card.append(make('p', 'formula-preview'));
          card.append(make('p', 'calculation-note'));
          metricsGrid.append(card);
          updateMetricResult(card, terms, context, metricIndex);
        });
      }

      function clearData(status) {
        flushPendingSave();
        loadId += 1;
        workbook = null;
        textRows = null;
        sheetNames = [];
        sheetCache = new Map();
        configs = [];
        currentFileKey = null;
        currentFile = null;
        choiceDirty = false;
        choiceVersion = 0;
        journalDate = '';
        journalDateManual = false;
        dateSource = { sheetName: '', headerRow: 1, column: null };
        renderDateSource();
        journalDateInput.value = '';
        metricValues = Array(metricNames.length).fill(null);
        fileInput.value = '';
        selectedFile.hidden = true;
        sheetWorkspace.hidden = true;
        renderJournal();
        message.textContent = status || '';
        message.classList.remove('error');
      }

      async function handleFile(file) {
        if (!file) return false;
        const extension = file.name.split('.').pop().toLowerCase();
        if (!extensions.has(extension)) {
          clearData('');
          message.textContent = 'Choose an XLSX, XLS, XLSM, CSV, TSV, or ODS file.';
          message.classList.add('error');
          return false;
        }
        flushPendingSave();
        const request = ++loadId;
        currentFileKey = null;
        currentFile = null;
        selectedFile.hidden = true;
        sheetWorkspace.hidden = true;
        metricValues = Array(metricNames.length).fill(null);
        journalDate = '';
        journalDateManual = false;
        journalDateInput.value = '';
        renderJournal();
        message.textContent = 'Reading file...';
        message.classList.remove('error');
        try {
          const bytes = await file.arrayBuffer();
          const recordKey = await fileContentKey(bytes);
          if (request !== loadId) return false;
          let nextWorkbook = null;
          let nextTextRows = null;
          if (extension === 'csv' || extension === 'tsv') {
            nextTextRows = parseDelimited(new TextDecoder().decode(bytes), extension === 'tsv' ? '\t' : ',');
            if (!nextTextRows.some(hasValues)) throw new Error('This file is empty.');
          } else {
            await ensureSheetJS();
            if (request !== loadId) return false;
            nextWorkbook = XLSX.read(bytes, { type: 'array', cellDates: true });
          }
          if (request !== loadId) return false;
          const nextNames = nextWorkbook ? nextWorkbook.SheetNames : [file.name];
          if (!nextNames.length) throw new Error('This file has no sheets.');
          let cloudError = null;
          try {
            await DataBook.data.ready();
          } catch (error) {
            cloudError = error;
          }
          if (request !== loadId) return false;
          let cloudRecord = null;
          const user = cloudError ? null : DataBook.data.user();
          if (user) {
            try {
              cloudRecord = await DataBook.data.getRecord('order-summary', recordKey);
            } catch (error) {
              cloudError = error;
            }
          }
          if (request !== loadId) return false;
          workbook = nextWorkbook;
          textRows = nextTextRows;
          sheetNames = nextNames;
          sheetCache = new Map();
          currentFileKey = recordKey;
          currentFile = file;
          choiceDirty = false;
          choiceVersion = 0;
          if (user && cloudRecord?.file_path) uploadedKeys.add(user.id + '/' + recordKey);
          const savedEntry = cloudRecord?.payload?.choices || readLegacyChoices(file.name);
          const hadSavedChoices = Boolean(savedEntry);
          sheetMeta.textContent = sheetNames.length + (sheetNames.length === 1 ? ' sheet' : ' sheets') + ' available · Choose a sheet in each calculation.';
          sheetWorkspace.hidden = false;
          applyChoices(savedEntry, file.name);
          document.getElementById('file-name').textContent = file.name;
          document.getElementById('file-size').textContent = formatSize(file.size);
          selectedFile.hidden = false;
          if (cloudError) {
            message.textContent = 'File loaded locally. Could not read Data: ' + (cloudError.message || 'Please try again.');
            message.classList.add('error');
          } else if (!user) {
            message.textContent = hadSavedChoices ? 'File loaded. Previous browser choices restored. Sign in from Data to save.' : 'File loaded locally. Sign in from Data to save it and its choices.';
          } else {
            message.textContent = hadSavedChoices ? 'File loaded. Saved choices restored.' : 'File loaded. Choose columns below.';
            if (!cloudRecord?.file_path) scheduleSave(0);
          }
          return true;
        } catch (error) {
          if (request !== loadId) return false;
          clearData('');
          message.textContent = error.message || 'Could not read this file.';
          message.classList.add('error');
          return false;
        }
      }

      async function reconcileSignedInFile(user) {
        if (!currentFile || !currentFileKey) return;
        const request = loadId;
        const key = currentFileKey;
        const file = currentFile;
        const dirtyAtStart = choiceDirty;
        try {
          const record = await DataBook.data.getRecord('order-summary', key);
          if (request !== loadId || DataBook.data.user()?.id !== user.id) return;
          if (record?.file_path) uploadedKeys.add(user.id + '/' + key);
          if (record?.payload?.choices && !dirtyAtStart && !choiceDirty) {
            applyChoices(record.payload.choices, file.name);
            message.textContent = 'File loaded. Saved choices restored from Data.';
            message.classList.remove('error');
          } else {
            scheduleSave(0);
          }
        } catch (error) {
          if (request !== loadId) return;
          message.textContent = 'Could not read Data: ' + (error.message || 'Please try again.');
          message.classList.add('error');
        }
      }

      DataBook.data.onChange(({ type, user, record }) => {
        if (type === 'deleted' && record?.module === 'order-summary' && record.record_key === currentFileKey && user) {
          if (pendingSave?.key === currentFileKey) {
            clearTimeout(saveTimer);
            pendingSave = null;
          }
          uploadedKeys.delete(user.id + '/' + currentFileKey);
          return;
        }
        if (type !== 'auth' || !currentFile) return;
        if (user) void reconcileSignedInFile(user);
        else {
          message.textContent = 'Signed out. Sign in from Data to save further changes.';
          message.classList.remove('error');
        }
      });
      DataBook.data.registerOpener('order-summary', async record => {
        if (!record.file_path) throw new Error('This saved item has no order file.');
        const blob = await DataBook.data.downloadFile(record);
        const name = record.file_name || record.title || 'order-summary.csv';
        const file = new File([blob], name, { type: blob.type });
        if (!await handleFile(file)) throw new Error(message.textContent || 'Could not open the saved order file.');
      });

      journalDateInput.addEventListener('change', () => {
        journalDate = normalizeDate(journalDateInput.value);
        journalDateManual = true;
        journalDateInput.value = journalDate;
        if (configs.length) renderMetrics();
        else renderJournal();
        saveConfigs();
      });
      dateSheetSelect.addEventListener('change', () => {
        dateSource.sheetName = dateSheetSelect.value;
        dateSource.headerRow = 1;
        dateSource.column = detectDateColumn(getDateContext());
        applyDateSourceChange();
      });
      dateHeaderRowInput.addEventListener('change', () => {
        dateSource.headerRow = dateHeaderRowInput.value;
        const context = getDateContext();
        if (dateSource.column === null) dateSource.column = detectDateColumn(context);
        applyDateSourceChange();
      });
      dateColumnSelect.addEventListener('change', () => {
        dateSource.column = dateColumnSelect.value === '' ? null : Number(dateColumnSelect.value);
        applyDateSourceChange();
        dateColumnSelect.focus();
      });
      fileInput.addEventListener('change', () => handleFile(fileInput.files[0]));
      ['dragenter', 'dragover'].forEach(type => dropzone.addEventListener(type, event => {
        event.preventDefault();
        dropzone.classList.add('dragging');
      }));
      ['dragleave', 'drop'].forEach(type => dropzone.addEventListener(type, event => {
        event.preventDefault();
        dropzone.classList.remove('dragging');
      }));
      dropzone.addEventListener('drop', event => {
        if (event.dataTransfer.files.length !== 1) {
          clearData('');
          message.textContent = 'Drop one file at a time.';
          message.classList.add('error');
          return;
        }
        try {
          fileInput.files = event.dataTransfer.files;
          handleFile(fileInput.files[0]);
        } catch {
          clearData('');
          message.textContent = 'Drag and drop is unavailable here. Use browse files instead.';
          message.classList.add('error');
        }
      });
      document.getElementById('remove-file').addEventListener('click', () => {
        clearData('File removed.');
        fileInput.focus();
      });
      renderJournal();
    }
  };
})();
