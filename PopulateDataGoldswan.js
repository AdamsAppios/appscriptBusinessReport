function handleGoldswanPopulate(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'Goldswan') return;

    const HEADER_ROW = 2;
    const row = e.range.getRow();
    if (row <= HEADER_ROW) return;

    // Re-entry guard used when we clear the paste cell
    const sp = PropertiesService.getScriptProperties();
    if (sp.getProperty('GS_POPULATE_SUSPEND') === '1') return;

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

    const headerToCol = {};
    headers.forEach((h, i) => {
      if (h && String(h).trim() !== '') headerToCol[String(h).trim()] = i + 1;
    });

    const pasteCol = headerToCol['PasteStringToPopulate'];
    if (!pasteCol || e.range.getColumn() !== pasteCol) return;

    const raw = String(e.value || '').trim();
    if (!raw) return;

    // ---------- helpers ----------
    const toNumber = (s) => {
      if (s == null) return null;
      const n = Number(String(s).replace(/,/g, '').trim());
      return isFinite(n) ? n : null;
    };

    // get first number appearing *after* the keyword
    const firstNumberAfterKey = (line, keyRe) => {
      const m = line.match(keyRe);
      if (!m) return null;
      const tail = line.slice(m.index + m[0].length);
      const m2 = tail.match(/(\d[\d,]*)/);  // first number after the keyword (commas ok)
      return m2 ? toNumber(m2[1]) : null;
    };

    // sum all numbers appearing *after* the keyword
    const sumNumbersAfterKey = (line, keyRe) => {
      const m = line.match(keyRe);
      if (!m) return null;
      const tail = line.slice(m.index + m[0].length);
      const nums = tail.match(/(\d[\d,]*)/g) || [];
      const total = nums.reduce((acc, s) => acc + (toNumber(s) || 0), 0);
      return isFinite(total) ? total : null;
    };

    const lines = raw
      .split(/\r?\n/)
      .map((s) => s.trim())
      .filter(Boolean);

    const updates = [];

    // flexible keyword regexes (handle dash, colon, or spaces after the key)
    const P = [
      { re: /collectible\s*unpaid\b/i, header: 'Collectibles',   fn: firstNumberAfterKey },
      { re: /collectible\s*paid\b/i,   header: 'Paid',           fn: firstNumberAfterKey },
      { re: /\brectangle\b/i,          header: 'Rect',           fn: firstNumberAfterKey },
      { re: /\bround\b/i,              header: 'Rnd',            fn: firstNumberAfterKey },
      { re: /\bpick\s*up\b/i,          header: 'Pickup',         fn: firstNumberAfterKey },
      { re: /\bplus\s*1\b/i,           header: 'Plus 1',         fn: sumNumbersAfterKey },
      { re: /badger\s*meter\b/i,       header: 'Badger Meter',   fn: firstNumberAfterKey },
      { re: /total\s*sales\b/i,        header: 'T. Sales  Text', fn: firstNumberAfterKey },
      { re: /net\s*amount\b/i,         header: 'CTO Text',       fn: firstNumberAfterKey },
    ];

    // Duty: e.g., "nag duty lito" → lito
    const dutyLine = lines.find((l) => /duty/i.test(l));
    if (dutyLine) {
      const dutyTailMatch = dutyLine.match(/\bduty\b(.*)$/i);
      if (dutyTailMatch && headerToCol['Duty']) {
        const dutyTail = dutyTailMatch[1] || '';
        const names = dutyTail
          .replace(/[&,]/g, ' ')
          .split(/\s+/)
          .map((n) => n.replace(/[^A-Za-z]/g, '').trim())
          .filter(Boolean)
          .map((n) => n.charAt(0).toUpperCase() + n.slice(1).toLowerCase());

        if (names.length) {
          updates.push({ col: headerToCol['Duty'], value: names.join(', ') });
        }
      }
    }

    // Simple mappings
    for (const line of lines) {
      for (const p of P) {
        if (p.re.test(line)) {
          const value =
            p.fn === sumNumbersAfterKey
              ? sumNumbersAfterKey(line, p.re)
              : firstNumberAfterKey(line, p.re);
          const col = headerToCol[p.header];
          if (col && value != null) updates.push({ col, value });
        }
      }
    }

    // Expenses block
    const expStartIdx = lines.findIndex((l) => /^expenses\s*:?/i.test(l));
    if (expStartIdx !== -1) {
      const expenseLines = [];
      for (let i = expStartIdx + 1; i < lines.length; i++) {
        const ln = lines[i];
        // stop when another section begins
        if (/(^total\s*sales|^net\s*amount|^beg\s|^badger\s*meter|^end\b|^nag\s*duty)/i.test(ln)) break;
        // accept forms like "item-12", "item:12", or "item 12"
        if (/(?:-|:|\s)\s*\d/.test(ln)) expenseLines.push(ln);
        else break;
      }

      if (expenseLines.length) {
        // convert separators to '=' for the text version
        const expenseText = expenseLines
          .map((l) => l.replace(/(?:-|:|\s)\s*/, '=')) // first sep to '='
          .join(';');

        // sum numeric values (first number on each line)
        const expenseTotal = expenseLines.reduce((sum, l) => {
          const n =
            firstNumberAfterKey(l, /^(.*?)(?:-|:|\s)/i) ?? // generic: anything until first sep
            firstNumberAfterKey(l, /.+/i);
          return sum + (n || 0);
        }, 0);

        if (headerToCol['Expenses'])
          updates.push({ col: headerToCol['Expenses'], value: expenseText });
        if (headerToCol['TotalExpenses'])
          updates.push({ col: headerToCol['TotalExpenses'], value: expenseTotal });
      }
    }

    // Write all updates
    if (updates.length) {
      updates.forEach((u) => sheet.getRange(row, u.col).setValue(u.value));
    }

    // Clear the pasted cell WITHOUT retriggering onEdit
    sp.setProperty('GS_POPULATE_SUSPEND', '1');
    try {
      sheet.getRange(row, pasteCol).clearContent();
    } finally {
      // tiny delay then release
      Utilities.sleep(50);
      sp.deleteProperty('GS_POPULATE_SUSPEND');
    }
  } catch (err) {
    console.error('handleGoldswanPopulate error:', err && err.message);
  }
}
