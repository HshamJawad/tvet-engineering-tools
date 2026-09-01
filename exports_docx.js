/* ============================================================
   exports_docx.js — Training Resources Management Tool
   Word (.docx) export, English and Arabic
   ============================================================

   A classic script, not an ES module, because index.html loads its
   other libraries the same way and the buttons call global functions
   through inline onclick. Adding a module here would mean the export
   is unreachable from the markup that triggers it.

   It reads three globals defined in index.html: _t(), _isRTL() and
   currencyLabel(). Keeping the dictionary there rather than duplicating
   it here means a term is worded once — the Excel export, the interface
   and this file cannot drift into three spellings of the same label.

   ── WHY ARABIC NEEDS FOUR SEPARATE THINGS ────────────────────
   Each of these fixes a different failure, and three of the four look
   identical when they are missing — the text simply sits on the wrong
   side — which is why they are set together at one choke point rather
   than sprinkled across call sites:

     bidirectional  on a Paragraph  → <w:bidi/>, reading order
     visuallyRightToLeft on a Table → <w:bidiVisual/>, column order
     rightToLeft    on a TextRun    → <w:rtl/>, marks the run as
                                      complex script so Word applies
                                      the cs font and the Arabic
                                      proofing language
     columnWidths   on a Table      → <w:tblGrid>

   The last one is not about direction at all, but omitting it is the
   most destructive of the four: docx@7.8.2 builds <w:tblGrid> from
   columnWidths AND NOTHING ELSE, so leaving it out writes a grid of
   100 twips per column. Under <w:tblLayout w:type="fixed"/> that grid
   outranks every declared cell width and the table collapses — which
   reads on screen as an RTL bug and is not one.

   ── AND ONE THING THAT MUST *NOT* BE SET ─────────────────────
   No `alignment` for start-aligned text. ECMA-376 makes w:jc LOGICAL
   inside a bidi paragraph: in an RTL paragraph w:val="right" means END
   and renders at the LEFT edge. Setting RIGHT "to be safe" pushes every
   Arabic cell to the wrong side. Left unset, a paragraph sits at its
   natural start edge — right under RTL, left under LTR — in both
   languages, with no branching.
   ============================================================ */

(function () {
  'use strict';

  /* Arial rather than a prettier face: Word substitutes silently when a
     named font is missing, which is how a document ends up full of
     boxes on a machine that never had it. Arial ships everywhere and
     has full Arabic coverage. */
  const FONT_AR = 'Arial';
  const FONT_EN = 'Calibri';

  const HEADER_FILL = 'DCDCDC';

  /* Landscape A4 (16838 twips) less 2 × 720 margin ≈ 15398 usable.
     Declared once and used for BOTH the cells and the table grid, so
     the two can never disagree. The two cost columns are wide because
     they carry the currency code in their heading. */
  const COLS = [3100, 1350, 1200, 3600, 2100, 2100, 1948];
  const TABLE_W = COLS.reduce((a, b) => a + b, 0);

  /* Falls back to alert() only if index.html has not defined notify()
     — a half-loaded page should still be able to report a failure. */
  const _say  = (msg, type) => (typeof notify === 'function' ? notify(msg, type) : alert(msg));
  const _rtl  = () => (typeof _isRTL === 'function' ? _isRTL() : false);
  const _tr   = (k) => (typeof _t === 'function' ? _t(k) : k);
  const _cur  = () => (typeof currencyLabel === 'function' ? currencyLabel() : '');
  const _font = () => (_rtl() ? FONT_AR : FONT_EN);

  /* Filename: /[^a-z0-9]/gi turns an Arabic title into a row of
     underscores, so every Arabic export arrived named "______.docx".
     Keep Unicode letters and digits; strip only what a filesystem
     actually objects to. */
  function safeFilename(title, fallback) {
    const base = String(title || '')
      .replace(/[\\/:*?"<>|\u0000-\u001F]/g, '')
      .replace(/\s+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '')
      .slice(0, 80);
    return (base || fallback) + '.docx';
  }

  function num(v) {
    const n = Number(v || 0);
    return isFinite(n) ? n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0.00';
  }

  async function exportToDOCX() {
    if (typeof window.docx === 'undefined') {
      _say(_tr('msgDocxLoading'), 'error');
      return;
    }

    const {
      Document, Paragraph, TextRun, Table, TableRow, TableCell,
      WidthType, AlignmentType, ImageRun, Packer, ShadingType,
    } = window.docx;

    const rtl = _rtl();

    /* THE run constructor. Every run in this file goes through it, so a
       run added later cannot forget the complex-script mark. */
    const run = (text, opts) => {
      opts = opts || {};
      return new TextRun({
        text: String(text == null ? '' : text),
        bold: !!opts.bold,
        size: opts.size || 20,
        font: _font(),
        rightToLeft: rtl,
      });
    };

    /* THE cell-paragraph constructor. Note the absent `alignment` for
       the start case — see the header note; this is the single place
       that decision is made. */
    const cellPara = (text, opts) => {
      opts = opts || {};
      const p = {
        children: [run(text, opts)],
        bidirectional: rtl,
        spacing: { before: 30, after: 30 },
      };
      if (opts.center) p.alignment = AlignmentType.CENTER;
      return new Paragraph(p);
    };

    const cell = (text, opts) => {
      opts = opts || {};
      const tc = {
        children: Array.isArray(text) ? text : [cellPara(text, opts)],
        width: { size: opts.width, type: WidthType.DXA },
      };
      if (opts.fill) tc.shading = { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' };
      if (opts.span) tc.columnSpan = opts.span;
      return new TableCell(tc);
    };

    const heading = (text, size) => new Paragraph({
      children: [run(text, { bold: true, size: size })],
      alignment: AlignmentType.CENTER,
      bidirectional: rtl,
      spacing: { after: 200 },
    });

    const line = (text) => new Paragraph({
      children: [run(text, { size: 22 })],
      bidirectional: rtl,
      spacing: { after: 100 },
    });

    try {
      const orgName =
        (document.getElementById('orgName') || {}).value || '';
      const trainingTitle =
        (document.getElementById('trainingTitle') || {}).value || '';
      const preparedBy =
        (document.getElementById('preparedBy') || {}).value || '';
      const prepDate =
        (document.getElementById('prepDate') || {}).value || '';

      const children = [];

      if (orgName) children.push(heading(orgName, 32));
      if (trainingTitle) children.push(heading(trainingTitle, 28));
      if (preparedBy) children.push(line(_tr('docPreparedBy') + ': ' + preparedBy));
      if (prepDate) children.push(line(_tr('docDate') + ': ' + prepDate));
      children.push(new Paragraph({ text: '', bidirectional: rtl }));

      const code = _cur();
      /* An en dash, not parentheses. Brackets are MIRRORED characters
         under the bidi algorithm, so "(IQD)" inside an Arabic heading
         renders with its parentheses swapped and, once the heading
         wraps, the closing bracket lands on the line above its content.
         A dash is directionally neutral and reads the same either way. */
      const withCur = (label) => (code ? `${label} \u2013 ${code}` : label);

      const headers = [
        _tr('colItem'),
        _tr('colQty'),
        _tr('colUnit'),
        _tr('colSpec'),
        withCur(_tr('colUnitCost')),
        withCur(_tr('colTotal')),
        _tr('colImage'),
      ];

      const tableRows = [new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => cell(h, {
          bold: true, size: 22, width: COLS[i], fill: HEADER_FILL,
          center: i === 1 || i === 4 || i === 5 || i === 6,
        })),
      })];

      const rows = document.querySelectorAll('#tableBody tr');
      let grandTotal = 0;

      rows.forEach(tr => {
        const rowId = tr.id.split('-')[1];
        const preview = document.getElementById('imagePreview-' + rowId);
        const imageBase64 = preview && preview.dataset ? preview.dataset.image : null;

        const q = (sel) => (tr.querySelector(sel) || {}).value || '';
        const total = Number(q('.total-cost') || 0);
        grandTotal += isFinite(total) ? total : 0;

        let imageCell;
        if (imageBase64) {
          try {
            const b64 = String(imageBase64).split(',')[1];
            imageCell = cell([new Paragraph({
              children: [new ImageRun({
                data: Uint8Array.from(atob(b64), c => c.charCodeAt(0)),
                transformation: { width: 90, height: 90 },
              })],
              alignment: AlignmentType.CENTER,
              bidirectional: rtl,
            })], { width: COLS[6] });
          } catch (e) {
            console.error('Image skipped:', e);
          }
        }
        if (!imageCell) imageCell = cell(_tr('docNoImage'), { width: COLS[6], center: true });

        tableRows.push(new TableRow({
          children: [
            cell(q('.item-name'),     { width: COLS[0] }),
            /* Quantities and money are centred and never carry the
               complex-script mark implicitly — a price whose decimal
               separator has been mirrored is a different number to the
               reader, and that is the one error in a costing sheet that
               nobody catches by eye. */
            cell(q('.quantity'),      { width: COLS[1], center: true }),
            cell(q('.unit'),          { width: COLS[2] }),
            cell(q('.specification'), { width: COLS[3] }),
            cell(num(q('.unit-cost')), { width: COLS[4], center: true }),
            cell(num(total),           { width: COLS[5], center: true }),
            imageCell,
          ],
        }));
      });

      tableRows.push(new TableRow({
        children: [
          cell(withCur(_tr('docGrandTotal')), {
            bold: true, size: 22, span: 5,
            width: COLS[0] + COLS[1] + COLS[2] + COLS[3] + COLS[4],
            fill: HEADER_FILL,
          }),
          cell(num(grandTotal), { bold: true, size: 22, width: COLS[5], center: true, fill: HEADER_FILL }),
          cell('', { width: COLS[6], fill: HEADER_FILL }),
        ],
      }));

      children.push(new Table({
        visuallyRightToLeft: rtl,
        width: { size: TABLE_W, type: WidthType.DXA },
        columnWidths: COLS,
        layout: 'fixed',
        rows: tableRows,
      }));

      const doc = new Document({
        styles: { default: { document: { run: { font: _font() } } } },
        sections: [{
          properties: {
            page: {
              margin: { top: 720, right: 720, bottom: 720, left: 720 },
              size: { orientation: 'landscape' },
            },
            bidi: rtl,
          },
          children: children,
        }],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = safeFilename(trainingTitle, _tr('docFallbackName'));
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      _say(_tr('msgDocxOk'), 'success');

    } catch (error) {
      console.error('Export error:', error);
      _say(_tr('msgDocxError') + error.message, 'error');
    }
  }

  window.exportToDOCX = exportToDOCX;
})();
