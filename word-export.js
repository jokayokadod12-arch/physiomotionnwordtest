/* ============================================================
   word-export.js — تصدير خطة النظام الغذائي (Lifestyle) كملف Word (.docx)
   الاستخدام:  downloadPlanWord(plan)
   بيحمّل مكتبة docx (js/vendor/docx.umd.js) أول مرة بس عند الضغط على الزر.
   ============================================================ */
(function () {
  'use strict';

  function loadDocxLib() {
    return new Promise(function (resolve, reject) {
      if (window.docx) return resolve(window.docx);
      var s = document.createElement('script');
      s.src = 'js/vendor/docx.umd.js';
      s.onload = function () { window.docx ? resolve(window.docx) : reject(new Error('docx not found')); };
      s.onerror = function () { reject(new Error('Failed to load Word library')); };
      document.head.appendChild(s);
    });
  }

  // نص عادي: يشيل أي وسوم HTML ويحوّل الـ entities (&lt; &amp; ...) لحروفها
  function clean(v) {
    if (v === undefined || v === null) return '';
    var t = document.createElement('textarea');
    t.innerHTML = String(v).replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]*>/g, '');
    return t.value;
  }

  function safeName(s) {
    return clean(s).replace(/[\\/:*?"<>|]+/g, '').trim().replace(/\s+/g, '-') || 'Patient';
  }

  async function buildBlob(p) {
    var d = await loadDocxLib();
    var Document = d.Document, Packer = d.Packer, Paragraph = d.Paragraph, TextRun = d.TextRun,
        Table = d.Table, TableRow = d.TableRow, TableCell = d.TableCell, WidthType = d.WidthType,
        AlignmentType = d.AlignmentType, BorderStyle = d.BorderStyle, ShadingType = d.ShadingType,
        VerticalAlign = d.VerticalAlign, Footer = d.Footer;

    var GREEN = '5A7A62', DARK = '2D4A35', LIGHT = 'F0F7F2', WARNBG = 'FFF8E1', GREY = '666666';
    var W = 9746; // A4 content width (margins 1080)
    var FONT = 'Arial';

    var goalTxt = p.goal === 'lose' ? 'Weight Loss' : p.goal === 'gain' ? 'Weight Gain' : 'Maintain Weight';
    var diseases = (p.diseases || []).filter(function (x) { return x && x !== 'none'; }).map(clean).join(', ') || 'None';
    var macros = p.macros || {};

    var thin = { style: BorderStyle.SINGLE, size: 4, color: 'D0D7D2' };
    var borders = { top: thin, bottom: thin, left: thin, right: thin };
    var cm = { top: 80, bottom: 80, left: 110, right: 110 };

    function run(text, o) {
      o = o || {};
      return new TextRun({ text: String(text), font: FONT, size: o.size || 20, bold: !!o.bold, color: o.color || '222222' });
    }
    function para(children, o) {
      o = o || {};
      return new Paragraph({
        children: Array.isArray(children) ? children : [children],
        spacing: { before: o.before || 0, after: o.after === undefined ? 80 : o.after },
        alignment: o.align || AlignmentType.LEFT,
        shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
        border: o.border
      });
    }
    function h2(text) {
      return para(run(text, { bold: true, size: 26, color: DARK }), {
        before: 280, after: 120,
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: GREEN, space: 4 } }
      });
    }
    function cell(children, width, o) {
      o = o || {};
      return new TableCell({
        children: children, width: { size: width, type: WidthType.DXA }, borders: borders, margins: cm,
        verticalAlign: o.vAlign || VerticalAlign.TOP, rowSpan: o.rowSpan, columnSpan: o.columnSpan,
        shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined
      });
    }

    var kids = [];

    // العنوان
    kids.push(para(run('Lifestyle Modification Plan', { bold: true, size: 40, color: GREEN }), { after: 60 }));
    kids.push(para(run('Personalized Nutrition Plan', { size: 22, color: GREY }), { after: 160 }));

    // بيانات المريض
    function info(label, val) { return [run(label + ': ', { bold: true }), run(val)]; }
    var sep = function () { return run('   |   ', { color: 'AAAAAA' }); };
    kids.push(para([].concat(
      info('Patient', clean(p.patientName)), [sep()], info('Age', clean(p.age) + ' yrs'), [sep()],
      info('Gender', clean(p.gender)), [sep()], info('Weight', clean(p.weight) + ' kg'), [sep()],
      info('Height', clean(p.height) + ' cm')
    )));
    kids.push(para([].concat(
      info('Goal', goalTxt), [sep()], info('BMI', clean(p.bmi) + ' (' + clean(p.bmiCat) + ')')
    )));
    kids.push(para(info('Conditions', diseases)));
    kids.push(para([].concat(info('Prepared by Dr.', clean(p.doctorName)), [sep()], info('Date', clean(p.date))), { after: 160 }));

    // تحذير
    kids.push(para(run('Warning: This is a general nutritional guideline. Individual needs may vary. Consult a licensed dietitian.',
      { size: 18, color: '795548' }), { fill: WARNBG, after: 160 }));

    // التحليل الأيضي
    kids.push(h2('Metabolic Analysis'));
    var metrics = [
      ['BMR', clean(p.bmr) + ' kcal'], ['TDEE', clean(p.tdee) + ' kcal'], ['Daily Target', clean(p.target) + ' kcal'],
      ['Protein', clean(macros.protein) + ' g'], ['Carbs', clean(macros.carbs) + ' g'], ['Fat', clean(macros.fat) + ' g']
    ];
    var mw = Math.floor(W / 3);
    var mrows = [0, 3].map(function (st) {
      return new TableRow({
        children: metrics.slice(st, st + 3).map(function (m, i) {
          return cell([
            para(run(m[1], { bold: true, size: 28, color: GREEN }), { align: AlignmentType.CENTER, after: 20 }),
            para(run(m[0].toUpperCase(), { size: 16, color: GREY }), { align: AlignmentType.CENTER, after: 0 })
          ], i === 2 ? W - 2 * mw : mw, { fill: LIGHT });
        })
      });
    });
    kids.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [mw, mw, W - 2 * mw], rows: mrows }));

    // خطة الوجبات
    kids.push(h2('7-Day Meal Plan'));
    var cw = [1300, 1700, 5546, 1200];
    var head = new TableRow({
      tableHeader: true,
      children: ['Day', 'Meal', 'Food', 'Calories'].map(function (t, i) {
        return cell([para(run(t, { bold: true, size: 20, color: 'FFFFFF' }), { after: 0 })], cw[i], { fill: GREEN });
      })
    });
    var rows = [head];
    (p.meals || []).forEach(function (day, di) {
      var fill = di % 2 === 1 ? LIGHT : undefined;
      var ms = day.meals || [];
      ms.forEach(function (m, i) {
        var cs = [];
        if (i === 0) {
          cs.push(cell([
            para(run(clean(day.day), { bold: true, color: '006064' }), { after: 20 }),
            para(run(clean(day.totalCals) + ' kcal', { size: 16, color: GREY }), { after: 0 })
          ], cw[0], { rowSpan: ms.length, vAlign: VerticalAlign.CENTER, fill: 'E0F2F1' }));
        }
        cs.push(cell([para(run(clean(m.type), { bold: true }), { after: 0 })], cw[1], { fill: fill }));
        cs.push(cell([para(run(clean(m.items), { size: 19 }), { after: 0 })], cw[2], { fill: fill }));
        cs.push(cell([para(run(clean(m.cals) + ' kcal', { size: 19 }), { after: 0 })], cw[3], { fill: fill }));
        rows.push(new TableRow({ children: cs, cantSplit: true }));
      });
    });
    kids.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cw, rows: rows }));

    // النصائح
    kids.push(h2('Lifestyle Tips'));
    (p.tips || []).forEach(function (t) {
      kids.push(new Paragraph({
        keepNext: false,
        spacing: { before: 60, after: 100 },
        shading: { type: ShadingType.CLEAR, fill: LIGHT, color: 'auto' },
        border: { left: { style: BorderStyle.SINGLE, size: 24, color: GREEN, space: 8 } },
        children: [run(clean(t.title), { bold: true, color: DARK, size: 21 }), new TextRun({ text: '', break: 1 }), run(clean(t.txt), { size: 19 })]
      }));
    });

    var doc = new Document({
      creator: 'Physio Motion Analyzer',
      title: 'Lifestyle Plan - ' + clean(p.patientName),
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1080, right: 1080 } } },
        footers: {
          default: new Footer({
            children: [para(run('Physio Motion Analyzer  ·  Generated by Dr. ' + clean(p.doctorName) + '  ·  ' + clean(p.date),
              { size: 16, color: '888888' }), { align: AlignmentType.CENTER, after: 0 })]
          })
        },
        children: kids
      }]
    });
    return Packer.toBlob(doc);
  }

  window.downloadPlanWord = async function (plan, btn) {
    if (!plan) return;
    var old = btn ? btn.innerHTML : null;
    if (btn) { btn.disabled = true; btn.innerHTML = '⏳ Word...'; }
    try {
      var blob = await buildBlob(plan);
      var name = 'Nutrition-Plan-' + safeName(plan.patientName) + '-' + safeName(plan.date || '') + '.docx';
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = name;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    } catch (e) {
      console.error(e);
      alert('Could not create the Word file. Please try again.');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = old; }
    }
  };
  window.__buildPlanWordBlob = buildBlob; // للاختبار
})();
