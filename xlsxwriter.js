"use strict";
/* Простая запись .xlsx (JSZip): листы с заголовком, ширинами столбцов, закреплённой первой строкой.
   Все ячейки — текст, чтобы номера ККТ и S/N не превращались в 1,06E+11.
   sheets: [{name, headers:[...], widths:[...], rows:[[...]], wrap:[индексы столбцов с переносом], plain:true — без заголовка (инструкция)}] */
(function(){
  const esc = s => String(s==null?"":s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]))
                                        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  const col = n => { let s=""; n++; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };

  // стили: 0 обычный, 1 заголовок, 2 текст с рамкой, 3 текст с рамкой и переносом, 4 крупный синий, 5 синий жирный, 6 перенос без рамки
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="14"/><color rgb="FF1F5FA8"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FF1F5FA8"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F5FA8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFC7D1"/></left><right style="thin"><color rgb="FFBFC7D1"/></right><top style="thin"><color rgb="FFBFC7D1"/></top><bottom style="thin"><color rgb="FFBFC7D1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="49" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="49" fontId="3" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const BLANK_ROWS=30;   // пустые строки с рамкой под таблицей — удобно дописывать
  function cell(ref, v, s){
    if(v==null || v==="") return `<c r="${ref}" s="${s}"/>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  }
  function sheetXml(sh){
    const n=sh.plain ? 1 : sh.headers.length;
    const cols=(sh.widths||[]).map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("");
    let rows="", r=1;
    if(sh.plain){
      sh.rows.forEach(([text, kind])=>{ rows+=`<row r="${r}">${cell("A"+r, text, kind==="title"?4:kind==="head"?5:6)}</row>`; r++; });
    } else {
      rows+=`<row r="1" ht="32" customHeight="1">`+sh.headers.map((h,i)=>cell(col(i)+"1", h, 1)).join("")+`</row>`; r=2;
      const data=sh.rows.concat(Array.from({length:BLANK_ROWS},()=>[]));
      data.forEach(row=>{ let h=`<row r="${r}">`; for(let i=0;i<n;i++) h+=cell(col(i)+r, row[i], (sh.wrap||[]).includes(i)?3:2); rows+=h+"</row>"; r++; });
    }
    const pane = sh.plain ? "" : `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`;
    const filter = sh.plain || !sh.rows.length ? "" : `<autoFilter ref="A1:${col(n-1)}${sh.rows.length+1}"/>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${pane}${cols?`<cols>${cols}</cols>`:""}<sheetData>${rows}</sheetData>${filter}</worksheet>`;
  }

  window.buildXlsx = function(sheets){
    const z=new JSZip();
    z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${
      sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`);
    z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${
      sheets.map((s,i)=>`<sheet name="${esc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets></workbook>`);
    z.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${
      sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")
      }<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    z.file("xl/styles.xml", STYLES);
    sheets.forEach((s,i)=>z.file(`xl/worksheets/sheet${i+1}.xml`, sheetXml(s)));
    return z.generateAsync({type:"blob", compression:"DEFLATE", mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  };
})();
