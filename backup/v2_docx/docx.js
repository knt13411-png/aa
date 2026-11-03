'use strict';
/* 최소 Word(.docx) 작성기: 문단·표·그림·머리글/바닥글. U.zip(무압축)으로 묶어 오프라인에서 만든다. */

const Docx = {};
const DX = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const PAGE_W = 9638; // A4(11906) − 좌우 여백 20mm×2, 단위 twip

/* run: 문자열 또는 {t, b, color, sz(반각 pt), i} */
Docx.run = function (r) {
  if (typeof r === 'string') r = { t: r };
  let pr = '';
  if (r.b) pr += '<w:b/><w:bCs/>';
  if (r.i) pr += '<w:i/>';
  if (r.color) pr += `<w:color w:val="${r.color}"/>`;
  if (r.sz) pr += `<w:sz w:val="${r.sz}"/><w:szCs w:val="${r.sz}"/>`;
  if (r.shd) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${r.shd}"/>`;
  return String(r.t).split('\n').map((line, i) =>
    `${i ? '<w:r><w:br/></w:r>' : ''}<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ''}<w:t xml:space="preserve">${DX(line)}</w:t></w:r>`).join('');
};

/* o: {style, align, left, hanging, before, after, keepNext, pageBreakBefore} */
Docx.para = function (runs, o) {
  o = o || {};
  let pr = '';
  if (o.style) pr += `<w:pStyle w:val="${o.style}"/>`;
  if (o.keepNext) pr += '<w:keepNext/>';
  if (o.pageBreakBefore) pr += '<w:pageBreakBefore/>';
  if (o.before != null || o.after != null) pr += `<w:spacing${o.before != null ? ` w:before="${o.before}"` : ''}${o.after != null ? ` w:after="${o.after}"` : ''}/>`;
  if (o.left || o.hanging) pr += `<w:ind w:left="${o.left || 0}"${o.hanging ? ` w:hanging="${o.hanging}"` : ''}/>`;
  if (o.align) pr += `<w:jc w:val="${o.align}"/>`;
  const body = (Array.isArray(runs) ? runs : [runs]).filter(x => x != null && x !== '').map(Docx.run).join('');
  return `<w:p>${pr ? `<w:pPr>${pr}</w:pPr>` : ''}${body}</w:p>`;
};

class DocBuilder {
  constructor(title) { this.title = title; this.body = []; this.media = []; this.figN = 0; this.tblN = 0; }

  p(runs, o) { this.body.push(Docx.para(runs, o)); return this; }
  title1(t) { return this.p({ t, b: true, sz: 36 }, { style: 'Title', align: 'center', after: 240 }); }
  h1(t) { return this.p(t, { style: 'Heading1', keepNext: true }); }
  h2(t) { return this.p(t, { style: 'Heading2', keepNext: true }); }
  /* 한국 군 보고서식 위계 기호: □ → ○ → - → · */
  item(level, runs, o) {
    const sym = ['□', '○', '-', '·'][level] || '·';
    const arr = Array.isArray(runs) ? runs : [runs];
    return this.p([sym + ' ', ...arr], Object.assign({ left: 200 + level * 340 + 300, hanging: 300, after: 60 }, o || {}));
  }
  note(t) { return this.p({ t, sz: 18, color: '595959' }, { after: 60 }); }
  caption(t) { return this.p({ t, sz: 18, b: true, color: '404040' }, { align: 'center', after: 200 }); }
  pageBreak() { this.body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'); return this; }

  /* rows: [[cell]], cell = 문자열 | {t, b, color, fill, align, span}. o: {widths(비율), header(첫 행 머리), fill(머리 배경), firstColFill, fontSz} */
  table(rows, o) {
    o = o || {};
    const n = Math.max(...rows.map(r => r.reduce((s, c) => s + ((c && c.span) || 1), 0)));
    const ws = o.widths || Array(n).fill(1);
    const tot = ws.reduce((a, b) => a + b, 0);
    const tw = ws.map(w => Math.round((PAGE_W * w) / tot));
    const sz = o.fontSz || 19;
    const border = k => `<w:${k} w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>`;
    let x = `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` +
      `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${tw.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
    rows.forEach((row, ri) => {
      const isHead = o.header !== false && ri === 0;
      x += `<w:tr>${isHead ? '<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' : '<w:trPr><w:cantSplit/></w:trPr>'}`;
      let ci = 0;
      for (let c of row) {
        if (c == null || typeof c !== 'object') c = { t: c == null ? '' : String(c) };
        const span = c.span || 1;
        const w = tw.slice(ci, ci + span).reduce((a, b) => a + b, 0);
        const fill = c.fill || (isHead ? o.fill || 'D9E2F3' : ci === 0 && o.firstColFill ? o.firstColFill : null);
        x += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}` +
          `${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}<w:vAlign w:val="center"/></w:tcPr>` +
          Docx.para({ t: c.t, b: c.b || isHead, color: c.color, sz }, { align: c.align || (isHead ? 'center' : 'left'), after: 0, before: 0 }) + '</w:tc>';
        ci += span;
      }
      x += '</w:tr>';
    });
    this.body.push(x + '</w:tbl>');
    return this.p('', { after: 120 });
  }

  /* 한 칸짜리 강조 상자 (보고 요지 등) */
  box(paras, fill) {
    const inner = paras.join('');
    this.body.push(`<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right'].map(k => `<w:${k} w:val="single" w:sz="8" w:space="0" w:color="1F3864"/>`).join('')}</w:tblBorders>` +
      `<w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="${PAGE_W}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${PAGE_W}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${fill || 'F2F5FA'}"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`);
    return this.p('', { after: 120 });
  }

  /* png: {data: Uint8Array, w, h}, widthCm: 문서상 너비 */
  image(png, widthCm) {
    const i = this.media.length + 1;
    this.media.push(png.data);
    const cx = Math.round((widthCm || 16) * 360000);
    const cy = Math.round((cx * png.h) / png.w);
    this.body.push(`<w:p><w:pPr><w:keepNext/><w:jc w:val="center"/><w:spacing w:before="60" w:after="40"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${i}" name="그림 ${i}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${i}" name="image${i}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rIdImg${i}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`);
    return this;
  }
  figure(png, cap, widthCm) { this.figN++; this.image(png, widthCm); return this.caption(`[그림 ${this.figN}] ${cap}`); }
  tcap(cap) { this.tblN++; return this.p({ t: `<표 ${this.tblN}> ${cap}`, b: true, sz: 19 }, { keepNext: true, after: 60, before: 80 }); }

  build() {
    const enc = new TextEncoder();
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    const doc = head + `<w:document ${NS}><w:body>${this.body.join('')}` +
      `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/>` +
      `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const header = head + `<w:hdr ${NS}>${Docx.para({ t: this.title, sz: 16, color: '7F7F7F' }, { align: 'right', after: 0 })}</w:hdr>`;
    const footer = head + `<w:ftr ${NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">- </w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> -</w:t></w:r></w:p></w:ftr>`;
    const font = '<w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕" w:eastAsia="맑은 고딕" w:cs="맑은 고딕"/>';
    const styles = head + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr>${font}<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1F3864"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="140"/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="2" w:color="1F3864"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="1F3864"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="220" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>` +
      `</w:styles>`;
    const rels = this.media.map((_, i) => `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${i + 1}.png"/>`).join('');
    const docRels = head + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>${rels}</Relationships>`;
    const ct = head + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
    const rootRels = head + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;
    const core = head + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${DX(this.title)}</dc:title><dc:creator>사후검토 분석기</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`;
    const files = [
      { path: '[Content_Types].xml', data: enc.encode(ct) },
      { path: '_rels/.rels', data: enc.encode(rootRels) },
      { path: 'docProps/core.xml', data: enc.encode(core) },
      { path: 'word/document.xml', data: enc.encode(doc) },
      { path: 'word/styles.xml', data: enc.encode(styles) },
      { path: 'word/header1.xml', data: enc.encode(header) },
      { path: 'word/footer1.xml', data: enc.encode(footer) },
      { path: 'word/_rels/document.xml.rels', data: enc.encode(docRels) },
      ...this.media.map((d, i) => ({ path: `word/media/image${i + 1}.png`, data: d })),
    ];
    const blob = U.zip(files);
    return new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }
}
Docx.Builder = DocBuilder;
