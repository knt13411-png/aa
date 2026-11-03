'use strict';
/* 최소 PowerPoint(.pptx) 작성기: 16:9 슬라이드에 도형·글상자·표·그림을 절대 위치(인치)로 배치. U.zip(무압축)으로 오프라인 생성. */

const Pptx = {};
const PX = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const EMU = v => Math.round(v * 914400); // 인치 → EMU
const NS_P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const PFONT = '맑은 고딕';
const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const fillXml = c => (c ? `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` : '<a:noFill/>');

/* run: 문자열 | {t, b, color, sz(pt)} */
Pptx.run = function (r, d) {
  if (typeof r === 'string') r = { t: r };
  const sz = Math.round((r.sz || d.sz || 14) * 100);
  const color = r.color || d.color || '1A1A1A';
  const b = r.b != null ? r.b : d.b;
  return `<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="${sz}"${b ? ' b="1"' : ''} dirty="0">${fillXml(color)}<a:latin typeface="${PFONT}"/><a:ea typeface="${PFONT}"/></a:rPr><a:t>${PX(r.t)}</a:t></a:r>`;
};

/* para: 문자열 | {t | runs, b, color, sz, align(l|ctr|r), lvl(□○- 위계 0~2), after(pt)} */
Pptx.para = function (p, d) {
  if (typeof p === 'string') p = { t: p };
  const dd = Object.assign({}, d, { sz: p.sz || d.sz, color: p.color || d.color, b: p.b != null ? p.b : d.b });
  let runs = p.runs ? p.runs.slice() : [p.t == null ? '' : p.t];
  let ind = '';
  if (p.lvl != null) {
    const sym = ['□', '○', '-'][p.lvl] || '·';
    runs = [sym + ' '].concat(runs);
    const mar = 0.05 + p.lvl * 0.28 + 0.22;
    ind = ` marL="${EMU(mar)}" indent="${EMU(-0.22)}"`;
  }
  const after = p.after != null ? p.after : d.after != null ? d.after : 4;
  const pPr = `<a:pPr algn="${p.align || d.align || 'l'}"${ind}><a:lnSpc><a:spcPct val="${d.line || 110}000"/></a:lnSpc><a:spcAft><a:spcPts val="${after * 100}"/></a:spcAft><a:buNone/></a:pPr>`;
  return `<a:p>${pPr}${runs.filter(x => x != null && x !== '').map(r => Pptx.run(r, dd)).join('')}<a:endParaRPr lang="ko-KR" sz="${Math.round((dd.sz || 14) * 100)}" dirty="0"/></a:p>`;
};

class PptSlide {
  constructor(deck) { this.deck = deck; this.items = []; this.rels = []; this.nid = 2; }
  id() { return this.nid++; }

  /* 사각형. o: {fill, line, lw(pt), round} */
  rect(x, y, w, h, o) {
    o = o || {};
    const id = this.id();
    const ln = o.line ? `<a:ln w="${Math.round((o.lw || 1) * 12700)}">${fillXml(o.line)}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
    this.items.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="도형 ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${EMU(x)}" y="${EMU(y)}"/><a:ext cx="${EMU(w)}" cy="${EMU(h)}"/></a:xfrm>` +
      `<a:prstGeom prst="${o.round ? 'roundRect' : 'rect'}"><a:avLst>${o.round ? `<a:gd name="adj" fmla="val ${o.round}"/>` : ''}</a:avLst></a:prstGeom>${fillXml(o.fill)}${ln}</p:spPr></p:sp>`);
    return this;
  }

  /* 글상자. content: 문자열 | [para]. o: {sz, color, b, align, anchor(t|ctr|b), fill, line, inset, line(줄간격%)} */
  text(x, y, w, h, content, o) {
    o = o || {};
    const id = this.id();
    const paras = Array.isArray(content) ? content : [content];
    const ins = EMU(o.inset != null ? o.inset : 0.06);
    const ln = o.border ? `<a:ln w="12700">${fillXml(o.border)}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
    this.items.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="글상자 ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${EMU(x)}" y="${EMU(y)}"/><a:ext cx="${EMU(w)}" cy="${EMU(h)}"/></a:xfrm><a:prstGeom prst="${o.round ? 'roundRect' : 'rect'}"><a:avLst/></a:prstGeom>${fillXml(o.fill)}${ln}</p:spPr>` +
      `<p:txBody><a:bodyPr wrap="square" lIns="${ins}" tIns="${ins}" rIns="${ins}" bIns="${ins}" anchor="${o.anchor || 't'}" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/>` +
      paras.map(p => Pptx.para(p, { sz: o.sz, color: o.color, b: o.b, align: o.align, after: o.after, line: o.lineSp })).join('') + '</p:txBody></p:sp>');
    return this;
  }

  /* png: {data, w, h}. h 생략 시 비율 유지 */
  image(png, x, y, w, h) {
    const id = this.id();
    const rid = 'rIdImg' + (this.rels.length + 1);
    const name = this.deck.addMedia(png.data);
    this.rels.push({ id: rid, type: 'image', target: '../media/' + name });
    if (h == null) h = (w * png.h) / png.w;
    this.items.push(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="그림 ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="${EMU(x)}" y="${EMU(y)}"/><a:ext cx="${EMU(w)}" cy="${EMU(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
    return this;
  }

  /* rows: [[cell]], cell = 문자열 | {t, b, color, fill, align, span}. o: {widths(비율), rowH, sz, head(머리 배경), firstColFill} */
  table(x, y, w, rows, o) {
    o = o || {};
    const id = this.id();
    const n = Math.max(...rows.map(r => r.reduce((s, c) => s + ((c && c.span) || 1), 0)));
    const ws = o.widths || Array(n).fill(1);
    const tot = ws.reduce((a, b) => a + b, 0);
    const cw = ws.map(v => EMU((w * v) / tot));
    const rowH = o.rowH || 0.34;
    const sz = o.sz || 11;
    const line = k => `<a:${k} w="9525" cap="flat" cmpd="sng" algn="ctr">${fillXml('A6A6A6')}<a:prstDash val="solid"/></a:${k}>`;
    let body = '';
    rows.forEach((row, ri) => {
      const isHead = o.header !== false && ri === 0;
      body += `<a:tr h="${EMU(rowH)}">`;
      let ci = 0;
      for (let c of row) {
        if (c == null || typeof c !== 'object') c = { t: c == null ? '' : String(c) };
        const span = c.span || 1;
        const fill = c.fill || (isHead ? o.head || '1F3864' : ci === 0 && o.firstColFill ? o.firstColFill : 'FFFFFF');
        const color = c.color || (isHead ? 'FFFFFF' : '1A1A1A');
        const para = Pptx.para({ t: c.t, b: c.b || isHead, color, sz: c.sz || sz, align: c.align || (isHead ? 'ctr' : 'l') }, { after: 0, line: 100 });
        body += `<a:tc${span > 1 ? ` gridSpan="${span}"` : ''}><a:txBody><a:bodyPr/><a:lstStyle/>${para}</a:txBody>` +
          `<a:tcPr marL="${EMU(0.07)}" marR="${EMU(0.07)}" marT="${EMU(0.03)}" marB="${EMU(0.03)}" anchor="ctr">${line('lnL')}${line('lnR')}${line('lnT')}${line('lnB')}${fillXml(fill)}</a:tcPr></a:tc>`;
        for (let k = 1; k < span; k++) body += `<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR"/></a:p></a:txBody><a:tcPr/></a:tc>`;
        ci += span;
      }
      body += '</a:tr>';
    });
    this.items.push(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="표 ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="${EMU(x)}" y="${EMU(y)}"/><a:ext cx="${EMU(w)}" cy="${EMU(rowH * rows.length)}"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="0"/>` +
      `<a:tblGrid>${cw.map(v => `<a:gridCol w="${v}"/>`).join('')}</a:tblGrid>${body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
    return this;
  }

  xml() { return XML_HEAD + `<p:sld ${NS_P}><p:cSld><p:spTree>${GRP}${this.items.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`; }
  relsXml() {
    return XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      this.rels.map(r => `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${r.type}" Target="${r.target}"/>`).join('') + '</Relationships>';
  }
}

class PptDeck {
  constructor(title) { this.title = title; this.slides = []; this.media = []; this.W = 13.333; this.H = 7.5; }
  slide() { const s = new PptSlide(this); this.slides.push(s); return s; }
  addMedia(data) { this.media.push(data); return `image${this.media.length}.png`; }

  build() {
    const enc = new TextEncoder();
    const n = this.slides.length;
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const rel = (id, type, target) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    const rels = inner => XML_HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
    const CT = 'application/vnd.openxmlformats-officedocument.presentationml';

    const pres = XML_HEAD + `<p:presentation ${NS_P} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst>` +
      `<p:sldIdLst>${this.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rIdS${i + 1}"/>`).join('')}</p:sldIdLst>` +
      `<p:sldSz cx="${EMU(this.W)}" cy="${EMU(this.H)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
    const presRels = rels(rel('rIdM', 'slideMaster', 'slideMasters/slideMaster1.xml') + rel('rIdT', 'theme', 'theme/theme1.xml') +
      rel('rIdP', 'presProps', 'presProps.xml') + rel('rIdV', 'viewProps', 'viewProps.xml') + rel('rIdTS', 'tableStyles', 'tableStyles.xml') +
      this.slides.map((_, i) => rel(`rIdS${i + 1}`, 'slide', `slides/slide${i + 1}.xml`)).join(''));
    const master = XML_HEAD + `<p:sldMaster ${NS_P}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>${GRP}</p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`;
    const masterRels = rels(rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml') + rel('rId2', 'theme', '../theme/theme1.xml'));
    const layout = XML_HEAD + `<p:sldLayout ${NS_P} type="blank" preserve="1"><p:cSld name="빈 화면"><p:spTree>${GRP}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
    const layoutRels = rels(rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml'));
    const clr = (name, v) => `<a:${name}><a:srgbClr val="${v}"/></a:${name}>`;
    const fillSt = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    const theme = XML_HEAD + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="AAR"><a:themeElements>` +
      `<a:clrScheme name="AAR"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>${clr('dk2', '1F3864')}${clr('lt2', 'EEECE1')}` +
      `${clr('accent1', '2A5DB0')}${clr('accent2', 'C03A32')}${clr('accent3', '1A9B6C')}${clr('accent4', 'D79A00')}${clr('accent5', '7F7F7F')}${clr('accent6', 'C75B8F')}${clr('hlink', '0563C1')}${clr('folHlink', '954F72')}</a:clrScheme>` +
      `<a:fontScheme name="AAR"><a:majorFont><a:latin typeface="${PFONT}"/><a:ea typeface="${PFONT}"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${PFONT}"/><a:ea typeface="${PFONT}"/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="AAR"><a:fillStyleLst>${fillSt}${fillSt}${fillSt}</a:fillStyleLst>` +
      `<a:lnStyleLst>${[6350, 12700, 19050].map(w => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`).join('')}</a:lnStyleLst>` +
      `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
      `<a:bgFillStyleLst>${fillSt}${fillSt}${fillSt}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
    const presProps = XML_HEAD + `<p:presentationPr ${NS_P}/>`;
    const viewProps = XML_HEAD + `<p:viewPr ${NS_P}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;
    const tableStyles = XML_HEAD + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
    const core = XML_HEAD + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${PX(this.title)}</dc:title><dc:creator>사후검토 분석기</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`;
    const ct = XML_HEAD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentation.main+xml"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.slideMaster+xml"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${CT}.slideLayout+xml"/>` +
      this.slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${CT}.slide+xml"/>`).join('') +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
      `<Override PartName="/ppt/presProps.xml" ContentType="${CT}.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${CT}.viewProps+xml"/>` +
      `<Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.tableStyles+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
    const rootRels = rels(rel('rId1', 'officeDocument', 'ppt/presentation.xml') +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>');

    const files = [
      ['[Content_Types].xml', ct], ['_rels/.rels', rootRels], ['docProps/core.xml', core],
      ['ppt/presentation.xml', pres], ['ppt/_rels/presentation.xml.rels', presRels],
      ['ppt/slideMasters/slideMaster1.xml', master], ['ppt/slideMasters/_rels/slideMaster1.xml.rels', masterRels],
      ['ppt/slideLayouts/slideLayout1.xml', layout], ['ppt/slideLayouts/_rels/slideLayout1.xml.rels', layoutRels],
      ['ppt/theme/theme1.xml', theme], ['ppt/presProps.xml', presProps], ['ppt/viewProps.xml', viewProps], ['ppt/tableStyles.xml', tableStyles],
    ].map(([path, s]) => ({ path, data: enc.encode(s) }));
    this.slides.forEach((s, i) => {
      files.push({ path: `ppt/slides/slide${i + 1}.xml`, data: enc.encode(s.xml()) });
      files.push({ path: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: enc.encode(s.relsXml()) });
    });
    this.media.forEach((d, i) => files.push({ path: `ppt/media/image${i + 1}.png`, data: d }));
    if (!n) throw new Error('슬라이드가 없습니다.');
    return new Blob([U.zip(files)], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }
}
Pptx.Deck = PptDeck;
