// Documento em branco no padrão do Word em português, com os estilos principais.
import JSZip from 'jszip';

export const NS_DECL = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
  + 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
  + 'xmlns:o="urn:schemas-microsoft-com:office:office" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
  + 'xmlns:v="urn:schemas-microsoft-com:vml" '
  + 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
  + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:w10="urn:schemas-microsoft-com:office:word" '
  + 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
  + 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
  + 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
  + 'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
  + 'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
  + 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
  + 'mc:Ignorable="w14 w15 wp14"';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const CT = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  + '<Default Extension="xml" ContentType="application/xml"/>'
  + '<Default Extension="png" ContentType="image/png"/>'
  + '<Default Extension="jpeg" ContentType="image/jpeg"/>'
  + '<Default Extension="jpg" ContentType="image/jpeg"/>'
  + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
  + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
  + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
  + '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
  + '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
  + '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
  + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
  + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
  + '</Types>';

const ROOT_RELS = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
  + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
  + '</Relationships>';

const DOC_RELS = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
  + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
  + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
  + '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
  + '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
  + '</Relationships>';

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';

const DOCUMENT = XML + '<w:document ' + NS_DECL + '><w:body><w:p/>' + SECT + '</w:body></w:document>';

const B = (sides, attrs) => sides.map((s) => '<w:' + s + ' ' + attrs + '/>').join('');

const tableStyle = (id, name, prio, body) => '<w:style w:type="table" w:styleId="' + id + '"><w:name w:val="' + name + '"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="' + prio + '"/>' + body + '</w:style>';

export const STYLES = XML + '<w:styles ' + W + ' ' + R + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">'
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="pt-BR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
  + '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + [1, 2, 3, 4, 5, 6].map((n) => {
    const sz = { 1: 32, 2: 26, 3: 24, 4: 22, 5: 22, 6: 22 }[n];
    const color = n === 3 || n === 6 ? '<w:color w:val="1F3763" w:themeColor="accent1" w:themeShade="7F"/>' : '<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>';
    return '<w:style w:type="paragraph" w:styleId="Heading' + n + '"><w:name w:val="heading ' + n + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="Heading' + n + 'Char"/><w:uiPriority w:val="9"/>' + (n > 2 ? '<w:unhideWhenUsed/>' : '') + '<w:qFormat/>'
      + '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="' + (n === 1 ? 240 : 40) + '" w:after="0"/><w:outlineLvl w:val="' + (n - 1) + '"/></w:pPr>'
      + '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' + (n === 4 ? '<w:i/><w:iCs/>' : '') + color + '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/></w:rPr></w:style>';
  }).join('')
  + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:color w:val="5A5A5A" w:themeColor="text1" w:themeTint="A5"/><w:spacing w:val="15"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="NoSpacing"><w:name w:val="No Spacing"/><w:uiPriority w:val="1"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="IntenseQuote"><w:name w:val="Intense Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="30"/><w:qFormat/><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="4472C4" w:themeColor="accent1"/><w:bottom w:val="single" w:sz="4" w:space="10" w:color="4472C4" w:themeColor="accent1"/></w:pBdr><w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:link w:val="HeaderChar"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:link w:val="FooterChar"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:link w:val="FootnoteTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="44546A" w:themeColor="text2"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>'
  + [1, 2, 3].map((n) => '<w:style w:type="paragraph" w:styleId="TOC' + n + '"><w:name w:val="toc ' + n + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8494"/></w:tabs><w:spacing w:after="100"/>' + (n > 1 ? '<w:ind w:left="' + (n - 1) * 220 + '"/>' : '') + '</w:pPr></w:style>').join('')
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="40"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="113" w:right="113"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BlockQuote"><w:name w:val="Block Quote"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="40"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="D0D7DE"/></w:pBdr><w:ind w:left="340"/></w:pPr><w:rPr><w:color w:val="595959"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
  + [1, 2, 3, 4, 5, 6].map((n) => '<w:style w:type="character" w:customStyle="1" w:styleId="Heading' + n + 'Char"><w:name w:val="Heading ' + n + ' Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Heading' + n + '"/><w:uiPriority w:val="9"/></w:style>').join('')
  + '<w:style w:type="character" w:customStyle="1" w:styleId="HeaderChar"><w:name w:val="Header Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Header"/><w:uiPriority w:val="99"/></w:style>'
  + '<w:style w:type="character" w:customStyle="1" w:styleId="FooterChar"><w:name w:val="Footer Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Footer"/><w:uiPriority w:val="99"/></w:style>'
  + '<w:style w:type="character" w:customStyle="1" w:styleId="FootnoteTextChar"><w:name w:val="Footnote Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="FootnoteText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="FollowedHyperlink"><w:name w:val="FollowedHyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:color w:val="954F72" w:themeColor="followedHyperlink"/><w:u w:val="single"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="22"/><w:qFormat/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="20"/><w:qFormat/><w:rPr><w:i/><w:iCs/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="SubtleEmphasis"><w:name w:val="Subtle Emphasis"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="19"/><w:qFormat/><w:rPr><w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="IntenseEmphasis"><w:name w:val="Intense Emphasis"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="21"/><w:qFormat/><w:rPr><w:i/><w:iCs/><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="SubtleReference"><w:name w:val="Subtle Reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="31"/><w:qFormat/><w:rPr><w:smallCaps/><w:color w:val="5A5A5A" w:themeColor="text1" w:themeTint="A5"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="IntenseReference"><w:name w:val="Intense Reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="32"/><w:qFormat/><w:rPr><w:b/><w:bCs/><w:smallCaps/><w:color w:val="4472C4" w:themeColor="accent1"/><w:spacing w:val="5"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="BookTitle"><w:name w:val="Book Title"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="33"/><w:qFormat/><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:spacing w:val="5"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:customStyle="1" w:styleId="InlineCode"><w:name w:val="Inline Code"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="40"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
  + tableStyle('TableGrid', 'Table Grid', 39, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>' + B(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], 'w:val="single" w:sz="4" w:space="0" w:color="auto"') + '</w:tblBorders></w:tblPr>')
  + tableStyle('GridTable1Light', 'Grid Table 1 Light', 46, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>' + B(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], 'w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"') + '</w:tblBorders></w:tblPr>'
    + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="12" w:space="0" w:color="999999"/></w:tcBorders></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>')
  + tableStyle('GridTable4-Accent1', 'Grid Table 4 Accent 1', 49, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>' + B(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], 'w:val="single" w:sz="4" w:space="0" w:color="8EAADB"') + '</w:tblBorders></w:tblPr>'
    + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders>' + B(['top', 'left', 'bottom', 'right'], 'w:val="single" w:sz="4" w:space="0" w:color="4472C4"') + '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="4472C4"/></w:tcBorders></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>')
  + tableStyle('GridTable5Dark-Accent1', 'Grid Table 5 Dark Accent 1', 50, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>' + B(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], 'w:val="single" w:sz="4" w:space="0" w:color="FFFFFF"') + '</w:tblBorders></w:tblPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr>'
    + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="B4C6E7"/></w:tcPr></w:tblStylePr>')
  + tableStyle('ListTable3-Accent2', 'List Table 3 Accent 2', 48, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>' + B(['top', 'left', 'bottom', 'right'], 'w:val="single" w:sz="4" w:space="0" w:color="ED7D31"') + '</w:tblBorders></w:tblPr>'
    + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="ED7D31"/></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="ED7D31"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="ED7D31"/></w:tcBorders></w:tcPr></w:tblStylePr>')
  + tableStyle('PlainTable1', 'Plain Table 1', 41, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>' + B(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], 'w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"') + '</w:tblBorders></w:tblPr>'
    + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:tcPr></w:tblStylePr>')
  + '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
  + '</w:styles>';

const SETTINGS = XML + '<w:settings ' + W + ' ' + R + ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">'
  + '<w:zoom w:percent="100"/><w:proofState w:spelling="clean" w:grammar="clean"/><w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>'
  + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>'
  + '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>'
  + '<w:themeFontLang w:val="pt-BR"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>'
  + '<w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/></w:settings>';

const FONTS = XML + '<w:fonts ' + W + ' ' + R + '>'
  + '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
  + '<w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
  + '<w:font w:name="Calibri Light"><w:panose1 w:val="020F0302020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
  + '<w:font w:name="Consolas"><w:panose1 w:val="020B0609020204030204"/><w:charset w:val="00"/><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>'
  + '<w:font w:name="Symbol"><w:panose1 w:val="05050102010706020507"/><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
  + '<w:font w:name="Courier New"><w:panose1 w:val="02070309020205020404"/><w:charset w:val="00"/><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>'
  + '<w:font w:name="Wingdings"><w:panose1 w:val="05000000000000000000"/><w:charset w:val="02"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>'
  + '</w:fonts>';

const clr = (name, val) => '<a:' + name + '><a:srgbClr val="' + val + '"/></a:' + name + '>';

const THEME = XML + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tema do Office"><a:themeElements>'
  + '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
  + clr('dk2', '44546A') + clr('lt2', 'E7E6E6') + clr('accent1', '4472C4') + clr('accent2', 'ED7D31') + clr('accent3', 'A5A5A5')
  + clr('accent4', 'FFC000') + clr('accent5', '5B9BD5') + clr('accent6', '70AD47') + clr('hlink', '0563C1') + clr('folHlink', '954F72') + '</a:clrScheme>'
  + '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
  + '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
  + '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
  + '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
  + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
  + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
  + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';

const NUMBERING = XML + '<w:numbering ' + W + ' ' + R + '></w:numbering>';

function coreXml(author) {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const esc = (s) => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;');
  return XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    + '<dc:title></dc:title><dc:creator>' + esc(author) + '</dc:creator><cp:lastModifiedBy>' + esc(author) + '</cp:lastModifiedBy><cp:revision>1</cp:revision>'
    + '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>';
}

const APP = XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>';

export async function blankDocxBytes(author) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/document.xml', DOCUMENT);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/settings.xml', SETTINGS);
  zip.file('word/numbering.xml', NUMBERING);
  zip.file('word/fontTable.xml', FONTS);
  zip.file('word/theme/theme1.xml', THEME);
  zip.file('docProps/core.xml', coreXml(author));
  zip.file('docProps/app.xml', APP);
  return zip.generateAsync({ type: 'uint8array' });
}
