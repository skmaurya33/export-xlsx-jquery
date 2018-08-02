/*
 * 
####################################################################################################
Built in styles
The following indexes are available form the styles that are predefined in the Editor XLSX style file. These indexes can be applied to any cells in the generated spreadsheet, altering their appearance.
0 - Normal text
1 - White text
2 - Bold
3 - Italic
4 - Underline
5 - Normal text, grey background
6 - White text, grey background
7 - Bold, grey background
8 - Italic, grey background
9 - Underline, grey background
10 - Normal text, red background
11 - White text, red background
12 - Bold, red background
13 - Italic, red background
14 - Underline, red background
15 - Normal text, green background
16 - White text, green background
17 - Bold, green background
18 - Italic, green background
19 - Underline, green background
20 - Normal text, blue background
21 - White text, blue background
22 - Bold, blue background
23 - Italic, blue background
24 - Underline, blue background
25 - Normal text, thin black border
26 - White text, thin black border
27 - Bold, thin black border
28 - Italic, thin black border
29 - Underline, thin black border
30 - Normal text, grey background, thin black border
31 - White text, grey background, thin black border
32 - Bold, grey background, thin black border
33 - Italic, grey background, thin black border
34 - Underline, grey background, thin black border
35 - Normal text, red background, thin black border
36 - White text, red background, thin black border
37 - Bold, red background, thin black border
38 - Italic, red background, thin black border
39 - Underline, red background, thin black border
40 - Normal text, green background, thin black border
41 - White text, green background, thin black border
42 - Bold, green background, thin black border
43 - Italic, green background, thin black border
44 - Underline, green background, thin black border
45 - Normal text, blue background, thin black border
46 - White text, blue background, thin black border
47 - Bold, blue background, thin black border
48 - Italic, blue background, thin black border
49 - Underline, blue background, thin black border
50 - Left aligned text (since 1.2.2)
51 - Centred text (since 1.2.2)
52 - Right aligned text (since 1.2.2)
53 - Justified text (since 1.2.2)
54 - Text rotated 90° (since 1.2.2)
55 - Wrapped text (since 1.2.2)
56 - Percentage integer value (automatically detected and used by buttons - since 1.2.3)
57 - Dollar currency values (automatically detected and used by buttons - since 1.2.3)
58 - Pound currency values (automatically detected and used by buttons - since 1.2.3)
59 - Euro currency values (automatically detected and used by buttons - since 1.2.3)
60 - Percentage with 1 decimal place (automatically detected and used by buttons - since 1.2.3)
61 - Negative numbers indicated by brackets (automatically detected and used by buttons - since 1.2.3)
62 - Negative numbers indicated by brackets - 2 decimal places (automatically detected and used by buttons - since 1.2.3)
63 - Numbers with thousand separators (automatically detected and used by buttons - since 1.2.3)
64 - Numbers with thousand separators - 2 decimal places (automatically detected and used by buttons - since 1.2.3)
65 - Numbers without thousand separators (automatically detected and used by buttons - since 1.2.4)
66 - Numbers without thousand separators - 2 decimal places (automatically detected and used by buttons - since 1.2.4)
####################################################################################################
 * 
 * https://datatables.net/reference/button/excelHtml5
 */
try {
    var _serialiser = new XMLSerializer();
    var _ieExcel;
} catch (t) {
}

// Excel - Pre-defined strings to build a basic XLSX file
var excelStrings = {
    "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>',

    "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
            '</Relationships>',

    "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="xml" ContentType="application/xml" />' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' +
            '<Default Extension="jpeg" ContentType="image/jpeg" />' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />' +
            '</Types>',

    "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="24816"/>' +
            '<workbookPr showInkAnnotation="0" autoCompressPictures="0"/>' +
            '<bookViews>' +
            '<workbookView xWindow="0" yWindow="0" windowWidth="25600" windowHeight="19020" tabRatio="500"/>' +
            '</bookViews>' +
            '<sheets>' +
            '<sheet name="" sheetId="1" r:id="rId1"/>' +
            '</sheets>' +
            '</workbook>',

    "xl/worksheets/sheet1.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' +
            '<sheetData/>' +
            '<mergeCells count="0"/>' +
            '</worksheet>',

    "xl/styles.xml":
            '<?xml version="1.0" encoding="UTF-8"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' +
            '<numFmts count="6">' +
            '<numFmt numFmtId="164" formatCode="#,##0.00_-\ [$$-45C]"/>' +
            '<numFmt numFmtId="165" formatCode="&quot;£&quot;#,##0.00"/>' +
            '<numFmt numFmtId="166" formatCode="[$€-2]\ #,##0.00"/>' +
            '<numFmt numFmtId="167" formatCode="0.0%"/>' +
            '<numFmt numFmtId="168" formatCode="#,##0;(#,##0)"/>' +
            '<numFmt numFmtId="169" formatCode="#,##0.00;(#,##0.00)"/>' +
            '</numFmts>' +
            '<fonts count="5" x14ac:knownFonts="1">' +
            '<font>' +
            '<sz val="11" />' +
            '<name val="Calibri" />' +
            '</font>' +
            '<font>' +
            '<sz val="11" />' +
            '<name val="Calibri" />' +
            '<color rgb="FFFFFFFF" />' +
            '</font>' +
            '<font>' +
            '<sz val="11" />' +
            '<name val="Calibri" />' +
            '<b />' +
            '</font>' +
            '<font>' +
            '<sz val="11" />' +
            '<name val="Calibri" />' +
            '<i />' +
            '</font>' +
            '<font>' +
            '<sz val="11" />' +
            '<name val="Calibri" />' +
            '<u />' +
            '</font>' +
            '</fonts>' +
            '<fills count="6">' +
            '<fill>' +
            '<patternFill patternType="none" />' +
            '</fill>' +
            '<fill>' + // Excel appears to use this as a dotted background regardless of values but
            '<patternFill patternType="none" />' + // to be valid to the schema, use a patternFill
            '</fill>' +
            '<fill>' +
            '<patternFill patternType="solid">' +
            '<fgColor rgb="FFD9D9D9" />' +
            '<bgColor indexed="64" />' +
            '</patternFill>' +
            '</fill>' +
            '<fill>' +
            '<patternFill patternType="solid">' +
            '<fgColor rgb="FFD99795" />' +
            '<bgColor indexed="64" />' +
            '</patternFill>' +
            '</fill>' +
            '<fill>' +
            '<patternFill patternType="solid">' +
            '<fgColor rgb="ffc6efce" />' +
            '<bgColor indexed="64" />' +
            '</patternFill>' +
            '</fill>' +
            '<fill>' +
            '<patternFill patternType="solid">' +
            '<fgColor rgb="ffc6cfef" />' +
            '<bgColor indexed="64" />' +
            '</patternFill>' +
            '</fill>' +
            '</fills>' +
            '<borders count="2">' +
            '<border>' +
            '<left />' +
            '<right />' +
            '<top />' +
            '<bottom />' +
            '<diagonal />' +
            '</border>' +
            '<border diagonalUp="false" diagonalDown="false">' +
            '<left style="thin">' +
            '<color auto="1" />' +
            '</left>' +
            '<right style="thin">' +
            '<color auto="1" />' +
            '</right>' +
            '<top style="thin">' +
            '<color auto="1" />' +
            '</top>' +
            '<bottom style="thin">' +
            '<color auto="1" />' +
            '</bottom>' +
            '<diagonal />' +
            '</border>' +
            '</borders>' +
            '<cellStyleXfs count="1">' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />' +
            '</cellStyleXfs>' +
            '<cellXfs count="67">' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="2" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="2" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="3" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="3" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="4" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="4" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="4" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="4" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="5" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="5" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="5" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="5" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="5" borderId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="0" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="5" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="5" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="5" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="5" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyAlignment="1">' +
            '<alignment horizontal="left"/>' +
            '</xf>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyAlignment="1">' +
            '<alignment horizontal="center"/>' +
            '</xf>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyAlignment="1">' +
            '<alignment horizontal="right"/>' +
            '</xf>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyAlignment="1">' +
            '<alignment horizontal="fill"/>' +
            '</xf>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyAlignment="1">' +
            '<alignment textRotation="90"/>' +
            '</xf>' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyAlignment="1">' +
            '<alignment wrapText="1"/>' +
            '</xf>' +
            '<xf numFmtId="9"   fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="167" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="168" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="169" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" xfId="0" applyNumberFormat="1"/>' +
            '</cellXfs>' +
            '<cellStyles count="1">' +
            '<cellStyle name="Normal" xfId="0" builtinId="0" />' +
            '</cellStyles>' +
            '<dxfs count="0" />' +
            '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4" />' +
            '</styleSheet>'
};
// Note we could use 3 `for` loops for the styles, but when gzipped there is
// virtually no difference in size, since the above can be easily compressed

// Pattern matching for special number formats. Perhaps this should be exposed
// via an API in future?
// Ref: section 3.8.30 - built in formatters in open spreadsheet
//   https://www.ecma-international.org/news/TC45_current_work/Office%20Open%20XML%20Part%204%20-%20Markup%20Language%20Reference.pdf
var _excelSpecials = [
    {match: /^\-?\d+\.\d%$/, style: 60, fmt: function (d) {
            return d / 100;
        }}, // Precent with d.p.
    {match: /^\-?\d+\.?\d*%$/, style: 56, fmt: function (d) {
            return d / 100;
        }}, // Percent
    {match: /^\-?\$[\d,]+.?\d*$/, style: 57}, // Dollars
    {match: /^\-?£[\d,]+.?\d*$/, style: 58}, // Pounds
    {match: /^\-?€[\d,]+.?\d*$/, style: 59}, // Euros
    {match: /^\-?\d+$/, style: 65}, // Numbers without thousand separators
    {match: /^\-?\d+\.\d{2}$/, style: 66}, // Numbers 2 d.p. without thousands separators
    {match: /^\([\d,]+\)$/, style: 61, fmt: function (d) {
            return -1 * d.replace(/[\(\)]/g, '');
        }}, // Negative numbers indicated by brackets
    {match: /^\([\d,]+\.\d{2}\)$/, style: 62, fmt: function (d) {
            return -1 * d.replace(/[\(\)]/g, '');
        }}, // Negative numbers indicated by brackets - 2d.p.
    {match: /^\-?[\d,]+$/, style: 63}, // Numbers with thousand separators
    {match: /^\-?[\d,]+\.\d{2}$/, style: 64}  // Numbers with 2 d.p. and thousands separators
];

var _sheetname = function (config)
{
    var sheetName = 'Sheet1';

    if (config.sheetName) {
        sheetName = config.sheetName.replace(/[\[\]\*\/\\\?\:]/g, '');
    }

    return sheetName;
};

/**
 * Create an XML node and add any children, attributes, etc without needing to
 * be verbose in the DOM.
 *
 * @param  {object} doc      XML document
 * @param  {string} nodeName Node name
 * @param  {object} opts     Options - can be `attr` (attributes), `children`
 *   (child nodes) and `text` (text content)
 * @return {node}            Created node
 */
function _createNode(doc, nodeName, opts) {
    var tempNode = doc.createElement(nodeName);

    if (opts) {
        if (opts.attr) {
            $(tempNode).attr(opts.attr);
        }

        if (opts.children) {
            $.each(opts.children, function (key, value) {
                tempNode.appendChild(value);
            });
        }

        if (opts.text !== null && opts.text !== undefined) {
            tempNode.appendChild(doc.createTextNode(opts.text));
        }
    }

    return tempNode;
}
/**
 * Convert from numeric position to letter for column names in Excel
 * @param  {int} n Column number
 * @return {string} Column letter(s) name
 */
function createCellPos(n) {
    var ordA = 'A'.charCodeAt(0);
    var ordZ = 'Z'.charCodeAt(0);
    var len = ordZ - ordA + 1;
    var s = "";

    while (n >= 0) {
        s = String.fromCharCode(n % len + ordA) + s;
        n = Math.floor(n / len) - 1;
    }

    return s;
}
/**
 * Get the width for an Excel column based on the contents of that column
 * @param  {object} data Data for export
 * @param  {int}    col  Column index
 * @return {int}         Column width
 */
function _excelColWidth(data, col) {
    var max = data.header[col].length;
    var len, lineSplit, str;

    if (data.footer && data.footer[col].length > max) {
        max = data.footer[col].length;
    }

    for (var i = 0, ien = data.body.length; i < ien; i++) {
        var point = data.body[i][col];
        str = point !== null && point !== undefined ?
                point.toString() :
                '';

        // If there is a newline character, workout the width of the column
        // based on the longest line in the string
        if (str.indexOf('\n') !== -1) {
            lineSplit = str.split('\n');
            lineSplit.sort(function (a, b) {
                return b.length - a.length;
            });

            len = lineSplit[0].length;
        } else {
            len = str.length;
        }

        if (len > max) {
            max = len;
        }

        // Max width rather than having potentially massive column widths
        if (max > 40) {
            return 52; // 40 * 1.3
        }
    }

    max *= 1.3;

    // And a min width
    return max > 6 ? max : 6;
}

// Allow the constructor to pass in JSZip and PDFMake from external requires.
// Otherwise, use globally defined variables, if they are available.
function _jsZip() {
    return jszip || window.JSZip;
}

/**
 * Recursively add XML files from an object's structure to a ZIP file. This
 * allows the XSLX file to be easily defined with an object's structure matching
 * the files structure.
 *
 * @param {JSZip} zip ZIP package
 * @param {object} obj Object to add (recursive)
 */
function _addToZip(zip, obj) {
    if (_ieExcel === undefined) {
        // Detect if we are dealing with IE's _awful_ serialiser by seeing if it
        // drop attributes
        _ieExcel = _serialiser
                .serializeToString(
                        $.parseXML(excelStrings['xl/worksheets/sheet1.xml'])
                        )
                .indexOf('xmlns:r') === -1;
    }

    $.each(obj, function (name, val) {
        if ($.isPlainObject(val)) {
            var newDir = zip.folder(name);
            _addToZip(newDir, val);
        } else {
            if (_ieExcel) {
                // IE's XML serialiser will drop some name space attributes from
                // from the root node, so we need to save them. Do this by
                // replacing the namespace nodes with a regular attribute that
                // we convert back when serialised. Edge does not have this
                // issue
                var worksheet = val.childNodes[0];
                var i, ien;
                var attrs = [];

                for (i = worksheet.attributes.length - 1; i >= 0; i--) {
                    var attrName = worksheet.attributes[i].nodeName;
                    var attrValue = worksheet.attributes[i].nodeValue;

                    if (attrName.indexOf(':') !== -1) {
                        attrs.push({name: attrName, value: attrValue});

                        worksheet.removeAttribute(attrName);
                    }
                }

                for (i = 0, ien = attrs.length; i < ien; i++) {
                    var attr = val.createAttribute(attrs[i].name.replace(':', '_dt_b_namespace_token_'));
                    attr.value = attrs[i].value;
                    worksheet.setAttributeNode(attr);
                }
            }

            var str = _serialiser.serializeToString(val);

            // Fix IE's XML
            if (_ieExcel) {
                // IE doesn't include the XML declaration
                if (str.indexOf('<?xml') === -1) {
                    str = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + str;
                }

                // Return namespace attributes to being as such
                str = str.replace(/_dt_b_namespace_token_/g, ':');
            }

            // Safari, IE and Edge will put empty name space attributes onto
            // various elements making them useless. This strips them out
            str = str.replace(/<([^<>]*?) xmlns=""([^<>]*?)>/g, '<$1 $2>');

            zip.file(name, str);
        }
    });
}
/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * FileSaver.js dependency
 */

/*jslint bitwise: true, indent: 4, laxbreak: true, laxcomma: true, smarttabs: true, plusplus: true */

var _saveAs = (function (view) {
    "use strict";
    // IE <10 is explicitly unsupported
    if (typeof view === "undefined" || typeof navigator !== "undefined" && /MSIE [1-9]\./.test(navigator.userAgent)) {
        return;
    }
    var
            doc = view.document
            // only get URL when necessary in case Blob.js hasn't overridden it yet
            , get_URL = function () {
                return view.URL || view.webkitURL || view;
            }
    , save_link = doc.createElementNS("http://www.w3.org/1999/xhtml", "a")
            , can_use_save_link = "download" in save_link
            , click = function (node) {
                var event = new MouseEvent("click");
                node.dispatchEvent(event);
            }
    , is_safari = /constructor/i.test(view.HTMLElement) || view.safari
            , is_chrome_ios = /CriOS\/[\d]+/.test(navigator.userAgent)
            , throw_outside = function (ex) {
                (view.setImmediate || view.setTimeout)(function () {
                    throw ex;
                }, 0);
            }
    , force_saveable_type = "application/octet-stream"
            // the Blob API is fundamentally broken as there is no "downloadfinished" event to subscribe to
            , arbitrary_revoke_timeout = 1000 * 40 // in ms
            , revoke = function (file) {
                var revoker = function () {
                    if (typeof file === "string") { // file is an object URL
                        get_URL().revokeObjectURL(file);
                    } else { // file is a File
                        file.remove();
                    }
                };
                setTimeout(revoker, arbitrary_revoke_timeout);
            }
    , dispatch = function (filesaver, event_types, event) {
        event_types = [].concat(event_types);
        var i = event_types.length;
        while (i--) {
            var listener = filesaver["on" + event_types[i]];
            if (typeof listener === "function") {
                try {
                    listener.call(filesaver, event || filesaver);
                } catch (ex) {
                    throw_outside(ex);
                }
            }
        }
    }
    , auto_bom = function (blob) {
        // prepend BOM for UTF-8 XML and text/* types (including HTML)
        // note: your browser will automatically convert UTF-16 U+FEFF to EF BB BF
        if (/^\s*(?:text\/\S*|application\/xml|\S*\/\S*\+xml)\s*;.*charset\s*=\s*utf-8/i.test(blob.type)) {
            return new Blob([String.fromCharCode(0xFEFF), blob], {type: blob.type});
        }
        return blob;
    }
    , FileSaver = function (blob, name, no_auto_bom) {
        if (!no_auto_bom) {
            blob = auto_bom(blob);
        }
        // First try a.download, then web filesystem, then object URLs
        var
                filesaver = this
                , type = blob.type
                , force = type === force_saveable_type
                , object_url
                , dispatch_all = function () {
                    dispatch(filesaver, "writestart progress write writeend".split(" "));
                }
        // on any filesys errors revert to saving with object URLs
        , fs_error = function () {
            if ((is_chrome_ios || (force && is_safari)) && view.FileReader) {
                // Safari doesn't allow downloading of blob urls
                var reader = new FileReader();
                reader.onloadend = function () {
                    var url = is_chrome_ios ? reader.result : reader.result.replace(/^data:[^;]*;/, 'data:attachment/file;');
                    var popup = view.open(url, '_blank');
                    if (!popup)
                        view.location.href = url;
                    url = undefined; // release reference before dispatching
                    filesaver.readyState = filesaver.DONE;
                    dispatch_all();
                };
                reader.readAsDataURL(blob);
                filesaver.readyState = filesaver.INIT;
                return;
            }
            // don't create more object URLs than needed
            if (!object_url) {
                object_url = get_URL().createObjectURL(blob);
            }
            if (force) {
                view.location.href = object_url;
            } else {
                var opened = view.open(object_url, "_blank");
                if (!opened) {
                    // Apple does not allow window.open, see https://developer.apple.com/library/safari/documentation/Tools/Conceptual/SafariExtensionGuide/WorkingwithWindowsandTabs/WorkingwithWindowsandTabs.html
                    view.location.href = object_url;
                }
            }
            filesaver.readyState = filesaver.DONE;
            dispatch_all();
            revoke(object_url);
        }
        ;
        filesaver.readyState = filesaver.INIT;

        if (can_use_save_link) {
            object_url = get_URL().createObjectURL(blob);
            setTimeout(function () {
                save_link.href = object_url;
                save_link.download = name;
                click(save_link);
                dispatch_all();
                revoke(object_url);
                filesaver.readyState = filesaver.DONE;
            });
            return;
        }

        fs_error();
    }
    , FS_proto = FileSaver.prototype
            , saveAs = function (blob, name, no_auto_bom) {
                return new FileSaver(blob, name || blob.name || "download", no_auto_bom);
            }
    ;
    // IE 10+ (native saveAs)
    if (typeof navigator !== "undefined" && navigator.msSaveOrOpenBlob) {
        return function (blob, name, no_auto_bom) {
            name = name || blob.name || "download";

            if (!no_auto_bom) {
                blob = auto_bom(blob);
            }
            return navigator.msSaveOrOpenBlob(blob, name);
        };
    }

    FS_proto.abort = function () {};
    FS_proto.readyState = FS_proto.INIT = 0;
    FS_proto.WRITING = 1;
    FS_proto.DONE = 2;

    FS_proto.error =
            FS_proto.onwritestart =
            FS_proto.onprogress =
            FS_proto.onwrite =
            FS_proto.onabort =
            FS_proto.onerror =
            FS_proto.onwriteend =
            null;

    return saveAs;
}(
        typeof self !== "undefined" && self
        || typeof window !== "undefined" && window
        || this.content
        ));


function skm_gs(table, key) {
    var style = (key == 'header' || key == 'footer') ? 7 : 0;
     if (table.hasOwnProperty('styles') && table.styles !=null && table.styles.hasOwnProperty(key)) {
        style = table.styles[key];
    }
    return style;
}
function customizeCallBack(xlsx, config) {

    //console.log(xlsx);
    //console.log(aa);
    //console.log(bb);
    var sheet = xlsx.xl.worksheets['sheet1.xml'];
    console.log(sheet);
    //console.log($(sheet));

    $('row', sheet).each(function (index) {
        //console.log(index);
        if (index % 2 == 0) {
            $(this).find('c').attr('s', '12');
        }
    });


    // Loop over the cells in column `C`
    /*$('row c[r^="C"]', sheet).each(function () {
     // Get the value
     if ($('is t', this).text() == 'New York') {
     $(this).attr('s', '12');
     //$(this).attr('s', '2');
     }
     });*/
}



function generateXLSX(config2, tableData) {

    var config1 = {
        // className: "buttons-excel buttons-html5",
        filename: "report",
        sheetName: "report",
        extension: ".xlsx",
        //exportOptions: {},
        header: true,
        footer: false,
        // title: "*",
        //messageTop: "*",
        // messageBottom: "*",
        createEmptyCells: true,
        createEmptyRow: true,
        //namespace: ".dt-button-0",
        customize: false,
        customizeCallBack: customizeCallBack,
    };

    var config = Object.assign(config1, config2);

    var rowPos = 0;
    var getXml = function (type) {
        var str = excelStrings[ type ];

        //str = str.replace( /xmlns:/g, 'xmlns_' ).replace( /mc:/g, 'mc_' );

        return $.parseXML(str);
    };
    var rels = getXml('xl/worksheets/sheet1.xml');
    var relsGet = rels.getElementsByTagName("sheetData")[0];

    var xlsx = {
        _rels: {
            ".rels": getXml('_rels/.rels')
        },
        xl: {
            _rels: {
                "workbook.xml.rels": getXml('xl/_rels/workbook.xml.rels')
            },
            "workbook.xml": getXml('xl/workbook.xml'),
            "styles.xml": getXml('xl/styles.xml'),
            "worksheets": {
                "sheet1.xml": rels
            }

        },
        "[Content_Types].xml": getXml('[Content_Types].xml')
    };

    var currentRow, rowNode;
    var addRow = function (row) {
        currentRow = rowPos + 1;
        rowNode = _createNode(rels, "row", {attr: {r: currentRow}});

        for (var i = 0, ien = row.length; i < ien; i++) {
            // Concat both the Cell Columns as a letter and the Row of the cell.
            var cellId = createCellPos(i) + '' + currentRow;
            var cell = null;

            // For null, undefined of blank cell, continue so it doesn't create the _createNode
            if (row[i] === null || row[i] === undefined || row[i] === '') {
                if (config.createEmptyCells === true) {
                    row[i] = '';
                } else {
                    continue;
                }
            }

            var originalContent = row[i];
            row[i] = $.trim(row[i]);

            // Special number formatting options
            for (var j = 0, jen = _excelSpecials.length; j < jen; j++) {
                var special = _excelSpecials[j];

                // TODO Need to provide the ability for the specials to say
                // if they are returning a string, since at the moment it is
                // assumed to be a number
                if (row[i].match && !row[i].match(/^0\d+/) && row[i].match(special.match)) {
                    var val = row[i].replace(/[^\d\.\-]/g, '');

                    if (special.fmt) {
                        val = special.fmt(val);
                    }

                    cell = _createNode(rels, 'c', {
                        attr: {
                            r: cellId,
                            s: special.style
                        },
                        children: [
                            _createNode(rels, 'v', {text: val})
                        ]
                    });

                    break;
                }
            }

            if (!cell) {
                if (typeof row[i] === 'number' || (
                        row[i].match &&
                        row[i].match(/^-?\d+(\.\d+)?$/) &&
                        !row[i].match(/^0\d+/))
                        ) {
                    // Detect numbers - don't match numbers with leading zeros
                    // or a negative anywhere but the start
                    cell = _createNode(rels, 'c', {
                        attr: {
                            t: 'n',
                            r: cellId
                        },
                        children: [
                            _createNode(rels, 'v', {text: row[i]})
                        ]
                    });
                } else {
                    // String output - replace non standard characters for text output
                    var text = !originalContent.replace ?
                            originalContent :
                            originalContent.replace(/[\x00-\x09\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, '');

                    cell = _createNode(rels, 'c', {
                        attr: {
                            t: 'inlineStr',
                            r: cellId
                        },
                        children: {
                            row: _createNode(rels, 'is', {
                                children: {
                                    row: _createNode(rels, 't', {
                                        text: text,
                                        attr: {
                                            'xml:space': 'preserve'
                                        }
                                    })
                                }
                            })
                        }
                    });
                }
            }
            rowNode.appendChild(cell);
        }

        relsGet.appendChild(rowNode);
        rowPos++;
    };

    $('sheets sheet', xlsx.xl['workbook.xml']).attr('name', _sheetname(config));



    var mergeCells = function (row, colspan) {
        var mergeCells = $('mergeCells', rels);

        mergeCells[0].appendChild(_createNode(rels, 'mergeCell', {
            attr: {
                ref: 'A' + row + ':' + createCellPos(colspan) + row
            }
        }));
        mergeCells.attr('count', parseFloat(mergeCells.attr('count')) + 1);
        $('row:eq(' + (row - 1) + ') c', rels).attr('s', '51'); // centre
    };

    $.each(tableData, function (index, table) {

        if (table.title) {
            $.each(table.title, function (ind, text) {
                addRow([text], rowPos);
                mergeCells(rowPos, table.header.length - 1);
                $('row:last c', rels).attr('s', skm_gs(table, 'title'));
            });
        }

        if (table.messageTop) {
            addRow([table.messageTop], rowPos);
            mergeCells(rowPos, table.header.length - 1);
            $('row:last c', rels).attr('s', skm_gs(table, 'messageTop'));
        }

        // Table itself
        if (config.header) {
            addRow(table.header, rowPos);
            //$('row:last c', rels).attr('s', '7'); // bold
            $('row:last c', rels).attr('s', skm_gs(table, 'header'));
        }

        for (var n = 0, ie = table.body.length; n < ie; n++) {
            addRow(table.body[n], rowPos);
        }

        if (config.footer && table.footer) {
            addRow(table.footer, rowPos);
            //$('row:last c', rels).attr('s', '7'); // bold
            $('row:last c', rels).attr('s', skm_gs(table, 'footer'));
        }

        // Below the table
        if (table.messageBottom) {
            addRow([table.messageBottom], rowPos);
            mergeCells(rowPos, table.header.length - 1);
            $('row:last c', rels).attr('s', skm_gs(table, 'messageBottom'));
        }

        if (config.createEmptyRow) {
            addRow([""], rowPos);
            mergeCells(rowPos, table.header.length - 1);
        }
        // Set column widths
        var cols = _createNode(rels, 'cols');
        $('worksheet', rels).prepend(cols);

        for (var i = 0, ien = table.header.length; i < ien; i++) {
            cols.appendChild(_createNode(rels, 'col', {
                attr: {
                    min: i + 1,
                    max: i + 1,
                    width: _excelColWidth(table, i),
                    customWidth: 1
                }
            }));
        }
    });


    // Let the developer customise the document if they want to
    if (config.customize) {
        config.customizeCallBack(xlsx, config);
    }

// Excel doesn't like an empty mergeCells tag
    if ($('mergeCells', rels).children().length === 0) {
        $('mergeCells', rels).remove();
    }

//var jszip = _jsZip();
    var zip = new JSZip();
    var zipConfig = {
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };

    _addToZip(zip, xlsx);

    if (zip.generateAsync) {
        // JSZip 3+
        zip
                .generateAsync(zipConfig)
                .then(function (blob) {
                    _saveAs(blob, config.filename);

                });
    } else {
        // JSZip 2.5
        _saveAs(
                zip.generate(zipConfig),
                config.filename
                );

    }
}
