// Code.gs
// job    : creates google sheets add-on for exporting markdown table text file
// git:   https://github.com/motetpaper/marta-duo
// lic:   MIT https://opensource.org/license/mit
// ver:   v0.3
//

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Download as Markdown text file')
    .addItem('Download the active range ...', 'exportMarkdownTable')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function exportMarkdownTable() {
  SpreadsheetApp.getUi()
    .showSidebar(HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setTitle('Exporting Markdown table ...'));
}

function mkmarta() {

  const obj = {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getActiveSheet();
  const rng = sht.getDataRange();
  const rows = rng.getNumRows();
  const cols = rng.getNumColumns();
 
  // sheet data area
  obj.name = sht.getSheetName();
  obj.notation = rng.getA1Notation();

  obj.rows = rows;
  obj.cols = cols;


  obj.lastcol = rng.getLastColumn();
  obj.lastrow = rng.getLastRow();

  obj.fcols = sht.getFrozenColumns();
  obj.frows = sht.getFrozenRows(); 
  obj.aligns = rng.getHorizontalAlignments();

  // hidden rows

  obj.hiderows = [];
  let isRowHidden = null;
  for(let j = 1; j < rows+1; j++) {
    isRowHidden = sht.isRowHiddenByFilter(j) || sht.isRowHiddenByUser(j);
    obj.hiderows.push([`ROW-${j}`, isRowHidden]);
  }
  
  // hidden columns

  obj.hidecols = [];
  let isColHidden = null;
  for(let i = 1; i < cols+1; i++) {
    isColHidden = sht.isColumnHiddenByUser(i);
    obj.hidecols.push([`COL-${i}`,isColHidden]);
  }  

  // cell values area
  obj.dvals = rng.getDisplayValues();
  obj.vals = rng.getValues();
  
  // fonts area
  obj.fweights = rng.getFontWeights();
  obj.fstyles = rng.getFontStyles();
  obj.fontlines = rng.getFontLines();    
  obj.ffams = rng.getFontFamilies(); 
  obj.rtvals = rng.getRichTextValues();
  obj.tstyles = rng.getTextStyles();
  
  // detects if text style is underline
  obj.ts_underline = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isUnderline();
    });
  });

  // detects if text style is bold
  obj.ts_bold = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isBold();
    });
  });

  // detects if text style is strikethrough (line-through)
  obj.ts_strikethrough = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isStrikethrough();
    });
  });


  // detects if text style is italic
  obj.ts_italic = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isItalic();
    });
  });    

  // validations area
  obj.valids = rng.getDataValidations();

  // detects if validation is checkbox
  obj.ischeckbox = obj.valids.map((a,j)=>{
    return a.map((b,i)=>{
      return (!!b) ? b.getCriteriaType().toJSON() === 'CHECKBOX': false;
    });
  });

  // formulas area
  obj.formulas = rng.getFormulas();
  obj.formulasrc = rng.getFormulasR1C1();

  // footnotes area
  obj.notes = rng.getNotes();
  obj.hasnotes = obj.notes.map((a,j)=>{
    return a.map((b,i)=>{
      return !!b.length;
    });
  });

  return obj;
}
