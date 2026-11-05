import {getMovies,subscribeToMovies} from './data-management.js';

const LANGUAGES=['Bengali','Hindi','English'];
const EXCELJS_CDN='https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
let excelJsPromise=null;

function loadExcelJS(){
  if(window.ExcelJS)return Promise.resolve(window.ExcelJS);
  if(excelJsPromise)return excelJsPromise;
  excelJsPromise=new Promise((resolve,reject)=>{
    const script=document.createElement('script');
    script.src=EXCELJS_CDN;
    script.async=true;
    script.onload=()=>window.ExcelJS?resolve(window.ExcelJS):(excelJsPromise=null,reject(new Error('Excel workbook engine could not be initialized.')));
    script.onerror=()=>{excelJsPromise=null;reject(new Error('Could not load the Excel workbook engine. Check your internet connection.'));};
    document.head.appendChild(script);
  });
  return excelJsPromise;
}

const sortMovies=movies=>[...movies].sort((a,b)=>{
  const yearA=Number(a.year)||9999,yearB=Number(b.year)||9999;
  return yearA-yearB||a.name.localeCompare(b.name,undefined,{sensitivity:'base'});
});
const getPosterExtension=data=>{
  const match=String(data||'').match(/^data:image\\/(png|jpeg|jpg|gif|webp);base64,/i);
  if(!match)return null;
  const ext=match[1].toLowerCase();
  return ext==='jpeg'?'jpeg':ext;
};

export function initExcelInterface(){
  const fileInput=document.getElementById('excelFile');
  const fileName=document.getElementById('excelFileName');
  const fileHint=document.getElementById('excelFileHint');
  const exportButton=document.getElementById('exportUpdatedExcel');
  const completeStage=document.getElementById('completeStage');
  const confirmPreviewButton=document.getElementById('confirmExcelPreview');
  const status=document.getElementById('excelStatus');
  const sheetSelectWrap=document.getElementById('sheetSelectWrap');
  const sheetSelect=document.getElementById('excelSheetSelect');
  const showPreviewButton=document.getElementById('showExcelPreview');
  const previewWrap=document.getElementById('excelPreviewWrap');
  const preview=document.getElementById('excelPreview');
  if(!fileInput||!sheetSelectWrap||!sheetSelect||!showPreviewButton||!previewWrap||!preview)return;

  let selectedFile=null;
  let workbook=null;
  let currentPreviewSheet='';
  let uploadAttempt=0;
  const showMessage=message=>{if(status)status.textContent=message;};

  function getOrCreateLanguageSheet(language){
    let sheet=workbook.getWorksheet(language);
    if(!sheet)sheet=workbook.addWorksheet(language);
    return sheet;
  }

  function buildLanguageSheet(language){
    const sheet=getOrCreateLanguageSheet(language);
    sheet.getImages().forEach(image=>sheet.removeImage(image.imageId));
    [...(sheet.model.merges||[])].forEach(range=>sheet.unMergeCells(range));
    if(sheet.rowCount)sheet.spliceRows(1,sheet.rowCount);
    sheet.columns=[{width:9},...Array.from({length:Math.max(1,sortMovies(getMovies().filter(m=>m.language===language)).length)},()=>({width:27}))];
    sheet.views=[{state:'frozen',xSplit:1,ySplit:0,showGridLines:false}];
    sheet.properties.defaultRowHeight=20;
    const languageMovies=sortMovies(getMovies().filter(movie=>movie.language===language));
    const groups=new Map();
    languageMovies.forEach(movie=>{
      const letter=(movie.name.trim().charAt(0)||'#').toLocaleUpperCase();
      if(!groups.has(letter))groups.set(letter,[]);
      groups.get(letter).push(movie);
    });
    let row=1;
    for(const [letter,items] of [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0]))){
      sheet.mergeCells(row,1,row+3,1);
      const groupCell=sheet.getCell(row,1);
      groupCell.value=letter;
      groupCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}};
      groupCell.font={name:'Arial',size:24,color:{argb:'FF111111'}};
      groupCell.alignment={vertical:'middle',horizontal:'center'};
      groupCell.border={top:{style:'medium',color:{argb:'FF333333'}},bottom:{style:'medium',color:{argb:'FF333333'}},left:{style:'medium',color:{argb:'FF333333'}},right:{style:'medium',color:{argb:'FF333333'}}};
      sheet.getRow(row).height=25;
      sheet.getRow(row+1).height=25;
      sheet.getRow(row+2).height=22;
      sheet.getRow(row+3).height=190;
      items.forEach((movie,index)=>{
        const col=index+2;
        const year=sheet.getCell(row,col);
        year.value=movie.year;
        year.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF990000'}};
        year.font={name:'Arial',size:15,bold:true,color:{argb:'FFFFFFFF'}};
        year.alignment={horizontal:'center',vertical:'middle'};
        const title=sheet.getCell(row+1,col);
        title.value=movie.name;
        title.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}};
        title.font={name:'Arial',size:12,bold:true,color:{argb:'FF333333'}};
        title.alignment={horizontal:'center',vertical:'middle',wrapText:true};
        const indexCell=sheet.getCell(row+2,col);
        indexCell.value=index+1;
        indexCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF00DDEB'}};
        indexCell.font={name:'Arial',size:11,bold:true,color:{argb:'FF174F58'}};
        indexCell.alignment={horizontal:'center',vertical:'middle'};
        const posterCell=sheet.getCell(row+3,col);
        posterCell.value=movie.poster?'':'Poster not supplied';
        posterCell.alignment={horizontal:'center',vertical:'middle',wrapText:true};
        posterCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}};
        posterCell.border={left:{style:'thin',color:{argb:'FFB7B7B7'}},right:{style:'thin',color:{argb:'FFB7B7B7'}},bottom:{style:'thin',color:{argb:'FFB7B7B7'}}};
        const ext=getPosterExtension(movie.poster);
        if(ext){
          try{
            const base64=movie.poster.substring(movie.poster.indexOf(',')+1);
            const imageId=workbook.addImage({base64,extension:ext==='jpg'?'jpeg':ext});
            sheet.addImage(imageId,{tl:{col:col-0.96,row:row+2.96},br:{col:col-0.04,row:row+3.98},editAs:'oneCell'});
          }catch(error){console.warn('Could not embed poster for '+movie.name,error);}
        }
      });
      row+=4;
      sheet.getRow(row).height=8;
      row++;
    }
    if(!languageMovies.length){
      sheet.getCell('A1').value=language+' MOVIE COLLECTION';
      sheet.getCell('A1').font={bold:true,size:16,color:{argb:'FF174F7D'}};
      sheet.getCell('A2').value='Movies added in this language will appear here, grouped alphabetically and sorted by release year.';
    }
    sheet.getColumn(1).width=9;
    sheet.autoFilter=null;
    return sheet;
  }

  function renderLanguagePreview(language){
    const movies=sortMovies(getMovies().filter(movie=>movie.language===language));
    preview.replaceChildren();
    preview.classList.add('movie-excel-preview');
    const table=document.createElement('table');
    table.className='movie-excel-table';
    const body=document.createElement('tbody');
    const groups=new Map();
    movies.forEach(movie=>{
      const letter=(movie.name.trim().charAt(0)||'#').toLocaleUpperCase();
      if(!groups.has(letter))groups.set(letter,[]);
      groups.get(letter).push(movie);
    });
    if(!movies.length){
      const empty=document.createElement('p');
      empty.className='movie-excel-empty';
      empty.textContent='No '+language+' movies yet. Add a movie using the form; this worksheet will update automatically.';
      preview.appendChild(empty);
      return;
    }
    for(const [letter,items] of [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0]))){
      const groupRow=document.createElement('tr');
      const alphabet=document.createElement('th');
      alphabet.className='movie-excel-letter';
      alphabet.rowSpan=4;
      alphabet.textContent=letter;
      groupRow.appendChild(alphabet);
      items.forEach(movie=>{
        const cell=document.createElement('th');
        cell.className='movie-excel-year';
        cell.textContent=String(movie.year);
        groupRow.appendChild(cell);
      });
      body.appendChild(groupRow);
      const titleRow=document.createElement('tr');
      items.forEach(movie=>{const cell=document.createElement('th');cell.className='movie-excel-title';cell.textContent=movie.name;titleRow.appendChild(cell);});
      body.appendChild(titleRow);
      const indexRow=document.createElement('tr');
      items.forEach((movie,index)=>{const cell=document.createElement('td');cell.className='movie-excel-index';cell.textContent=String(index+1);indexRow.appendChild(cell);});
      body.appendChild(indexRow);
      const posterRow=document.createElement('tr');
      items.forEach(movie=>{
        const cell=document.createElement('td');cell.className='movie-excel-poster';
        if(movie.poster){const img=document.createElement('img');img.src=movie.poster;img.alt=movie.name+' poster';cell.appendChild(img);}
        else{const empty=document.createElement('span');empty.textContent='Poster not supplied';cell.appendChild(empty);}
        posterRow.appendChild(cell);
      });
      body.appendChild(posterRow);
      const gap=document.createElement('tr');const gapCell=document.createElement('td');gapCell.colSpan=items.length+1;gapCell.className='movie-excel-gap';gap.appendChild(gapCell);body.appendChild(gap);
    }
    table.appendChild(body);
    const wrap=document.createElement('div');wrap.className='movie-excel-scroll';wrap.appendChild(table);preview.appendChild(wrap);
    const note=document.createElement('p');note.className='excel-preview-note';note.textContent=language+' worksheet · '+movies.length+' movie(s) · Alphabetical groups A–Z · Release year oldest to newest.';preview.appendChild(note);
  }

  function renderGenericSheet(sheet){
    preview.replaceChildren();
    preview.classList.remove('movie-excel-preview');
    const table=document.createElement('table');table.className='excel-grid-table';
    const range=sheet.dimensions?sheet.dimensions.bottomRight:{row:1,column:1};
    const maxRow=Math.min(range.row||1,100),maxCol=Math.min(range.column||1,30);
    const head=document.createElement('thead'),headRow=document.createElement('tr');
    const corner=document.createElement('th');corner.textContent='';headRow.appendChild(corner);
    for(let c=1;c<=maxCol;c++){const th=document.createElement('th');th.textContent=String.fromCharCode(64+c);headRow.appendChild(th);}
    head.appendChild(headRow);table.appendChild(head);
    const body=document.createElement('tbody');
    for(let r=1;r<=maxRow;r++){const tr=document.createElement('tr');const th=document.createElement('th');th.className='excel-grid-row-heading';th.textContent=String(r);tr.appendChild(th);for(let c=1;c<=maxCol;c++){const td=document.createElement('td');const value=sheet.getRow(r).getCell(c).value;td.textContent=value&&typeof value==='object'?(value.text||value.result||''):String(value??'');tr.appendChild(td);}body.appendChild(tr);}
    table.appendChild(body);const wrap=document.createElement('div');wrap.className='movie-excel-scroll';wrap.appendChild(table);preview.appendChild(wrap);
  }

  function renderSelectedPreview(){
    if(!workbook||!currentPreviewSheet)return;
    if(LANGUAGES.includes(currentPreviewSheet))renderLanguagePreview(currentPreviewSheet);
    else renderGenericSheet(workbook.getWorksheet(currentPreviewSheet));
  }

  function syncLanguageSheets(){
    if(!workbook)return;
    LANGUAGES.forEach(language=>buildLanguageSheet(language));
    refreshSheetOptions();
    if(currentPreviewSheet)renderSelectedPreview();
  }

  function refreshSheetOptions(){
    if(!workbook)return;
    const previous=sheetSelect.value;
    sheetSelect.replaceChildren();
    workbook.worksheets.forEach((sheet,index)=>{
      const option=document.createElement('option');option.value=sheet.name;option.textContent=(index+1)+'. '+sheet.name;sheetSelect.appendChild(option);
    });
    if(workbook.getWorksheet(previous))sheetSelect.value=previous;
    else if(workbook.getWorksheet('Bengali'))sheetSelect.value='Bengali';
    else sheetSelect.value=workbook.worksheets[0]?.name||'';
  }

  fileInput.addEventListener('change',async()=>{
    const attempt=++uploadAttempt;
    selectedFile=fileInput.files&&fileInput.files[0]||null;
    workbook=null;currentPreviewSheet='';sheetSelectWrap.hidden=true;previewWrap.hidden=true;
    if(completeStage)completeStage.hidden=true;
    if(!selectedFile){fileName.textContent='Choose Excel workbook';fileHint.textContent='Click to browse · Upload required each time';showMessage('Upload a reference workbook to begin.');return;}
    if(!/\\.xlsx$/i.test(selectedFile.name)){
      selectedFile=null;fileInput.value='';fileName.textContent='Choose Excel workbook';fileHint.textContent='Please select an .xlsx file';
      showMessage('This formatted movie-collection workflow requires an .xlsx workbook. Please save legacy .xls files as .xlsx and upload again.');return;
    }
    fileName.textContent=selectedFile.name;fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · File selected successfully';
    showMessage('Excel detected. Loading workbook…');
    try{
      const ExcelJS=await loadExcelJS();if(attempt!==uploadAttempt)return;
      workbook=new ExcelJS.Workbook();
      await workbook.xlsx.load(await selectedFile.arrayBuffer());if(attempt!==uploadAttempt)return;
      syncLanguageSheets();
      sheetSelectWrap.hidden=false;
      showMessage('Workbook loaded. Language worksheets are mapped to Bengali, Hindi and English. Select a sheet to preview.');
      sheetSelectWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      if(attempt!==uploadAttempt)return;
      workbook=null;showMessage('Excel upload failed: '+(error.message||'Unable to read this workbook.'));
    }
  });

  showPreviewButton.addEventListener('click',()=>{
    if(!workbook)return;
    currentPreviewSheet=sheetSelect.value;
    if(!currentPreviewSheet)return;
    try{
      renderSelectedPreview();previewWrap.hidden=false;
      if(completeStage)completeStage.hidden=true;
      showMessage('Showing worksheet: '+currentPreviewSheet+'. Add movies to see this preview update live.');
      previewWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){showMessage('Could not display this worksheet: '+(error.message||'Unknown error'));}
  });

  subscribeToMovies(()=>{
    if(!workbook)return;
    try{
      syncLanguageSheets();
      if(currentPreviewSheet)showMessage('Movie collection updated. '+currentPreviewSheet+' preview is synchronized with the latest entries.');
    }catch(error){showMessage('Could not refresh the Excel preview: '+(error.message||'Unknown error'));}
  });

  confirmPreviewButton?.addEventListener('click',()=>{
    if(previewWrap.hidden)return;
    if(completeStage)completeStage.hidden=false;
    showMessage('Preview confirmed. Export the formatted language worksheets below.');
    completeStage?.scrollIntoView({behavior:'smooth',block:'nearest'});
  });

  exportButton?.addEventListener('click',async()=>{
    if(!workbook||!selectedFile)return;
    try{
      exportButton.disabled=true;showMessage('Preparing formatted Excel workbook with embedded movie posters…');
      syncLanguageSheets();
      const buffer=await workbook.xlsx.writeBuffer();
      const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url=URL.createObjectURL(blob);const link=document.createElement('a');
      link.href=url;link.download=selectedFile.name.replace(/\\.xlsx$/i,'')+'_updated.xlsx';
      document.body.appendChild(link);link.click();link.remove();URL.revokeObjectURL(url);
      showMessage('Export started. Language worksheets include the alphabetical/yearly layout, colour bands and embedded posters.');
    }catch(error){showMessage('Could not export workbook: '+(error.message||'Unknown error'));}
    finally{exportButton.disabled=false;}
  });
}
