import {getMovies,subscribeToMovies} from './data-management.js';

const LANGUAGE_SHEETS={Bengali:'Bengali Movies',Hindi:'Hindi Movies',English:'English Movies'};
const LANGUAGES=Object.keys(LANGUAGE_SHEETS);
const WORKBOOK_DB='chayan-watchlist-excel-state';
const WORKBOOK_STORE='excel-workspace';
function openWorkbookDB(){return new Promise((resolve,reject)=>{const request=indexedDB.open(WORKBOOK_DB,1);request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains(WORKBOOK_STORE))request.result.createObjectStore(WORKBOOK_STORE);};request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);});}
async function saveWorkbookState(state){const db=await openWorkbookDB();await new Promise((resolve,reject)=>{const tx=db.transaction(WORKBOOK_STORE,'readwrite');tx.objectStore(WORKBOOK_STORE).put(state,'last-workbook');tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);});db.close();}
async function readWorkbookState(){const db=await openWorkbookDB();const value=await new Promise((resolve,reject)=>{const request=db.transaction(WORKBOOK_STORE,'readonly').objectStore(WORKBOOK_STORE).get('last-workbook');request.onsuccess=()=>resolve(request.result||null);request.onerror=()=>reject(request.error);});db.close();return value;}
async function clearWorkbookState(){const db=await openWorkbookDB();await new Promise((resolve,reject)=>{const tx=db.transaction(WORKBOOK_STORE,'readwrite');tx.objectStore(WORKBOOK_STORE).delete('last-workbook');tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);});db.close();}

const getLanguageForSheet=name=>LANGUAGES.find(language=>[LANGUAGE_SHEETS[language],language].includes(name))||null;
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
const posterCropCache=new Map();
function cropPosterToCell(data){
  if(posterCropCache.has(data))return posterCropCache.get(data);
  const task=new Promise((resolve,reject)=>{
    const image=new Image();
    image.onload=()=>{
      try{
        const width=380,height=550;
        const canvas=document.createElement('canvas');
        canvas.width=width;canvas.height=height;
        const context=canvas.getContext('2d');
        const scale=Math.max(width/image.naturalWidth,height/image.naturalHeight);
        const sourceWidth=width/scale,sourceHeight=height/scale;
        const sourceX=(image.naturalWidth-sourceWidth)/2;
        const sourceY=(image.naturalHeight-sourceHeight)/2;
        context.drawImage(image,sourceX,sourceY,sourceWidth,sourceHeight,0,0,width,height);
        resolve(canvas.toDataURL('image/png'));
      }catch(error){reject(error);}
    };
    image.onerror=()=>reject(new Error('Could not prepare a movie poster for the Excel cell.'));
    image.src=data;
  });
  posterCropCache.set(data,task);
  task.catch(()=>posterCropCache.delete(data));
  return task;
}

const getPosterExtension=data=>{
  const match=String(data||'').match(/^data:image\/(png|jpeg|jpg|gif|webp);base64,/i);
  if(!match)return null;
  const ext=match[1].toLowerCase();
  return ext==='jpeg'?'jpeg':ext;
};

export function initExcelInterface(){
  const fileInput=document.getElementById('excelFile');
  const resetButton=document.getElementById('resetAndUploadNewFile');
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
  let resetInProgress=false;
  const resetLanding=new URLSearchParams(window.location.search).has('reset');
  const showMessage=message=>{if(status)status.textContent=message;};
  async function persistCurrentWorkbook(){
    if(resetInProgress||localStorage.getItem('chayan-watchlist-reset-in-progress')==='1'||!workbook||!selectedFile)return;
    const buffer=await workbook.xlsx.writeBuffer();
    if(resetInProgress||localStorage.getItem('chayan-watchlist-reset-in-progress')==='1')return;
    await saveWorkbookState({buffer,fileName:selectedFile.name,sheetName:currentPreviewSheet,previewVisible:!!(previewWrap&&!previewWrap.hidden)});
  }
  // A reset invalidates in-flight uploads/restores before clearing every app-owned cache.
  if(resetButton&&!resetButton.dataset.resetHandlerBound){
    resetButton.dataset.resetHandlerBound='true';
    resetButton.addEventListener('click',async event=>{
      event.preventDefault();
      event.stopPropagation();
      if(resetInProgress)return;
      resetInProgress=true;
      uploadAttempt++;
      workbook=null;
      selectedFile=null;
      currentPreviewSheet='';
      resetButton.disabled=true;
      resetButton.textContent='Resetting…';
      if(preview)preview.replaceChildren();
      previewWrap.hidden=true;
      sheetSelectWrap.hidden=true;
      if(completeStage)completeStage.hidden=true;
      showMessage('Clearing previous workbook and saved application data…');
      try{
        localStorage.setItem('chayan-watchlist-reset-in-progress','1');
        // Remove every IndexedDB database owned by this application, including legacy versions.
        if(typeof indexedDB.databases==='function'){
          const databases=await indexedDB.databases();
          const owned=databases.map(db=>db.name).filter(name=>name&&/chayan-watchlist/i.test(name));
          await Promise.all(owned.map(name=>new Promise((resolve,reject)=>{
            const request=indexedDB.deleteDatabase(name);
            request.onsuccess=()=>resolve();
            request.onerror=()=>reject(request.error||new Error('Could not delete '+name));
            request.onblocked=()=>reject(new Error('A workbook database is still open. Close other app tabs and retry.'));
          })));
        }else{
          await clearWorkbookState();
        }
        // Clear all persistent and temporary app state, retaining only the reset lock until reload.
        for(const key of Object.keys(localStorage))localStorage.removeItem(key);
        sessionStorage.clear();
        localStorage.setItem('chayan-watchlist-reset-in-progress','1');
        window.location.replace(window.location.pathname+'?reset='+Date.now());
      }catch(error){
        localStorage.removeItem('chayan-watchlist-reset-in-progress');
        resetInProgress=false;
        resetButton.disabled=false;
        resetButton.textContent='↻ Reset and Ready for New Upload';
        showMessage('Reset failed: '+(error.message||'Unable to clear saved data. Please close other app tabs and retry.'));
        console.error('Full application reset failed:',error);
      }
    },true);
  }

  function getOrCreateLanguageSheet(language){
    const targetName=LANGUAGE_SHEETS[language];
    const existing=workbook.getWorksheet(targetName);
    const legacy=workbook.getWorksheet(language);
    if(existing)workbook.removeWorksheet(existing.id);
    if(legacy&&legacy.id!==existing?.id)workbook.removeWorksheet(legacy.id);
    return workbook.addWorksheet(targetName);
  }

  function removeDuplicateLanguageSheets(){
    LANGUAGES.forEach(language=>{
      const preferred=workbook.getWorksheet(LANGUAGE_SHEETS[language]);
      const duplicate=workbook.getWorksheet(language);
      if(preferred&&duplicate&&preferred.id!==duplicate.id)workbook.removeWorksheet(duplicate.id);
    });
  }

  async function buildLanguageSheet(language){
    const sheet=getOrCreateLanguageSheet(language);
    sheet.getImages().forEach(image=>sheet.removeImage(image.imageId));
    [...(sheet.model.merges||[])].forEach(range=>sheet.unMergeCells(range));
    if(sheet.rowCount)sheet.spliceRows(1,sheet.rowCount);
    sheet.columns=[{width:15.5},...Array.from({length:Math.max(1,sortMovies(getMovies().filter(m=>m.language===language)).length)},()=>({width:27}))];
    sheet.views=[{state:'frozen',xSplit:1,ySplit:0,showGridLines:false}];
    sheet.properties.defaultRowHeight=20;
    sheet.properties.defaultColWidth=27;
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
      const posterCount=items.filter(movie=>Boolean(movie.poster)).length;
      groupCell.value={richText:[
        {text:letter,font:{name:'Aptos Display',size:50,bold:true,color:{argb:'FF172B4D'}}},
        {text:'\n('+posterCount+')',font:{name:'Aptos',size:13,bold:true,italic:true,color:{argb:'FF147D75'}}}
      ]};
      groupCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFDCE5F0'}};
      groupCell.alignment={vertical:'middle',horizontal:'center',wrapText:true};
      groupCell.border={top:{style:'medium',color:{argb:'FFB8C6D8'}},bottom:{style:'medium',color:{argb:'FFB8C6D8'}},left:{style:'medium',color:{argb:'FFB8C6D8'}},right:{style:'medium',color:{argb:'FFB8C6D8'}}};
      sheet.getRow(row).height=19;
      sheet.getRow(row+1).height=16;
      sheet.getRow(row+2).height=15;
      sheet.getRow(row+3).height=190;
      for(const [index,movie] of items.entries()){
        const col=index+2;
        const year=sheet.getCell(row,col);
        year.value=movie.year;
        year.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF172B4D'}};
        year.font={name:'Aptos Display',size:15,bold:true,color:{argb:'FFFFFFFF'}};
        year.alignment={horizontal:'center',vertical:'middle'};
        const title=sheet.getCell(row+1,col);
        title.value=movie.name;
        title.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8D7A5'}};
        title.font={name:'Aptos',size:12,bold:true,color:{argb:'FF202A35'}};
        title.alignment={horizontal:'center',vertical:'middle',wrapText:false,shrinkToFit:true};
        const indexCell=sheet.getCell(row+2,col);
        indexCell.value=index+1;
        indexCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF9CCBC5'}};
        indexCell.font={name:'Aptos',size:11,bold:true,color:{argb:'FF202A35'}};
        indexCell.alignment={horizontal:'center',vertical:'middle'};
        const posterCell=sheet.getCell(row+3,col);
        posterCell.value=movie.poster?'':'Poster not supplied';
        posterCell.alignment={horizontal:'center',vertical:'middle',wrapText:true};
        posterCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}};
        posterCell.border={left:{style:'thin',color:{argb:'FFB8C6D8'}},right:{style:'thin',color:{argb:'FFB8C6D8'}},bottom:{style:'thin',color:{argb:'FFB8C6D8'}}};
        const ext=getPosterExtension(movie.poster);
        if(ext){
          try{
            const fittedPoster=await cropPosterToCell(movie.poster);
            const base64=fittedPoster.substring(fittedPoster.indexOf(',')+1);
            const imageId=workbook.addImage({base64,extension:'png'});
            // Anchor to the exact poster cell boundaries: move and size with cells.
            sheet.addImage(imageId,{tl:{col:col-1,row:row+2},br:{col:col,row:row+3},editAs:'twoCell'});
          }catch(error){console.warn('Could not embed fitted poster for '+movie.name,error);}
        }
      }
      row+=4;
      sheet.getRow(row).height=8;
      row++;
    }
    if(!languageMovies.length){
      sheet.getCell('A1').value=language+' MOVIE COLLECTION';
      sheet.getCell('A1').font={bold:true,size:16,color:{argb:'FF174F7D'}};
      sheet.getCell('A2').value='Movies added in this language will appear here, grouped alphabetically and sorted by release year.';
    }
    sheet.getColumn(1).width=15.5;
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
    if(!sheet){
      const message=document.createElement('p');
      message.className='movie-excel-empty';
      message.textContent='The selected worksheet could not be found. Please choose a worksheet again.';
      preview.appendChild(message);
      return;
    }

    // Always render a predictable 100 × 100 spreadsheet window, even when
    // the source worksheet is mostly blank or has a small used range.
    const maxRow=100;
    const maxCol=100;
    const columnLabel=number=>{
      let label='';
      while(number>0){const remainder=(number-1)%26;label=String.fromCharCode(65+remainder)+label;number=Math.floor((number-1)/26);}
      return label;
    };
    const table=document.createElement('table');
    table.className='excel-grid-table excel-grid-100';
    const head=document.createElement('thead');
    const headRow=document.createElement('tr');
    const corner=document.createElement('th');
    corner.className='excel-grid-corner';
    corner.textContent='';
    headRow.appendChild(corner);
    for(let col=1;col<=maxCol;col++){
      const th=document.createElement('th');
      th.className='excel-column-heading';
      th.textContent=columnLabel(col);
      headRow.appendChild(th);
    }
    head.appendChild(headRow);
    table.appendChild(head);

    const body=document.createElement('tbody');
    for(let rowIndex=1;rowIndex<=maxRow;rowIndex++){
      const row=document.createElement('tr');
      const rowHeading=document.createElement('th');
      rowHeading.className='excel-grid-row-heading';
      rowHeading.textContent=String(rowIndex);
      row.appendChild(rowHeading);
      for(let col=1;col<=maxCol;col++){
        const cell=document.createElement('td');
        const sourceCell=sheet.getCell(rowIndex,col);
        const value=sourceCell.value;
        if(value&&typeof value==='object'){
          if(value.text)cell.textContent=value.text;
          else if(Array.isArray(value.richText))cell.textContent=value.richText.map(part=>part.text||'').join('');
          else if(value.result!==undefined)cell.textContent=String(value.result);
          else if(value.hyperlink)cell.textContent=value.text||value.hyperlink;
          else cell.textContent='';
        }else cell.textContent=String(value??'');
        row.appendChild(cell);
      }
      body.appendChild(row);
    }
    table.appendChild(body);
    const wrap=document.createElement('div');
    wrap.className='movie-excel-scroll excel-grid-scroll';
    wrap.appendChild(table);
    preview.appendChild(wrap);
    const note=document.createElement('p');
    note.className='excel-preview-note';
    note.textContent='Showing rows 1–100 and columns A–CV. Scroll horizontally and vertically to explore the full 100 × 100 preview.';
    preview.appendChild(note);
  }

  function renderSelectedPreview(){
    if(!workbook||!currentPreviewSheet)throw new Error('Select a worksheet to preview.');
    const sheet=workbook.getWorksheet(currentPreviewSheet);
    if(!sheet)throw new Error('The selected worksheet is no longer available. Please select it again.');
    const language=getLanguageForSheet(sheet.name);
    if(language)renderLanguagePreview(language);
    else renderGenericSheet(sheet);
    if(previewWrap)previewWrap.hidden=false;
  }

  function renderPreviewError(error){
    preview.replaceChildren();
    preview.classList.remove('movie-excel-preview');
    const message=document.createElement('p');
    message.className='movie-excel-empty';
    message.textContent='Preview could not be rendered: '+(error?.message||'Unknown worksheet error');
    preview.appendChild(message);
    previewWrap.hidden=false;
  }

  async function syncLanguageSheets(){
    if(!workbook)return;
    removeDuplicateLanguageSheets();
    for(const language of LANGUAGES)await buildLanguageSheet(language);
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
    else if(workbook.getWorksheet(LANGUAGE_SHEETS.Bengali))sheetSelect.value=LANGUAGE_SHEETS.Bengali;
    else sheetSelect.value=workbook.worksheets[0]?.name||'';
  }

  sheetSelect.addEventListener('change',()=>{
    if(!workbook)return;
    const previous=currentPreviewSheet;
    currentPreviewSheet=sheetSelect.value;
    persistCurrentWorkbook().catch(error=>{currentPreviewSheet=previous;console.warn('Could not save worksheet selection:',error);});
  });

  sheetSelect.addEventListener('change',()=>{if(!workbook)return;currentPreviewSheet=sheetSelect.value;persistCurrentWorkbook().catch(error=>console.warn('Could not save worksheet selection:',error));});

  fileInput.addEventListener('change',async()=>{
    const attempt=++uploadAttempt;
    selectedFile=fileInput.files&&fileInput.files[0]||null;
    workbook=null;currentPreviewSheet='';sheetSelectWrap.hidden=true;previewWrap.hidden=true;
    if(completeStage)completeStage.hidden=true;
    if(!selectedFile){fileName.textContent='Choose Excel workbook';fileHint.textContent='Click to browse · Upload required each time';showMessage('Upload a reference workbook to begin.');return;}
    if(!/\.xlsx$/i.test(selectedFile.name)){
      selectedFile=null;fileInput.value='';fileName.textContent='Choose Excel workbook';fileHint.textContent='Please select an .xlsx file';
      showMessage('This formatted movie-collection workflow requires an .xlsx workbook. Please save legacy .xls files as .xlsx and upload again.');return;
    }
    fileName.textContent=selectedFile.name;fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · File selected successfully';
    showMessage('Excel detected. Loading workbook…');
    try{
      const ExcelJS=await loadExcelJS();if(attempt!==uploadAttempt)return;
      workbook=new ExcelJS.Workbook();
      await workbook.xlsx.load(await selectedFile.arrayBuffer());if(attempt!==uploadAttempt)return;
      await syncLanguageSheets();
      sheetSelectWrap.hidden=false;
      await persistCurrentWorkbook();
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
      if(completeStage)completeStage.hidden=true;
      renderSelectedPreview();
      persistCurrentWorkbook().catch(error=>console.warn('Could not save Excel preview selection:',error));
      showMessage('Showing worksheet: '+currentPreviewSheet+'. Add movies to see this preview update live.');
      previewWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      renderPreviewError(error);
      showMessage('Could not display this worksheet: '+(error.message||'Unknown error'));
      previewWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }
  });

  let updateQueue=Promise.resolve();
  subscribeToMovies(()=>{
    if(!workbook)return;
    updateQueue=updateQueue.then(async()=>{
      if(!workbook)return;
      await syncLanguageSheets();
      await persistCurrentWorkbook();
      if(currentPreviewSheet){
        renderSelectedPreview();
        showMessage('Movie collection updated. '+currentPreviewSheet+' preview is synchronized with '+getMovies().filter(movie=>movie.language===getLanguageForSheet(currentPreviewSheet)).length+' movie(s).');
      }
    }).catch(error=>{
      renderPreviewError(error);
      showMessage('Could not refresh the Excel preview: '+(error.message||'Unknown error'));
    });
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
      await updateQueue;
      await syncLanguageSheets();
      const buffer=await workbook.xlsx.writeBuffer();
      const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url=URL.createObjectURL(blob);const link=document.createElement('a');
      link.href=url;link.download=selectedFile.name.replace(/\.xlsx$/i,'')+'_updated.xlsx';
      document.body.appendChild(link);link.click();link.remove();URL.revokeObjectURL(url);
      showMessage('Export started. Language worksheets include the alphabetical/yearly layout, colour bands and embedded posters.');
    }catch(error){showMessage('Could not export workbook: '+(error.message||'Unknown error'));}
    finally{exportButton.disabled=false;}
  });
  // A reset landing must never restore a workbook from an earlier page lifecycle.
  if(resetLanding){
    localStorage.removeItem('chayan-watchlist-reset-in-progress');
    window.history.replaceState({},'',window.location.pathname);
    fileInput.value='';
    fileName.textContent='Choose Excel workbook';
    fileHint.textContent='Click to browse · Upload required each time';
    sheetSelectWrap.hidden=true;
    previewWrap.hidden=true;
  }else{
    // Restore only the Excel workbook and its preview controls after a normal refresh.
    (async()=>{
      const restoreAttempt=uploadAttempt;
      try{
        if(localStorage.getItem('chayan-watchlist-reset-in-progress')==='1')return;
        const saved=await readWorkbookState();
        if(restoreAttempt!==uploadAttempt||resetInProgress||localStorage.getItem('chayan-watchlist-reset-in-progress')==='1')return;
        if(!saved||!saved.buffer)return;
        const ExcelJS=await loadExcelJS();
        if(restoreAttempt!==uploadAttempt||resetInProgress)return;
        workbook=new ExcelJS.Workbook();
        await workbook.xlsx.load(saved.buffer);
        if(restoreAttempt!==uploadAttempt||resetInProgress)return;
      selectedFile={name:saved.fileName||'Saved workbook.xlsx'};
      fileName.textContent=selectedFile.name;
      fileHint.textContent='Restored from this device · Ready to preview';
      refreshSheetOptions();
      sheetSelectWrap.hidden=false;
      currentPreviewSheet=workbook.getWorksheet(saved.sheetName)?.name||sheetSelect.value;
      if(currentPreviewSheet)sheetSelect.value=currentPreviewSheet;
      if(saved.previewVisible&&currentPreviewSheet){renderSelectedPreview();showMessage('Restored your Excel workbook and preview from this device.');}
      else showMessage('Restored your Excel workbook. Select a worksheet to preview.');
      }catch(error){console.warn('Could not restore saved Excel workspace:',error);showMessage('Saved Excel workbook could not be restored. Please upload it again.');}
    })();
  }

}
