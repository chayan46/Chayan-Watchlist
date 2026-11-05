import {getMovies,subscribeToMovies} from './data-management.js';

const LANGUAGE_SHEETS={Bengali:'Bengali Movies',Hindi:'Hindi Movies',English:'English Movies'};
const LANGUAGES=Object.keys(LANGUAGE_SHEETS);
const getLanguageForSheet=name=>LANGUAGES.find(language=>[LANGUAGE_SHEETS[language],language].includes(name))||null;
const EXCELJS_CDN='https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
const WORKBOOK_DB='chayan-watchlist-workbook';
const WORKBOOK_STORE='saved-workbook';
const WORKBOOK_KEY='last-upload';
function openWorkbookStore(){
  return new Promise((resolve,reject)=>{
    const request=indexedDB.open(WORKBOOK_DB,1);
    request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains(WORKBOOK_STORE))request.result.createObjectStore(WORKBOOK_STORE);};
    request.onsuccess=()=>resolve(request.result);
    request.onerror=()=>reject(request.error||new Error('Could not open local workbook storage.'));
  });
}
async function saveWorkbookLocally(file){
  const db=await openWorkbookStore();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction(WORKBOOK_STORE,'readwrite');
    tx.objectStore(WORKBOOK_STORE).put({name:file.name,type:file.type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',lastModified:file.lastModified||Date.now(),data:file},WORKBOOK_KEY);
    tx.oncomplete=()=>{db.close();resolve();};
    tx.onerror=()=>{const error=tx.error;db.close();reject(error||new Error('Could not save the workbook locally.'));};
    tx.onabort=()=>{const error=tx.error;db.close();reject(error||new Error('Workbook storage was interrupted.'));};
  });
}
async function loadSavedWorkbook(){
  const db=await openWorkbookStore();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction(WORKBOOK_STORE,'readonly');
    const request=tx.objectStore(WORKBOOK_STORE).get(WORKBOOK_KEY);
    request.onsuccess=()=>resolve(request.result||null);
    request.onerror=()=>reject(request.error||new Error('Could not restore the saved workbook.'));
    tx.oncomplete=()=>db.close();
    tx.onerror=()=>db.close();
  });
}
async function clearSavedWorkbook(){
  const db=await openWorkbookStore();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction(WORKBOOK_STORE,'readwrite');
    tx.objectStore(WORKBOOK_STORE).delete(WORKBOOK_KEY);
    tx.oncomplete=()=>{db.close();resolve();};
    tx.onerror=()=>{const error=tx.error;db.close();reject(error||new Error('Could not clear the saved workbook.'));};
  });
}
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
  const fileName=document.getElementById('excelFileName');
  const fileHint=document.getElementById('excelFileHint');
  const uploadNewButton=document.getElementById('uploadNewExcel');
  const uploadLabel=document.getElementById('excelUploadLabel');
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
  function setWorkbookLoaded(loaded){
    if(uploadLabel)uploadLabel.hidden=loaded;
    if(uploadNewButton)uploadNewButton.hidden=!loaded;
    sheetSelectWrap.hidden=!loaded;
  }
  async function restoreWorkbook(record){
    if(!record?.data)return;
    const attempt=++uploadAttempt;
    selectedFile=new File([record.data],record.name||'saved-workbook.xlsx',{type:record.type||record.data.type,lastModified:record.lastModified||Date.now()});
    fileName.textContent=selectedFile.name;
    fileHint.textContent='Restored from this device · Saved until you upload a new workbook';
    showMessage('Restoring your last uploaded Excel workbook…');
    const ExcelJS=await loadExcelJS();
    if(attempt!==uploadAttempt)return;
    workbook=new ExcelJS.Workbook();
    await workbook.xlsx.load(await selectedFile.arrayBuffer());
    if(attempt!==uploadAttempt)return;
    // Do not rebuild/overwrite the uploaded workbook's language sheets during
    // restore. The preview must reflect the exact file the user selected.
    refreshSheetOptions();
    setWorkbookLoaded(true);
    currentPreviewSheet=sheetSelect.value||'';
    if(currentPreviewSheet)renderSelectedPreview();
    showMessage('Your last uploaded workbook has been restored. Continue editing, previewing or exporting.');
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
    sheet.columns=[{width:9},...Array.from({length:Math.max(1,sortMovies(getMovies().filter(m=>m.language===language)).length)},()=>({width:27}))];
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
      groupCell.value=letter;
      groupCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFDCE5F0'}};
      groupCell.font={name:'Aptos Display',size:24,bold:true,color:{argb:'FF172B4D'}};
      groupCell.alignment={vertical:'middle',horizontal:'center'};
      groupCell.border={top:{style:'medium',color:{argb:'FFB8C6D8'}},bottom:{style:'medium',color:{argb:'FFB8C6D8'}},left:{style:'medium',color:{argb:'FFB8C6D8'}},right:{style:'medium',color:{argb:'FFB8C6D8'}}};
      sheet.getRow(row).height=25;
      sheet.getRow(row+1).height=25;
      sheet.getRow(row+2).height=22;
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
        title.alignment={horizontal:'center',vertical:'middle',wrapText:true};
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
            // ExcelJS expects the complete data URI in its base64 field.
            // Keeping the MIME prefix is essential for the generated XLSX media part.
            const imageId=workbook.addImage({base64:fittedPoster,extension:'png'});
            // Anchor to the exact poster cell boundaries. Omit editAs: ExcelJS defaults to the OOXML two-cell anchor; explicitly writing 'twoCell' can create an unsupported editAs value in some spreadsheet viewers.
            sheet.addImage(imageId,{tl:{col:col-1,row:row+2},br:{col:col,row:row+3}});
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
    if(!sheet){
      const message=document.createElement('p');
      message.className='movie-excel-empty';
      message.textContent='The selected worksheet could not be found. Please choose a worksheet again.';
      preview.appendChild(message);
      return;
    }
    const table=document.createElement('table');
    table.className='excel-grid-table';
    // Empty workbooks/sheets still need a useful spreadsheet-style canvas.
    // ExcelJS reports zero dimensions for a truly blank worksheet, so provide
    // a visible grid instead of collapsing the preview to a tiny empty cell.
    const hasCellContent=Boolean(sheet.actualRowCount||sheet.actualColumnCount);
    const maxRow=Math.min(Math.max(sheet.rowCount||sheet.actualRowCount||20,20),100);
    const maxCol=Math.min(Math.max(sheet.columnCount||sheet.actualColumnCount||10,10),30);
    const head=document.createElement('thead');
    const headRow=document.createElement('tr');
    const corner=document.createElement('th');
    corner.textContent='';
    headRow.appendChild(corner);
    for(let col=1;col<=maxCol;col++){
      const th=document.createElement('th');
      th.textContent=col<=26?String.fromCharCode(64+col):'C'+col;
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
        const value=sheet.getCell(rowIndex,col).value;
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
    wrap.className='movie-excel-scroll';
    wrap.appendChild(table);
    preview.appendChild(wrap);
    if(!hasCellContent){
      const note=document.createElement('p');
      note.className='excel-preview-note';
      note.textContent='Blank worksheet · Empty cells are shown in spreadsheet table view. Start entering data in Excel to populate this sheet.';
      preview.appendChild(note);
    }
  }

  function renderSelectedPreview(){
    if(!workbook||!currentPreviewSheet)throw new Error('Select a worksheet to preview.');
    const sheet=workbook.getWorksheet(currentPreviewSheet);
    if(!sheet)throw new Error('The selected worksheet is no longer available. Please select it again.');
    // Preview the actual worksheet parsed from the currently selected XLSX.
    // Language-sheet regeneration belongs to live movie updates/export, not upload.
    renderGenericSheet(sheet);
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

  function refreshSheetOptions(resetSelection=false){
    if(!workbook)return;
    const previous=resetSelection?'':sheetSelect.value;
    sheetSelect.replaceChildren();
    workbook.worksheets.forEach((sheet,index)=>{
      const option=document.createElement('option');option.value=sheet.name;option.textContent=(index+1)+'. '+sheet.name;sheetSelect.appendChild(option);
    });
    if(workbook.getWorksheet(previous))sheetSelect.value=previous;
    else if(workbook.getWorksheet(LANGUAGE_SHEETS.Bengali))sheetSelect.value=LANGUAGE_SHEETS.Bengali;
    else sheetSelect.value=workbook.worksheets[0]?.name||'';
  }

  fileInput.addEventListener('change',async()=>{
    const attempt=++uploadAttempt;
    // Clear the native input after capturing the File. This lets users select
    // the same workbook again and still receive a change event (including on Android).
    const newlySelectedFile=fileInput.files&&fileInput.files[0]||null;
    fileInput.value='';
    if(!newlySelectedFile)return; // Picker cancelled: preserve the current workbook and preview.
    selectedFile=newlySelectedFile;
    workbook=null;currentPreviewSheet='';setWorkbookLoaded(false);previewWrap.hidden=true;
    preview.replaceChildren();
    if(completeStage)completeStage.hidden=true;
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
      // Keep the newly uploaded workbook intact for preview. Rebuilding the
      // language sheets here replaced the selected file's contents with the
      // browser's older movie collection, which made the preview look stale.
      refreshSheetOptions(true);
      currentPreviewSheet='';
      previewWrap.hidden=true;
      preview.replaceChildren();
      preview.classList.remove('movie-excel-preview');
      await saveWorkbookLocally(selectedFile);
      setWorkbookLoaded(true);
      // Re-render from the newly parsed workbook, replacing any previous preview.
      currentPreviewSheet=sheetSelect.options[0]?.value||workbook.worksheets[0]?.name||'';
      if(currentPreviewSheet){
        sheetSelect.value=currentPreviewSheet;
        renderSelectedPreview();
      }
      showMessage('New workbook loaded and preview refreshed: '+selectedFile.name+'. It is saved on this device until you upload another workbook.');
      sheetSelectWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      if(attempt!==uploadAttempt)return;
      workbook=null;setWorkbookLoaded(false);showMessage('Excel upload failed: '+(error.message||'Unable to read this workbook.'));
    }
  });

  uploadNewButton?.addEventListener('click',()=>{
    // Keep the current workbook intact unless a replacement file is actually selected and loaded.
    fileInput.click();
  });

  showPreviewButton.addEventListener('click',()=>{
    if(!workbook)return;
    currentPreviewSheet=sheetSelect.value;
    if(!currentPreviewSheet)return;
    try{
      if(completeStage)completeStage.hidden=true;
      renderSelectedPreview();
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

  loadSavedWorkbook().then(record=>{if(record)restoreWorkbook(record);}).catch(error=>showMessage('Could not restore the saved workbook: '+(error.message||'Local storage error')));

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
}
