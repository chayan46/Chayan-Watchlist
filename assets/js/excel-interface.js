export function initExcelInterface(){
  const fileInput=document.getElementById('excelFile');
  const fileName=document.getElementById('excelFileName');
  const fileHint=document.getElementById('excelFileHint');
  const readyButton=document.getElementById('readyExcel');
  const completeButton=document.getElementById('completeExcel');
  const status=document.getElementById('excelStatus');
  const sheetSelectWrap=document.getElementById('sheetSelectWrap');
  const sheetSelect=document.getElementById('excelSheetSelect');
  const showPreviewButton=document.getElementById('showExcelPreview');
  const previewWrap=document.getElementById('excelPreviewWrap');
  const preview=document.getElementById('excelPreview');
  const downloadPanel=document.getElementById('excelDownload');
  if(!fileInput||!readyButton||!sheetSelectWrap||!sheetSelect||!showPreviewButton||!previewWrap||!preview)return;

  let selectedFile=null;
  let workbook=null;
  let xlsxLibraryPromise=null;

  const showMessage=message=>{status.textContent=message;};

  function loadXLSX(){
    if(window.XLSX)return Promise.resolve(window.XLSX);
    if(xlsxLibraryPromise)return xlsxLibraryPromise;
    xlsxLibraryPromise=new Promise((resolve,reject)=>{
      const urls=[
        'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
        'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
      ];
      const tryLoad=index=>{
        if(index>=urls.length){xlsxLibraryPromise=null;reject(new Error('Excel reader could not be loaded. Check your internet connection and try again.'));return;}
        const script=document.createElement('script');
        script.src=urls[index];
        script.async=true;
        script.onload=()=>window.XLSX?resolve(window.XLSX):tryLoad(index+1);
        script.onerror=()=>{script.remove();tryLoad(index+1);};
        document.head.appendChild(script);
      };
      tryLoad(0);
    });
    return xlsxLibraryPromise;
  }

  fileInput.addEventListener('change',()=>{
    selectedFile=fileInput.files&&fileInput.files[0]||null;
    workbook=null;
    sheetSelectWrap.hidden=true;
    previewWrap.hidden=true;
    downloadPanel.hidden=true;
    readyButton.disabled=true;
    if(!selectedFile){
      fileName.textContent='Choose Excel workbook';
      fileHint.textContent='Click to browse · Upload required each time';
      showMessage('Upload a reference workbook to begin.');
      return;
    }
    const ext=selectedFile.name.split('.').pop().toLowerCase();
    if(!['xlsx','xls'].includes(ext)){
      selectedFile=null;fileInput.value='';
      fileName.textContent='Choose Excel workbook';
      fileHint.textContent='Please select an .xlsx or .xls file';
      showMessage('Unsupported file type. Please upload an Excel workbook.');
      return;
    }
    fileName.textContent=selectedFile.name;
    fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · File selected successfully';
    readyButton.disabled=false;
    showMessage('Workbook selected. Click “Ready to update” to read its worksheet names.');
  });

  readyButton.addEventListener('click',async()=>{
    if(!selectedFile)return;
    readyButton.disabled=true;
    readyButton.textContent='Reading workbook…';
    sheetSelectWrap.hidden=true;previewWrap.hidden=true;
    showMessage('Opening workbook and detecting all worksheets…');
    try{
      const XLSX=await loadXLSX();
      const buffer=await selectedFile.arrayBuffer();
      workbook=XLSX.read(buffer,{type:'array',cellDates:true});
      if(!workbook.SheetNames||!workbook.SheetNames.length)throw new Error('No worksheets were found in this workbook.');
      sheetSelect.replaceChildren();
      workbook.SheetNames.forEach((name,index)=>{
        const option=document.createElement('option');
        option.value=name;
        option.textContent=(index+1)+'. '+name;
        sheetSelect.appendChild(option);
      });
      sheetSelect.value=workbook.SheetNames[0];
      sheetSelectWrap.hidden=false;
      showMessage(workbook.SheetNames.length+' worksheet(s) detected. Choose a sheet, then click “Show selected sheet preview”.');
      sheetSelectWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      workbook=null;
      showMessage(error&&error.message?error.message:'Unable to read this Excel workbook.');
    }finally{
      readyButton.disabled=!selectedFile;
      readyButton.innerHTML='<span aria-hidden="true">▷</span> Ready to update';
    }
  });

  showPreviewButton.addEventListener('click',()=>{
    if(!workbook||!window.XLSX)return;
    try{
      renderSheet(window.XLSX,workbook,sheetSelect.value);
      previewWrap.hidden=false;
      showMessage('Showing worksheet: '+sheetSelect.value);
      previewWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      showMessage('Could not display this worksheet: '+(error.message||'Unknown error'));
    }
  });

  function renderSheet(XLSX,book,sheetName){
    const sheet=book.Sheets[sheetName];
    if(!sheet)throw new Error('The selected worksheet was not found.');

    // Use the worksheet's declared Excel range, not only non-empty JSON rows.
    // Blank formatted cells may have a valid !ref even when sheet_to_json returns [].
    const hasRange=Boolean(sheet['!ref']);
    const range=hasRange?XLSX.utils.decode_range(sheet['!ref']):{s:{r:0,c:0},e:{r:19,c:9}};
    const startRow=range.s.r;
    const startCol=range.s.c;
    const rowCount=Math.min(Math.max(20,range.e.r-startRow+1),100);
    const colCount=Math.min(Math.max(10,range.e.c-startCol+1),20);
    const endRow=startRow+rowCount;
    const endCol=startCol+colCount;

    preview.replaceChildren();
    const scroll=document.createElement('div');
    scroll.className='excel-table-scroll';
    const table=document.createElement('table');
    table.className='excel-data-table excel-grid-table';
    const thead=document.createElement('thead');
    const headRow=document.createElement('tr');
    const corner=document.createElement('th');
    corner.className='excel-grid-corner';
    corner.textContent='';
    headRow.appendChild(corner);
    for(let c=startCol;c<endCol;c++){
      const th=document.createElement('th');
      th.className='excel-column-heading';
      th.textContent=XLSX.utils.encode_col(c);
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody=document.createElement('tbody');
    let populatedCells=0;
    for(let r=startRow;r<endRow;r++){
      const tr=document.createElement('tr');
      const rowHead=document.createElement('th');
      rowHead.className='excel-row-heading';
      rowHead.textContent=String(r+1);
      tr.appendChild(rowHead);
      for(let c=startCol;c<endCol;c++){
        const td=document.createElement('td');
        const address=XLSX.utils.encode_cell({r,c});
        const cell=sheet[address];
        const value=cell&&cell.v!==undefined?String(cell.w!==undefined?cell.w:cell.v):'';
        td.textContent=value;
        if(value!=='')populatedCells++;
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    scroll.appendChild(table);
    preview.appendChild(scroll);

    const note=document.createElement('p');
    note.className='excel-preview-note';
    if(populatedCells===0){
      note.textContent='This worksheet has no entered values. Showing an editable-style blank grid with '+rowCount+' rows and '+colCount+' columns.';
    }else{
      note.textContent='Showing '+rowCount+' rows and '+colCount+' columns from “'+sheetName+'” ('+populatedCells+' non-empty cells in this preview range).';
    }
    preview.appendChild(note);
  }

  completeButton?.addEventListener('click',()=>{
    downloadPanel.hidden=false;
    showMessage('Preview stage complete. Movie-to-workbook updates and download generation are still pending integration.');
  });
}
