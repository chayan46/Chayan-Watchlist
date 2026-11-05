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

    // A virtualized 1000 × 1000 spreadsheet: the scroll area represents
    // all cells, while only cells near the visible viewport are rendered.
    const ROWS=1000;
    const COLS=1000;
    const ROW_HEIGHT=35;
    const COL_WIDTH=92;
    const ROW_HEADER=42;
    const COL_HEADER=34;
    const totalWidth=ROW_HEADER+COLS*COL_WIDTH;
    const totalHeight=COL_HEADER+ROWS*ROW_HEIGHT;

    preview.replaceChildren();
    const scroll=document.createElement('div');
    scroll.className='excel-table-scroll excel-virtual-scroll';
    scroll.setAttribute('role','region');
    scroll.setAttribute('aria-label','Excel worksheet preview. Scroll horizontally and vertically to view 1000 rows and 1000 columns.');
    const canvas=document.createElement('div');
    canvas.className='excel-virtual-canvas';
    canvas.style.setProperty('display','block','important');
    canvas.style.setProperty('position','relative','important');
    canvas.style.setProperty('width',totalWidth+'px','important');
    canvas.style.setProperty('height',totalHeight+'px','important');
    canvas.style.setProperty('min-width',totalWidth+'px','important');
    canvas.style.setProperty('min-height',totalHeight+'px','important');
    canvas.style.setProperty('max-width','none','important');
    canvas.style.setProperty('flex','none','important');
    scroll.appendChild(canvas);
    preview.appendChild(scroll);

    let populatedCells=0;
    let renderFrame=0;
    function addCell(className,text,left,top,width,height){
      const el=document.createElement('div');
      el.className=className;
      el.textContent=text;
      el.style.setProperty('position','absolute','important');
      el.style.setProperty('display','flex','important');
      el.style.setProperty('left',left+'px','important');
      el.style.setProperty('top',top+'px','important');
      if(width)el.style.setProperty('width',width+'px','important');
      if(height)el.style.setProperty('height',height+'px','important');
      el.style.setProperty('box-sizing','border-box','important');
      canvas.appendChild(el);
      return el;
    }

    function drawViewport(){
      if(renderFrame)cancelAnimationFrame(renderFrame);
      renderFrame=requestAnimationFrame(()=>{
        renderFrame=0;
        canvas.replaceChildren();
        const viewHeight=scroll.clientHeight||440;
        const viewWidth=scroll.clientWidth||Math.min(window.innerWidth-80,700);
        const firstRow=Math.max(0,Math.floor(Math.max(0,scroll.scrollTop-COL_HEADER)/ROW_HEIGHT));
        const firstCol=Math.max(0,Math.floor(Math.max(0,scroll.scrollLeft-ROW_HEADER)/COL_WIDTH));
        const visibleRows=Math.ceil(viewHeight/ROW_HEIGHT)+2;
        const visibleCols=Math.ceil(viewWidth/COL_WIDTH)+2;
        const lastRow=Math.min(ROWS,firstRow+visibleRows);
        const lastCol=Math.min(COLS,firstCol+visibleCols);

        // Sticky worksheet column headers and row numbers.
        addCell('excel-virtual-corner','',scroll.scrollLeft,scroll.scrollTop,ROW_HEADER,COL_HEADER);
        for(let c=firstCol;c<lastCol;c++){
          const colName=XLSX.utils.encode_col(c);
          addCell('excel-virtual-colhead',colName,ROW_HEADER+c*COL_WIDTH,scroll.scrollTop,COL_WIDTH,COL_HEADER);
        }
        for(let r=firstRow;r<lastRow;r++){
          addCell('excel-virtual-rowhead',String(r+1),scroll.scrollLeft,COL_HEADER+r*ROW_HEIGHT,ROW_HEADER,ROW_HEIGHT);
          for(let c=firstCol;c<lastCol;c++){
            const address=XLSX.utils.encode_cell({r,c});
            const cell=sheet[address];
            const value=cell&&cell.v!==undefined?String(cell.w!==undefined?cell.w:cell.v):'';
            if(value!=='')populatedCells++;
            addCell('excel-virtual-cell',value,ROW_HEADER+c*COL_WIDTH,COL_HEADER+r*ROW_HEIGHT,COL_WIDTH,ROW_HEIGHT);
          }
        }
      });
    }
    scroll.addEventListener('scroll',drawViewport,{passive:true});
    drawViewport();

    const note=document.createElement('p');
    note.className='excel-preview-note';
    const declared=sheet['!ref']||'No populated cell range';
    note.textContent='1000 rows × 1000 columns available · Worksheet data range: '+declared+'. Scroll in both directions to explore the grid.';
    preview.appendChild(note);
  }

  completeButton?.addEventListener('click',()=>{
    downloadPanel.hidden=false;
    showMessage('Preview stage complete. Movie-to-workbook updates and download generation are still pending integration.');
  });
}
