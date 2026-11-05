export function initExcelInterface(){
  const fileInput=document.getElementById('excelFile');
  const fileName=document.getElementById('excelFileName');
  const fileHint=document.getElementById('excelFileHint');
  const readyButton=document.getElementById('readyExcel');
  const completeButton=document.getElementById('completeExcel');
  const status=document.getElementById('excelStatus');
  const previewWrap=document.getElementById('excelPreviewWrap');
  const preview=document.getElementById('excelPreview');
  const downloadPanel=document.getElementById('excelDownload');
  if(!fileInput||!readyButton||!previewWrap||!preview)return;

  let selectedFile=null;
  let workbook=null;

  const showMessage=(message)=>{status.textContent=message;};

  fileInput.addEventListener('change',()=>{
    selectedFile=fileInput.files&&fileInput.files[0]||null;
    workbook=null;
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
      selectedFile=null;
      fileInput.value='';
      fileName.textContent='Choose Excel workbook';
      fileHint.textContent='Please select an .xlsx or .xls file';
      showMessage('Unsupported file type. Please upload an Excel workbook.');
      return;
    }
    fileName.textContent=selectedFile.name;
    fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · File selected successfully';
    readyButton.disabled=false;
    showMessage('Workbook uploaded successfully. Click “Ready to update” to load its contents below.');
  });

  readyButton.addEventListener('click',async()=>{
    if(!selectedFile)return;
    readyButton.disabled=true;
    readyButton.textContent='Loading workbook…';
    showMessage('Reading Excel workbook and preparing preview…');
    try{
      const XLSX=await import('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm');
      const buffer=await selectedFile.arrayBuffer();
      workbook=XLSX.read(buffer,{type:'array'});
      if(!workbook.SheetNames.length)throw new Error('No worksheets were found in this workbook.');
      renderSheet(XLSX,workbook,workbook.SheetNames[0]);
      previewWrap.hidden=false;
      downloadPanel.hidden=true;
      showMessage('Workbook loaded: '+workbook.SheetNames.length+' worksheet(s). Preview is ready.');
      previewWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      previewWrap.hidden=false;
      preview.innerHTML='<div class="preview-error"><strong>Could not read this workbook</strong><p></p></div>';
      preview.querySelector('.preview-error p').textContent=error&&error.message?error.message:'Please check the file and try again.';
      showMessage('Excel preview could not be loaded. Please try another .xlsx or .xls file.');
    }finally{
      readyButton.disabled=!selectedFile;
      readyButton.innerHTML='<span aria-hidden="true">▷</span> Ready to update';
    }
  });

  function renderSheet(XLSX,book,sheetName){
    const sheet=book.Sheets[sheetName];
    const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});
    const maxRows=Math.min(rows.length,100);
    const maxCols=Math.min(Math.max(0,...rows.slice(0,maxRows).map(row=>row.length)),20);
    preview.replaceChildren();

    const bar=document.createElement('div');
    bar.className='excel-preview-toolbar';
    const label=document.createElement('strong');
    label.textContent='Worksheet';
    const selector=document.createElement('select');
    selector.className='excel-sheet-select';
    selector.setAttribute('aria-label','Select worksheet');
    book.SheetNames.forEach(name=>{
      const option=document.createElement('option');
      option.value=name;option.textContent=name;selector.appendChild(option);
    });
    selector.value=sheetName;
    selector.addEventListener('change',()=>renderSheet(XLSX,book,selector.value));
    bar.append(label,selector);
    preview.appendChild(bar);

    const scroll=document.createElement('div');
    scroll.className='excel-table-scroll';
    const table=document.createElement('table');
    table.className='excel-data-table';
    const tbody=document.createElement('tbody');
    for(let r=0;r<maxRows;r++){
      const tr=document.createElement('tr');
      for(let c=0;c<maxCols;c++){
        const cell=document.createElement(r===0?'th':'td');
        cell.textContent=rows[r]&&rows[r][c]!==undefined?String(rows[r][c]):'';
        tr.appendChild(cell);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    scroll.appendChild(table);
    preview.appendChild(scroll);
    const note=document.createElement('p');
    note.className='excel-preview-note';
    note.textContent='Showing '+maxRows+' of '+rows.length+' rows and up to '+maxCols+' columns from “'+sheetName+'”.';
    preview.appendChild(note);
  }

  completeButton?.addEventListener('click',()=>{
    downloadPanel.hidden=false;
    showMessage('Workbook preview is loaded. Movie-to-workbook updates and download generation are the next integration stage.');
  });
}
