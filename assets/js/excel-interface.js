export function initExcelInterface(){
  const fileInput=document.getElementById('excelFile');
  const fileName=document.getElementById('excelFileName');
  const fileHint=document.getElementById('excelFileHint');
  const readyButton=document.getElementById('readyExcel');
  const completeButton=document.getElementById('completeExcel');
  const status=document.getElementById('excelStatus');
  const previewWrap=document.getElementById('excelPreviewWrap');
  const downloadPanel=document.getElementById('excelDownload');
  if(!fileInput||!readyButton)return;
  let selectedFile=null;
  fileInput.addEventListener('change',()=>{
    selectedFile=fileInput.files&&fileInput.files[0]||null;
    previewWrap.hidden=true;
    downloadPanel.hidden=true;
    completeButton.disabled=false;
    if(!selectedFile){fileName.textContent='Choose Excel workbook';fileHint.textContent='.xlsx or .xls · Upload required each time';readyButton.disabled=true;status.textContent='Upload a reference workbook to begin.';return;}
    const ext=selectedFile.name.split('.').pop().toLowerCase();
    if(!['xlsx','xls'].includes(ext)){selectedFile=null;fileInput.value='';fileName.textContent='Choose Excel workbook';fileHint.textContent='Please select an .xlsx or .xls file';readyButton.disabled=true;status.textContent='Unsupported file type. Please upload an Excel workbook.';return;}
    fileName.textContent=selectedFile.name;
    fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · File selected successfully';
    readyButton.disabled=false;
    status.textContent='Reference workbook selected. Click “Ready to update” to continue.';
  });
  readyButton.addEventListener('click',()=>{
    if(!selectedFile)return;
    previewWrap.hidden=false;
    downloadPanel.hidden=true;
    status.textContent='Ready state confirmed. The live workbook preview will be connected in the Excel integration stage.';
    previewWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
  });
  completeButton.addEventListener('click',()=>{
    downloadPanel.hidden=false;
    status.textContent='Update completion stage selected. Excel editing and download are not enabled yet.';
  });
}