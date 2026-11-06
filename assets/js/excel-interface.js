import {getMovies,subscribeToMovies,deleteMovie} from './data-management.js';

const LANGUAGE_SHEETS={Bengali:'Bengali Movies',Hindi:'Hindi Movies',English:'English Movies'};
const LANGUAGES=Object.keys(LANGUAGE_SHEETS);
const ALPHABET_COLORS=["F4CCCC","FCE5CD","FFF2CC","D9EAD3","D0E0E3","CFE2F3","D9D2E9","EAD1DC","F4B6B6","F9CB9C","FFE599","B6D7A8","A2C4C9","9FC5E8","B4A7D6","D5A6BD","E6B8AF","F6D7B0","DDEBF7","D9EAD3","EADCF8","FCE4D6","C9DAF8","D0E0E3","E2F0D9","F4CCCC"];
const getAlphabetColor=letter=>{const index=String(letter||'').toUpperCase().charCodeAt(0)-65;return index>=0&&index<26?ALPHABET_COLORS[index]:'E8EDF3';};
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
  let previewSearchQuery='';
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
    async function writeMovieGroup(label,items,isFranchise=false){
      sheet.mergeCells(row,1,row+2,1);
      const groupCell=sheet.getCell(row,1);
      if(isFranchise){
        groupCell.value={richText:[
          {text:label,font:{name:'Aptos Display',size:32,bold:true,color:{argb:'FF263445'}}},
          {text:'\n('+items.length+')',font:{name:'Aptos',size:18,bold:true,italic:true,color:{argb:'FF147D75'}}}
        ]};
        groupCell.font={name:'Aptos Display',size:32,bold:true,color:{argb:'FF263445'}};
        groupCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+getAlphabetColor(label.charAt(0))}};
        groupCell.font={name:'Aptos Display',size:32,bold:true,color:{argb:'FF263445'}};
      }else{
        const movieCount=items.length;
        groupCell.value={richText:[
          {text:label,font:{name:'Aptos Display',size:50,bold:true,color:{argb:'FF172B4D'}}},
          {text:'\n('+movieCount+')',font:{name:'Aptos',size:25,bold:true,italic:true,color:{argb:'FF147D75'}}}
        ]};
        groupCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+getAlphabetColor(label.charAt(0))}};
      }
      groupCell.alignment={vertical:'middle',horizontal:'center',wrapText:true};
      groupCell.border={top:{style:'medium',color:{argb:'FFB8C6D8'}},bottom:{style:'medium',color:{argb:'FFB8C6D8'}},left:{style:'medium',color:{argb:'FFB8C6D8'}},right:{style:'medium',color:{argb:'FFB8C6D8'}}};
      sheet.getRow(row).height=19;
      sheet.getRow(row+1).height=16;
      sheet.getRow(row+2).height=190;
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
        const posterCell=sheet.getCell(row+2,col);
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
            sheet.addImage(imageId,{tl:{col:col-1,row:row+1},br:{col:col,row:row+2},editAs:'twoCell'});
          }catch(error){console.warn('Could not embed fitted poster for '+movie.name,error);}
        }
      }
      row+=3;
      sheet.getRow(row).height=8;
      row++;
    }
    for(const [letter,items] of [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0]))){
      const standardMovies=sortMovies(items.filter(movie=>movie.movieType!=='franchise'));
      const franchiseMap=new Map();
      items.filter(movie=>movie.movieType==='franchise'&&movie.franchiseName).forEach(movie=>{
        const key=movie.franchiseName.trim().toLocaleLowerCase();
        if(!franchiseMap.has(key))franchiseMap.set(key,{name:movie.franchiseName.trim(),movies:[]});
        franchiseMap.get(key).movies.push(movie);
      });
      if(standardMovies.length)await writeMovieGroup(letter,standardMovies,false);
      const legacyNames=[...new Set(languageMovies.filter(movie=>(movie.name.trim().charAt(0)||'#').toLocaleUpperCase()===letter&&movie.movieType==='franchise'&&movie.franchiseName&&!movie.franchiseCode).map(movie=>movie.franchiseName.trim()))].sort((a,b)=>a.localeCompare(b,undefined,{sensitivity:'base'}));
      const franchises=[...franchiseMap.values()].map(franchise=>{
        const explicit=franchise.movies.map(movie=>movie.franchiseCode).find(Boolean);
        const legacyIndex=legacyNames.findIndex(name=>name.toLocaleLowerCase()===franchise.name.toLocaleLowerCase());
        return {...franchise,code:explicit||letter+(legacyIndex+1)};
      }).sort((a,b)=>Number(a.code.match(/\d+$/)?.[0]||9999)-Number(b.code.match(/\d+$/)?.[0]||9999));
      for(const franchise of franchises){
        await writeMovieGroup(franchise.code,sortMovies(franchise.movies),true);
      }
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
    const allMovies=sortMovies(getMovies().filter(movie=>movie.language===language));
    const query=previewSearchQuery.trim();
    let movies=allMovies;
    if(query){
      if(/^[a-z]$/i.test(query)){
        movies=allMovies.filter(movie=>(movie.name.trim().charAt(0)||'#').toLocaleUpperCase()===query.toLocaleUpperCase());
      }else if(/^\d{4}$/.test(query)){
        movies=allMovies.filter(movie=>String(movie.year)===query);
      }else{
        movies=allMovies.filter(movie=>movie.name.toLocaleLowerCase().includes(query.toLocaleLowerCase()));
      }
    }
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
      empty.textContent=query?'No '+language+' movies match “'+query+'”. Clear the search to show all movies.':'No '+language+' movies yet. Add a movie using the form; this worksheet will update automatically.';
      preview.appendChild(empty);
      return;
    }
    function appendPreviewGroup(label,items,isFranchise=false){
      const groupRow=document.createElement('tr');
      const alphabet=document.createElement('th');
      alphabet.className=isFranchise?'movie-excel-letter movie-excel-franchise':'movie-excel-letter';
      alphabet.style.setProperty('--movie-alphabet-color','#'+getAlphabetColor(label.charAt(0)));
      alphabet.rowSpan=3;
      if(isFranchise){
        const subgroupLabel=document.createElement('span');
        subgroupLabel.className='movie-excel-franchise-label';
        subgroupLabel.textContent=label;
        const subgroupCount=document.createElement('span');
        subgroupCount.className='movie-excel-franchise-count';
        subgroupCount.textContent='('+items.length+')';
        alphabet.append(subgroupLabel,subgroupCount);
      }else{
        const movieCount=items.length;
        const letterLabel=document.createElement('span');
        letterLabel.className='movie-excel-letter-label';
        letterLabel.textContent=label;
        const posterCountLabel=document.createElement('span');
        posterCountLabel.className='movie-excel-poster-count';
        posterCountLabel.textContent='('+movieCount+')';
        alphabet.append(letterLabel,posterCountLabel);
      }
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
      const posterRow=document.createElement('tr');
      items.forEach(movie=>{
        const cell=document.createElement('td');cell.className='movie-excel-poster';
        cell.dataset.movieId=movie.id;
        if(movie.poster){const img=document.createElement('img');img.src=movie.poster;img.alt=movie.name+' poster';cell.appendChild(img);}
        else{const empty=document.createElement('span');empty.className='movie-excel-no-poster';empty.textContent='Poster not supplied';cell.appendChild(empty);}
        const deleteButton=document.createElement('button');
        deleteButton.type='button';
        deleteButton.className='movie-excel-delete';
        deleteButton.dataset.deleteMovieId=movie.id;
        deleteButton.textContent='Delete entry';
        deleteButton.setAttribute('aria-label','Delete '+movie.name+' from '+language+' movies');
        cell.appendChild(deleteButton);
        posterRow.appendChild(cell);
      });
      body.appendChild(posterRow);
      const gap=document.createElement('tr');const gapCell=document.createElement('td');gapCell.colSpan=items.length+1;gapCell.className='movie-excel-gap';gap.appendChild(gapCell);body.appendChild(gap);
    }
    for(const [letter,items] of [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0]))){
      const standardMovies=sortMovies(items.filter(movie=>movie.movieType!=='franchise'));
      const franchiseMap=new Map();
      items.filter(movie=>movie.movieType==='franchise'&&movie.franchiseName).forEach(movie=>{
        const key=movie.franchiseName.trim().toLocaleLowerCase();
        if(!franchiseMap.has(key))franchiseMap.set(key,{name:movie.franchiseName.trim(),movies:[]});
        franchiseMap.get(key).movies.push(movie);
      });
      if(standardMovies.length)appendPreviewGroup(letter,standardMovies,false);
      const legacyNames=[...new Set(allMovies.filter(movie=>(movie.name.trim().charAt(0)||'#').toLocaleUpperCase()===letter&&movie.movieType==='franchise'&&movie.franchiseName&&!movie.franchiseCode).map(movie=>movie.franchiseName.trim()))].sort((a,b)=>a.localeCompare(b,undefined,{sensitivity:'base'}));
      const franchises=[...franchiseMap.values()].map(franchise=>{
        const explicit=franchise.movies.map(movie=>movie.franchiseCode).find(Boolean);
        const legacyIndex=legacyNames.findIndex(name=>name.toLocaleLowerCase()===franchise.name.toLocaleLowerCase());
        return {...franchise,code:explicit||letter+(legacyIndex+1)};
      }).sort((a,b)=>Number(a.code.match(/\d+$/)?.[0]||9999)-Number(b.code.match(/\d+$/)?.[0]||9999));
      for(const franchise of franchises){
        appendPreviewGroup(franchise.code,sortMovies(franchise.movies),true);
      }
    }
    table.appendChild(body);
    const wrap=document.createElement('div');wrap.className='movie-excel-scroll';wrap.appendChild(table);preview.appendChild(wrap);
    const note=document.createElement('p');note.className='excel-preview-note';note.textContent=language+' worksheet · '+movies.length+' of '+allMovies.length+' movie(s) shown'+(query?' · Filter: '+(/^[a-z]$/i.test(query)?'first alphabet':/^\d{4}$/.test(query)?'release year':'movie name'):'')+' · Alphabetical groups A–Z · Release year oldest to newest.';preview.appendChild(note);
  }

  // One compact search field filters the selected movie worksheet by title, first letter, or year.
  const previewHeading=previewWrap.querySelector('.stage-heading');
  const previewSearch=document.createElement('div');
  previewSearch.className='movie-preview-search';
  const previewSearchInput=document.createElement('input');
  previewSearchInput.type='search';
  previewSearchInput.className='movie-preview-search-input';
  previewSearchInput.placeholder='Search movie name, A–Z or year…';
  previewSearchInput.setAttribute('aria-label','Search movies by name, first alphabet, or release year');
  previewSearchInput.autocomplete='off';
  previewSearchInput.value=previewSearchQuery;
  previewSearch.appendChild(previewSearchInput);
  if(previewHeading)previewHeading.appendChild(previewSearch);
  previewSearchInput.addEventListener('input',()=>{
    previewSearchQuery=previewSearchInput.value;
    // Source workbook previews are not filtered by the movie collection search.
  });

  function excelCellText(value){
    if(value===null||value===undefined)return '';
    if(value instanceof Date)return value.toLocaleString();
    if(typeof value==='object'){
      if(Array.isArray(value.richText))return value.richText.map(part=>part.text||'').join('');
      if(value.text!==undefined)return String(value.text);
      if(value.result!==undefined)return excelCellText(value.result);
      if(value.hyperlink)return String(value.text||value.hyperlink);
      if(value.formula)return value.result!==undefined?excelCellText(value.result):'='+value.formula;
      if(value.error)return String(value.error);return '';
    }return String(value);
  }
  function colorCss(color){if(!color)return '';if(color.argb)return '#'+String(color.argb).slice(-6);if(color.rgb)return '#'+String(color.rgb).slice(-6);return '';}
  function renderGenericSheet(sheet){
    preview.replaceChildren();preview.classList.remove('movie-excel-preview');previewSearch.hidden=true;
    if(!sheet){const p=document.createElement('p');p.className='movie-excel-empty';p.textContent='The selected worksheet could not be found.';preview.appendChild(p);return;}
    const maxRow=Math.max(1,sheet.rowCount||1),maxCol=Math.max(1,sheet.columnCount||1);
    const colLabel=n=>{let x='';while(n){const k=(n-1)%26;x=String.fromCharCode(65+k)+x;n=Math.floor((n-1)/26);}return x;};
    const table=document.createElement('table');table.className='excel-source-grid';table.style.width='max-content';table.style.minWidth='100%';
    const cg=document.createElement('colgroup'),rowCol=document.createElement('col');rowCol.style.width='44px';cg.appendChild(rowCol);
    for(let c=1;c<=maxCol;c++){const col=document.createElement('col'),w=sheet.getColumn(c).width;col.style.width=(w?Math.max(32,Math.round(w*7+5)):92)+'px';cg.appendChild(col);}table.appendChild(cg);
    const thead=document.createElement('thead'),hr=document.createElement('tr'),corner=document.createElement('th');corner.className='excel-source-corner';hr.appendChild(corner);
    for(let c=1;c<=maxCol;c++){const th=document.createElement('th');th.textContent=colLabel(c);hr.appendChild(th);}thead.appendChild(hr);table.appendChild(thead);
    const merges=new Map(),covered=new Set();
    (sheet.model.merges||[]).forEach(range=>{const m=String(range).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);if(!m)return;
      const num=x=>{let n=0;for(const ch of x.toUpperCase())n=n*26+ch.charCodeAt(0)-64;return n;};
      const r1=+m[2],c1=num(m[1]),r2=+m[4],c2=num(m[3]);merges.set(r1+':'+c1,{rowSpan:r2-r1+1,colSpan:c2-c1+1});
      for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)if(r!==r1||c!==c1)covered.add(r+':'+c);
    });
    const body=document.createElement('tbody');
    for(let r=1;r<=maxRow;r++){const row=sheet.getRow(r),tr=document.createElement('tr'),rh=document.createElement('th');rh.className='excel-source-row-heading';rh.textContent=String(r);tr.appendChild(rh);
      for(let c=1;c<=maxCol;c++){if(covered.has(r+':'+c))continue;const cell=row.getCell(c),td=document.createElement('td'),merge=merges.get(r+':'+c),style=cell.style||{};
        if(merge){td.rowSpan=merge.rowSpan;td.colSpan=merge.colSpan;}td.textContent=excelCellText(cell.value);
        if(style.font){if(style.font.bold)td.style.fontWeight='700';if(style.font.italic)td.style.fontStyle='italic';if(style.font.size)td.style.fontSize=Math.min(28,Math.max(8,style.font.size))+'px';if(style.font.name)td.style.fontFamily='"'+String(style.font.name).replace(/"/g,'')+'",Arial,sans-serif';const fc=colorCss(style.font.color);if(fc)td.style.color=fc;}
        if(style.fill&&style.fill.type==='pattern'){const fc=colorCss(style.fill.fgColor);if(fc)td.style.backgroundColor=fc;}
        if(style.alignment){td.style.textAlign=style.alignment.horizontal||'left';td.style.verticalAlign=style.alignment.vertical==='middle'?'middle':style.alignment.vertical||'middle';if(style.alignment.wrapText)td.style.whiteSpace='normal';if(style.alignment.indent)td.style.paddingLeft=(8+style.alignment.indent*8)+'px';}
        if(style.border)for(const side of ['top','right','bottom','left']){const b=style.border[side];if(b&&b.style)td.style['border'+side[0].toUpperCase()+side.slice(1)]='1px solid '+(colorCss(b.color)||'#aeb9c7');}
        tr.appendChild(td);
      }if(row.height)tr.style.height=Math.max(16,Math.round(row.height*4/3))+'px';body.appendChild(tr);
    }
    table.appendChild(body);const wrap=document.createElement('div');wrap.className='movie-excel-scroll excel-source-scroll';wrap.appendChild(table);preview.appendChild(wrap);
    const note=document.createElement('p');note.className='excel-preview-note';note.textContent='Original worksheet preview · '+sheet.name+' · '+maxRow+' row(s) × '+maxCol+' column(s). Source values, used range, merged cells, dimensions and supported cell formatting are rendered from the uploaded workbook.';preview.appendChild(note);
  }

  // Delete controls are delegated because the preview table is rebuilt after every change.
  preview.addEventListener('click',event=>{
    const button=event.target.closest('button[data-delete-movie-id]');
    if(!button||!preview.contains(button))return;
    const movieId=button.dataset.deleteMovieId;
    const movie=getMovies().find(item=>item.id===movieId);
    if(!movie)return;
    if(!window.confirm('Delete "'+movie.name+'" from the '+movie.language+' movie collection? This will also update the Excel worksheet and poster counter.'))return;
    button.disabled=true;
    if(!deleteMovie(movieId)){
      button.disabled=false;
      showMessage('The movie could not be deleted. Please try again.');
      return;
    }
    showMessage('"'+movie.name+'" deleted. The alphabet group, poster counter and following worksheet sections are being updated.');
  });

  function renderSelectedPreview(){
    if(!workbook||!currentPreviewSheet)throw new Error('Select a worksheet to preview.');
    const sheet=workbook.getWorksheet(currentPreviewSheet);
    if(!sheet)throw new Error('The selected worksheet is no longer available. Please select it again.');
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



  const readButton=document.getElementById('readUploadedFile');
  const progressWrap=document.getElementById('readingProgressWrap');
  const progressBar=document.getElementById('readingProgressBar');
  const progressText=document.getElementById('readingProgressText');
  const progressTrack=document.getElementById('readingProgressTrack');
  const progressDetail=document.getElementById('readingProgressDetail');
  let reading=false;
  const nextPaint=()=>new Promise(resolve=>requestAnimationFrame(resolve));
  async function setReadingProgress(percent,detail){
    const value=Math.max(0,Math.min(100,Math.round(percent)));
    if(progressBar)progressBar.style.width=value+'%';
    if(progressText)progressText.textContent=value+'%';
    if(progressTrack)progressTrack.setAttribute('aria-valuenow',String(value));
    if(progressDetail&&detail)progressDetail.textContent=detail;
    await nextPaint();
  }
  fileInput.addEventListener('change',()=>{
    uploadAttempt++;selectedFile=fileInput.files&&fileInput.files[0]||null;
    workbook=null;currentPreviewSheet='';sheetSelectWrap.hidden=true;previewWrap.hidden=true;
    if(completeStage)completeStage.hidden=true;if(progressWrap)progressWrap.hidden=true;
    if(readButton)readButton.disabled=!selectedFile;
    if(!selectedFile){fileName.textContent='Choose Excel workbook';fileHint.textContent='Click to browse · Upload required each time';showMessage('Upload a reference workbook to begin.');return;}
    if(!/\.xlsx$/i.test(selectedFile.name)){selectedFile=null;fileInput.value='';fileName.textContent='Choose Excel workbook';fileHint.textContent='Please select an .xlsx file';if(readButton)readButton.disabled=true;showMessage('This workflow requires an .xlsx workbook. Save legacy .xls files as .xlsx and upload again.');return;}
    fileName.textContent=selectedFile.name;fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · Ready to read';
    showMessage('File selected. Click “Read the Uploaded File” to process every worksheet.');
  });
  readButton?.addEventListener('click',async()=>{
    if(!selectedFile||reading)return;const attempt=uploadAttempt;
    reading=true;readButton.disabled=true;readButton.textContent='Reading Workbook…';
    sheetSelectWrap.hidden=true;previewWrap.hidden=true;if(completeStage)completeStage.hidden=true;
    if(progressWrap)progressWrap.hidden=false;await setReadingProgress(0,'Preparing Excel workbook reader…');
    try{
      const ExcelJS=await loadExcelJS();if(attempt!==uploadAttempt)return;
      await setReadingProgress(5,'Loading the uploaded .xlsx file…');
      const parsed=new ExcelJS.Workbook();await parsed.xlsx.load(await selectedFile.arrayBuffer());
      if(attempt!==uploadAttempt)return;await setReadingProgress(15,'Workbook opened. Reading worksheet contents…');
      const sheets=parsed.worksheets||[];if(!sheets.length)throw new Error('The uploaded workbook contains no readable worksheets.');
      const totalRows=sheets.reduce((n,sh)=>n+Math.max(1,sh.rowCount||1),0);let done=0;
      for(let si=0;si<sheets.length;si++){
        const sh=sheets[si],rows=Math.max(1,sh.rowCount||1),cols=Math.max(1,sh.columnCount||1);
        for(let rowIndex=1;rowIndex<=rows;rowIndex++){
          const row=sh.getRow(rowIndex);for(let colIndex=1;colIndex<=cols;colIndex++)void row.getCell(colIndex).value;
          done++;if(rowIndex===rows||rowIndex%10===0)await setReadingProgress(15+done/totalRows*80,'Reading worksheet '+(si+1)+' of '+sheets.length+': '+sh.name);
          if(attempt!==uploadAttempt)return;
        }
      }
      workbook=parsed;currentPreviewSheet='';refreshSheetOptions();sheetSelectWrap.hidden=false;
      await setReadingProgress(100,'Reading completed successfully. '+sheets.length+' worksheet(s) are ready to preview.');
      await persistCurrentWorkbook();fileHint.textContent=(selectedFile.size/1024/1024).toFixed(2)+' MB · Reading completed';
      showMessage('Reading completed successfully. Choose a worksheet in Proceed to Preview.');
      sheetSelectWrap.scrollIntoView({behavior:'smooth',block:'nearest'});
    }catch(error){
      if(attempt!==uploadAttempt)return;workbook=null;sheetSelectWrap.hidden=true;
      await setReadingProgress(0,'Reading failed. '+(error.message||'Unable to read this workbook.'));
      showMessage('Excel reading failed: '+(error.message||'Unable to read this workbook.'));
    }finally{if(attempt===uploadAttempt){reading=false;readButton.disabled=!selectedFile;readButton.textContent='Read the Uploaded File';}}
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
  subscribeToMovies(()=>{ /* Keep the uploaded source workbook unchanged until export. */ });

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
      link.href=url;
       const exportNow=new Date();
       const dateStamp=[exportNow.getFullYear(),String(exportNow.getMonth()+1).padStart(2,'0'),String(exportNow.getDate()).padStart(2,'0')].join('-');
       const timeStamp=[String(exportNow.getHours()).padStart(2,'0'),String(exportNow.getMinutes()).padStart(2,'0'),String(exportNow.getSeconds()).padStart(2,'0')].join('-');
       link.download='Chayan Watchlist_('+dateStamp+')_('+timeStamp+').xlsx';
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
