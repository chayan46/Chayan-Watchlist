import {addMovie,getMovies} from './data-management.js';
export function initInputSection(){
  const form=document.getElementById('movieForm');
  const resetButton=document.getElementById('resetMovieInput');
  const preview=document.getElementById('posterPreview');
  const pasteArea=document.getElementById('posterPasteArea');
  const pasteStatus=document.getElementById('posterPasteStatus');
  const movieNameInput=document.getElementById('movieName');
  const movieType=document.getElementById('movieType');
  const languageSelect=document.getElementById('language');
  const addMovieButton=document.getElementById('addMovieButton');
  let selectedWorksheetName='';
  function syncAddMovieAvailability(){
    const expectedWorksheet={Bengali:'Bengali Movies',Hindi:'Hindi Movies',English:'English Movies'}[languageSelect.value]||'';
    const matches=!!expectedWorksheet&&selectedWorksheetName===expectedWorksheet;
    if(addMovieButton)addMovieButton.disabled=!matches;
  }
  languageSelect.addEventListener('change',syncAddMovieAvailability);
  document.addEventListener('selected-worksheet-change',event=>{
    selectedWorksheetName=String(event.detail?.sheetName||'');
    syncAddMovieAvailability();
  });
  syncAddMovieAvailability();
  const franchiseAlphabetField=document.getElementById('franchiseAlphabetField');
  const franchiseLetterInput=document.getElementById('franchiseLetter');
  const franchiseSelectField=document.getElementById('franchiseSelectField');
  const franchiseAlphabet=document.getElementById('franchiseAlphabet');
  const franchiseAlphabetHint=document.getElementById('franchiseAlphabetHint');
  let workbookFranchiseSections=null;
  document.addEventListener('movie-workbook-franchise-sections',event=>{
    workbookFranchiseSections=Array.isArray(event.detail?.sections)?event.detail.sections:null;
    refreshFranchiseOptions();
  });
  let posterData='';
  const normalizedLetter=()=>franchiseLetterInput.value.trim().toLocaleUpperCase().charAt(0);
  function franchiseGroups(letter){
    let groups;
    if(workbookFranchiseSections!==null){
      groups=workbookFranchiseSections
        .filter(group=>String(group.code||'').toLocaleUpperCase().startsWith(letter)&&new RegExp('^'+letter+'\\d+$','i').test(String(group.code||'')))
        .map(group=>({name:String(group.name||group.code),code:String(group.code).toLocaleUpperCase()}));
    }else{
      const relevant=getMovies().filter(item=>(item.name.trim().charAt(0)||'').toLocaleUpperCase()===letter&&item.movieType==='franchise'&&item.franchiseName);
      const byName=new Map();
      relevant.forEach(item=>{
        const name=item.franchiseName.trim();
        const key=name.toLocaleLowerCase();
        if(!byName.has(key))byName.set(key,{name,code:item.franchiseCode||''});
      });
      groups=[...byName.values()];
      const usedCodes=new Set(groups.map(group=>group.code).filter(code=>new RegExp('^'+letter+'\\d+$').test(code)));
      let legacyNumber=1;
      groups.forEach(group=>{
        if(!group.code){while(usedCodes.has(letter+legacyNumber))legacyNumber++;group.code=letter+legacyNumber;usedCodes.add(group.code);legacyNumber++;}
      });
    }
    const byCode=new Map();groups.forEach(group=>{if(!byCode.has(group.code))byCode.set(group.code,group);});
    groups=[...byCode.values()].sort((a,b)=>Number(a.code.slice(1))-Number(b.code.slice(1))||a.name.localeCompare(b.name,undefined,{sensitivity:'base'}));
    const maxNumber=groups.reduce((max,group)=>Math.max(max,Number(group.code.slice(1))||0),0);
    return {groups,next:maxNumber+1};
  }
  function refreshFranchiseOptions(){
    if(movieType.value!=='franchise')return;
    const letter=normalizedLetter();
    franchiseAlphabet.replaceChildren();
    franchiseAlphabet.disabled=true;
    franchiseSelectField.hidden=true;
    if(!/^[A-Z]$/.test(letter)){
      franchiseAlphabet.add(new Option('Enter first letter first',''));
      franchiseAlphabetHint.textContent='Enter a single alphabet from A to Z to load its existing franchise sections.';
      return;
    }
    const {groups,next}=franchiseGroups(letter);
    franchiseSelectField.hidden=false;
    franchiseAlphabet.disabled=false;
    franchiseAlphabet.add(new Option('Select an existing section or create a new one',''));
    groups.forEach(group=>franchiseAlphabet.add(new Option(group.code+' — '+group.name,'existing|'+group.code)));
    const newCode=letter+next;
    franchiseAlphabet.add(new Option('＋ Add New Alphabet Section ('+newCode+')','new|'+newCode));
    franchiseAlphabetHint.textContent=groups.length
      ?'Found '+groups.length+' existing section(s) in the current saved movie collection. Next available section: '+newCode+'.'
      :'No existing franchise section for '+letter+'. The first section will be '+newCode+'.';
    franchiseAlphabet.value='';
  }
  function syncFranchiseFields(){
    const isFranchise=movieType.value==='franchise';
    franchiseAlphabetField.hidden=!isFranchise;
    franchiseSelectField.hidden=true;
    franchiseLetterInput.required=isFranchise;
    franchiseAlphabet.required=isFranchise;
    if(!isFranchise){
      franchiseLetterInput.value='';
      franchiseAlphabet.replaceChildren(new Option('Enter first letter first',''));
      franchiseAlphabet.value='';
      return;
    }
    refreshFranchiseOptions();
  }
  movieType.addEventListener('change',syncFranchiseFields);
  franchiseLetterInput.addEventListener('input',()=>{
    const value=franchiseLetterInput.value.toLocaleUpperCase().replace(/[^A-Z]/g,'').slice(0,1);
    franchiseLetterInput.value=value;
    refreshFranchiseOptions();
  });
  syncFranchiseFields();
  function showPoster(data){posterData=data;preview.src=data;preview.style.display='block';pasteStatus.textContent='Poster pasted successfully. It will be added with the movie.';}
  function usePosterBlob(blob){if(!blob||!blob.type?.startsWith('image/')){pasteStatus.textContent='Clipboard does not contain an image. Copy a poster image and press Ctrl + V here.';return;}if(blob.size>8*1024*1024){pasteStatus.textContent='Poster must be 8 MB or smaller.';return;}const reader=new FileReader();reader.onload=()=>showPoster(String(reader.result||''));reader.onerror=()=>{pasteStatus.textContent='Could not read this image. Please copy and paste it again.';};reader.readAsDataURL(blob);}
  function handlePosterPaste(event){const items=[...(event.clipboardData?.items||[])];const imageItem=items.find(item=>item.type.startsWith('image/'));if(imageItem){event.preventDefault();event.stopPropagation();usePosterBlob(imageItem.getAsFile());return;}const imageFile=[...(event.clipboardData?.files||[])].find(file=>file.type.startsWith('image/'));if(imageFile){event.preventDefault();event.stopPropagation();usePosterBlob(imageFile);return;}pasteStatus.textContent='No image detected. Copy an image first, then press Ctrl + V inside this area.';}
  pasteArea.addEventListener('paste',handlePosterPaste);
  pasteArea.addEventListener('click',()=>{pasteArea.focus();pasteStatus.textContent='Ready. Press Ctrl + V to paste your copied poster image.';});
  document.addEventListener('paste',event=>{if(document.activeElement===pasteArea)handlePosterPaste(event);});
  function resetPoster(){posterData='';preview.removeAttribute('src');preview.style.display='none';pasteStatus.textContent='Only image pasting is supported · Maximum 8 MB';}
  resetButton?.addEventListener('click',()=>{
    form.reset();
    resetPoster();
    syncFranchiseFields();
    form.querySelectorAll(':invalid').forEach(field=>field.setCustomValidity(''));
  });
  form.addEventListener('submit',event=>{
    event.preventDefault();
    syncAddMovieAvailability();
    if(!addMovieButton||addMovieButton.disabled)return;
    const language=document.getElementById('language').value;
    const name=movieNameInput.value.trim();
    const year=Number(document.getElementById('movieYear').value);
    const type=movieType.value;
    let franchiseCode='',franchiseName='';
    if(type==='franchise'){
      const letter=normalizedLetter();
      const selection=franchiseAlphabet.value;
      if(!/^[A-Z]$/.test(letter)||name.charAt(0).toLocaleUpperCase()!==letter){
        alert('The entered franchise alphabet must match the first letter of the movie name.');return;
      }
      if(selection.startsWith('existing|')){
        franchiseCode=selection.split('|')[1];
        const match=franchiseGroups(letter).groups.find(group=>group.code===franchiseCode);
        franchiseName=match?.name||'';
      }else if(selection.startsWith('new|')){
        franchiseCode=selection.split('|')[1];
        franchiseName=franchiseCode;
        const latestNext=franchiseGroups(letter).next;
        if(franchiseCode!==letter+latestNext){
          refreshFranchiseOptions();
          alert('The available franchise sections have changed. Please select the updated section.');
          return;
        }
      }
    }
    if(!['Bengali','Hindi','English'].includes(language)||!name||!Number.isInteger(year)||year<1888||year>2100||!['franchise','non-franchise'].includes(type)||(type==='franchise'&&(!franchiseCode||!franchiseName))){
      alert('Please enter a movie language, title, valid release year and select a valid franchise alphabet.');return;
    }
    const movie={id:Date.now().toString(36)+Math.random().toString(36).slice(2,7),language,name,year,poster:posterData,movieType:type,franchiseName,franchiseCode};
    if(!addMovie(movie))return;
    form.reset();resetPoster();syncFranchiseFields();
    const search=document.getElementById('search');if(search){search.value='';search.dispatchEvent(new Event('input'));}
  });
}
