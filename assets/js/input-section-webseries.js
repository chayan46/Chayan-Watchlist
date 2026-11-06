import {addWebseries,getWebseries} from './data-management-webseries.js';
export function initInputSection(){
  const form=document.getElementById('webseriesForm');
  const resetButton=document.getElementById('resetWebseriesInput');
  const preview=document.getElementById('posterPreview');
  const pasteArea=document.getElementById('posterPasteArea');
  const pasteStatus=document.getElementById('posterPasteStatus');
  const webseriesNameInput=document.getElementById('webseriesName');
  const webseriesType=document.getElementById('webseriesType');
  const languageSelect=document.getElementById('language');
  const addWebseriesButton=document.getElementById('addWebseriesButton');
  let selectedWorksheetName='';
  function syncAddWebseriesAvailability(){
    const expectedWorksheet={Bengali:'Bengali Webseries',Hindi:'Hindi Webseries',English:'English Webseries'}[languageSelect.value]||'';
    const matches=!!expectedWorksheet&&selectedWorksheetName===expectedWorksheet;
    if(addWebseriesButton)addWebseriesButton.disabled=!matches;
  }
  languageSelect.addEventListener('change',syncAddWebseriesAvailability);
  document.addEventListener('webseries-selected-worksheet-change',event=>{
    selectedWorksheetName=String(event.detail?.sheetName||'');
    syncAddWebseriesAvailability();
  });
  syncAddWebseriesAvailability();
  const franchiseAlphabetField=document.getElementById('franchiseAlphabetField');
  const franchiseLetterInput=document.getElementById('franchiseLetter');
  const franchiseSelectField=document.getElementById('franchiseSelectField');
  const franchiseAlphabet=document.getElementById('franchiseAlphabet');
  const franchiseAlphabetHint=document.getElementById('franchiseAlphabetHint');
  let posterData='';
  const normalizedLetter=()=>franchiseLetterInput.value.trim().toLocaleUpperCase().charAt(0);
  function franchiseGroups(letter){
    const relevant=getWebseries().filter(webseries=>(webseries.name.trim().charAt(0)||'').toLocaleUpperCase()===letter&&webseries.webseriesType==='franchise'&&webseries.franchiseName);
    const byName=new Map();
    relevant.forEach(webseries=>{
      const name=webseries.franchiseName.trim();
      const key=name.toLocaleLowerCase();
      if(!byName.has(key))byName.set(key,{name,code:webseries.franchiseCode||''});
    });
    const groups=[...byName.values()];
    const usedCodes=new Set(groups.map(group=>group.code).filter(code=>new RegExp('^'+letter+'\\d+$').test(code)));
    let legacyNumber=1;
    groups.forEach(group=>{
      if(!group.code){
        while(usedCodes.has(letter+legacyNumber))legacyNumber++;
        group.code=letter+legacyNumber;
        usedCodes.add(group.code);
        legacyNumber++;
      }
    });
    groups.sort((a,b)=>Number(a.code.slice(1))-Number(b.code.slice(1))||a.name.localeCompare(b.name,undefined,{sensitivity:'base'}));
    const maxNumber=groups.reduce((max,group)=>Math.max(max,Number(group.code.slice(1))||0),0);
    return {groups,next:maxNumber+1};
  }
  function refreshFranchiseOptions(){
    if(webseriesType.value!=='franchise')return;
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
      ?'Found '+groups.length+' existing section(s) in the current saved webseries collection. Next available section: '+newCode+'.'
      :'No existing franchise section for '+letter+'. The first section will be '+newCode+'.';
    franchiseAlphabet.value='';
  }
  function syncFranchiseFields(){
    const isFranchise=webseriesType.value==='franchise';
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
  webseriesType.addEventListener('change',syncFranchiseFields);
  franchiseLetterInput.addEventListener('input',()=>{
    const value=franchiseLetterInput.value.toLocaleUpperCase().replace(/[^A-Z]/g,'').slice(0,1);
    franchiseLetterInput.value=value;
    refreshFranchiseOptions();
  });
  syncFranchiseFields();
  function showPoster(data){posterData=data;preview.src=data;preview.style.display='block';pasteStatus.textContent='Poster pasted successfully. It will be added with the webseries.';}
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
    syncAddWebseriesAvailability();
    if(!addWebseriesButton||addWebseriesButton.disabled)return;
    const language=document.getElementById('language').value;
    const name=webseriesNameInput.value.trim();
    const year=Number(document.getElementById('webseriesYear').value);
    const type=webseriesType.value;
    let franchiseCode='',franchiseName='';
    if(type==='franchise'){
      const letter=normalizedLetter();
      const selection=franchiseAlphabet.value;
      if(!/^[A-Z]$/.test(letter)||name.charAt(0).toLocaleUpperCase()!==letter){
        alert('The entered franchise alphabet must match the first letter of the webseries name.');return;
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
      alert('Please enter a webseries language, title, valid release year and select a valid franchise alphabet.');return;
    }
    const webseries={id:Date.now().toString(36)+Math.random().toString(36).slice(2,7),language,name,year,poster:posterData,webseriesType:type,franchiseName,franchiseCode};
    if(!addWebseries(webseries))return;
    form.reset();resetPoster();syncFranchiseFields();
    const search=document.getElementById('search');if(search){search.value='';search.dispatchEvent(new Event('input'));}
  });
}
