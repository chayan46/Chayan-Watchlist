import {addMovie,getMovies} from './data-management.js';
export function initInputSection(){
  const form=document.getElementById('movieForm');
  const resetButton=document.getElementById('resetMovieInput');
  const fileInput=document.getElementById('posterFile');
  const preview=document.getElementById('posterPreview');
  const uploadTitle=document.getElementById('uploadTitle');
  const uploadHint=document.getElementById('uploadHint');
  const posterMethodTabs=[...document.querySelectorAll('[data-poster-method]')];
  const posterMethodPanels=[...document.querySelectorAll('[data-poster-panel]')];
  const pasteArea=document.getElementById('posterPasteArea');
  const pasteButton=document.getElementById('pastePosterButton');
  const internetSearch=document.getElementById('posterInternetSearch');
  const internetSearchButton=document.getElementById('searchPosterInternet');
  const searchResults=document.getElementById('posterSearchResults');
  const searchStatus=document.getElementById('posterSearchStatus');
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
  let posterData='';
  const normalizedLetter=()=>franchiseLetterInput.value.trim().toLocaleUpperCase().charAt(0);
  function franchiseGroups(letter){
    const relevant=getMovies().filter(movie=>(movie.name.trim().charAt(0)||'').toLocaleUpperCase()===letter&&movie.movieType==='franchise'&&movie.franchiseName);
    const byName=new Map();
    relevant.forEach(movie=>{
      const name=movie.franchiseName.trim();
      const key=name.toLocaleLowerCase();
      if(!byName.has(key))byName.set(key,{name,code:movie.franchiseCode||''});
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
  function showPoster(data,label='Poster selected'){posterData=data;preview.src=data;preview.style.display='block';uploadTitle.textContent=label;uploadHint.textContent='Poster ready to add';}
  function usePosterBlob(blob,label='Pasted poster'){if(!blob||!blob.type?.startsWith('image/'))return;if(blob.size>8*1024*1024){alert('Poster must be 8 MB or smaller.');return;}const reader=new FileReader();reader.onload=()=>showPoster(String(reader.result||''),label);reader.onerror=()=>alert('Could not read this poster image.');reader.readAsDataURL(blob);}
  function resetPoster(){posterData='';preview.removeAttribute('src');preview.style.display='none';uploadTitle.textContent='＋ Upload movie poster';uploadHint.textContent='JPG, PNG, WEBP, GIF · Max 8 MB';if(searchResults)searchResults.replaceChildren();}
  posterMethodTabs.forEach(tab=>tab.addEventListener('click',()=>{const method=tab.dataset.posterMethod;posterMethodTabs.forEach(item=>{const active=item===tab;item.classList.toggle('active',active);item.setAttribute('aria-selected',String(active));});posterMethodPanels.forEach(panel=>panel.hidden=panel.dataset.posterPanel!==method);}));
  pasteArea?.addEventListener('paste',event=>{const item=[...(event.clipboardData?.items||[])].find(entry=>entry.type.startsWith('image/'));if(item){event.preventDefault();usePosterBlob(item.getAsFile());}});
  pasteButton?.addEventListener('click',async()=>{try{if(!navigator.clipboard?.read)throw new Error('Clipboard image access is not supported here. Click the paste area and press Ctrl + V instead.');const items=await navigator.clipboard.read();for(const item of items){const type=item.types.find(value=>value.startsWith('image/'));if(type){usePosterBlob(await item.getType(type));return;}}throw new Error('No image was found in the clipboard.');}catch(error){alert(error.message||'Clipboard access was blocked. Click the paste area and press Ctrl + V instead.');}});
  async function searchInternetPosters(){const query=internetSearch?.value.trim();if(!query){searchStatus.textContent='Enter a poster name or movie title to search.';return;}searchStatus.textContent='Searching poster images…';searchResults.replaceChildren();const params=new URLSearchParams({action:'query',format:'json',origin:'*',generator:'search',gsrsearch:'filetype:bitmap '+query+' movie poster',gsrnamespace:'6',gsrlimit:'18',prop:'imageinfo',iiprop:'url',iiurlwidth:'240'});try{const response=await fetch('https://commons.wikimedia.org/w/api.php?'+params.toString());if(!response.ok)throw new Error('Poster search is temporarily unavailable.');const data=await response.json();const pages=Object.values(data.query?.pages||{}).filter(page=>page.imageinfo?.[0]?.thumburl||page.imageinfo?.[0]?.url);if(!pages.length){searchStatus.textContent='No poster images found. Try another search phrase.';return;}searchStatus.textContent=pages.length+' image results. Select one to use it as the movie poster.';pages.forEach(page=>{const info=page.imageinfo[0],src=info.thumburl||info.url;const button=document.createElement('button');button.type='button';button.className='poster-result';button.title='Use '+page.title.replace(/^File:/,'');const image=document.createElement('img');image.src=src;image.alt=page.title.replace(/^File:/,'');image.loading='lazy';const caption=document.createElement('span');caption.textContent=page.title.replace(/^File:/,'');button.append(image,caption);button.addEventListener('click',()=>{showPoster(info.url||src,page.title.replace(/^File:/,''));searchStatus.textContent='Poster selected and loaded into the preview.';});searchResults.append(button);});}catch(error){searchStatus.textContent=error.message||'Could not search poster images. Please try again.';}}
  internetSearchButton?.addEventListener('click',searchInternetPosters);
  internetSearch?.addEventListener('keydown',event=>{if(event.key==='Enter'){event.preventDefault();searchInternetPosters();}});
  resetButton?.addEventListener('click',()=>{
    form.reset();
    fileInput.value='';
    resetPoster();
    syncFranchiseFields();
    form.querySelectorAll(':invalid').forEach(field=>field.setCustomValidity(''));
  });
  fileInput.addEventListener('change',event=>{
    const file=event.target.files&&event.target.files[0];resetPoster();if(!file)return;
    if(!file.type.startsWith('image/')){alert('Please choose a valid image file.');event.target.value='';return;}
    if(file.size>8*1024*1024){alert('Poster must be 8 MB or smaller.');event.target.value='';return;}
    const reader=new FileReader();
    reader.onload=()=>showPoster(String(reader.result||''),file.name);
    reader.onerror=()=>alert('Could not read this poster. Please choose another image.');
    reader.readAsDataURL(file);
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
