import {addMovie,getMovies} from './data-management.js';
export function initInputSection(){
  const form=document.getElementById('movieForm');
  const resetButton=document.getElementById('resetMovieInput');
  const fileInput=document.getElementById('posterFile');
  const preview=document.getElementById('posterPreview');
  const uploadTitle=document.getElementById('uploadTitle');
  const uploadHint=document.getElementById('uploadHint');
  const movieNameInput=document.getElementById('movieName');
  const movieType=document.getElementById('movieType');
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
  function resetPoster(){posterData='';preview.removeAttribute('src');preview.style.display='none';uploadTitle.textContent='＋ Upload movie poster';uploadHint.textContent='JPG, PNG, WEBP, GIF · Max 8 MB';}
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
    reader.onload=()=>{posterData=String(reader.result||'');preview.src=posterData;preview.style.display='block';uploadTitle.textContent=file.name;uploadHint.textContent='Poster ready to add';};
    reader.onerror=()=>alert('Could not read this poster. Please choose another image.');
    reader.readAsDataURL(file);
  });
  form.addEventListener('submit',event=>{
    event.preventDefault();
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
