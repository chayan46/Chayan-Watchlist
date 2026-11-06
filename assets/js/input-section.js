import {addMovie,getMovies} from './data-management.js';
export function initInputSection(){
  const form=document.getElementById('movieForm');
  const fileInput=document.getElementById('posterFile');
  const preview=document.getElementById('posterPreview');
  const uploadTitle=document.getElementById('uploadTitle');
  const uploadHint=document.getElementById('uploadHint');
  const movieNameInput=document.getElementById('movieName');
  const movieType=document.getElementById('movieType');
  const franchiseAlphabetField=document.getElementById('franchiseAlphabetField');
  const franchiseAlphabet=document.getElementById('franchiseAlphabet');
  const franchiseAlphabetHint=document.getElementById('franchiseAlphabetHint');
  const newFranchiseNameField=document.getElementById('newFranchiseNameField');
  const newFranchiseName=document.getElementById('newFranchiseName');
  let posterData='';
  const firstLetter=()=>((movieNameInput.value.trim().charAt(0)||'').toLocaleUpperCase());
  function franchiseGroups(letter){
    const relevant=getMovies().filter(movie=>(movie.name.trim().charAt(0)||'').toLocaleUpperCase()===letter&&movie.movieType==='franchise'&&movie.franchiseName);
    const byName=new Map();
    relevant.forEach(movie=>{
      const name=movie.franchiseName.trim();
      const key=name.toLocaleLowerCase();
      if(!byName.has(key))byName.set(key,{name,code:movie.franchiseCode||''});
    });
    const groups=[...byName.values()].sort((a,b)=>{
      const na=Number((a.code.match(/\d+$/)||[])[0])||9999;
      const nb=Number((b.code.match(/\d+$/)||[])[0])||9999;
      return na-nb||a.name.localeCompare(b.name,undefined,{sensitivity:'base'});
    });
    let next=groups.reduce((max,group)=>Math.max(max,Number((group.code.match(/^[A-Z](\d+)$/)||[])[1])||0),0)+1;
    groups.forEach(group=>{if(!group.code)group.code=letter+(groups.indexOf(group)+1);});
    return {groups,next:Math.max(next,groups.length+1)};
  }
  function refreshFranchiseOptions(){
    if(movieType.value!=='franchise')return;
    const letter=firstLetter();
    franchiseAlphabet.replaceChildren();
    if(!/^[A-Z]$/.test(letter)){
      franchiseAlphabet.add(new Option('Enter a movie name first',''));
      franchiseAlphabet.disabled=true;
      franchiseAlphabetHint.textContent='The subgroup alphabet is determined by the first letter of the movie title.';
      return;
    }
    franchiseAlphabet.disabled=false;
    const {groups,next}=franchiseGroups(letter);
    franchiseAlphabet.add(new Option('Select an existing franchise or create a new one',''));
    groups.forEach(group=>franchiseAlphabet.add(new Option(group.code+' — '+group.name,'existing|'+group.code)));
    const newCode=letter+next;
    franchiseAlphabet.add(new Option('＋ Create New Franchise Alphabet ('+newCode+')','new|'+newCode));
    franchiseAlphabetHint.textContent=groups.length
      ?'Existing '+letter+' franchise alphabets are listed above. The next new subgroup will be '+newCode+'.'
      :'No '+letter+' franchise alphabet exists yet. Create the first subgroup: '+newCode+'.';
    if(franchiseAlphabet.value&&!Array.from(franchiseAlphabet.options).some(option=>option.value===franchiseAlphabet.value))franchiseAlphabet.value='';
    syncNewFranchiseName();
  }
  function syncNewFranchiseName(){
    const creating=franchiseAlphabet.value.startsWith('new|');
    newFranchiseNameField.hidden=!creating;
    newFranchiseName.required=creating;
    if(!creating)newFranchiseName.value='';
  }
  function syncFranchiseFields(){
    const isFranchise=movieType.value==='franchise';
    franchiseAlphabetField.hidden=!isFranchise;
    if(!isFranchise){
      franchiseAlphabet.value='';
      newFranchiseNameField.hidden=true;
      newFranchiseName.required=false;
      newFranchiseName.value='';
      return;
    }
    refreshFranchiseOptions();
  }
  movieType.addEventListener('change',syncFranchiseFields);
  movieNameInput.addEventListener('input',refreshFranchiseOptions);
  franchiseAlphabet.addEventListener('change',syncNewFranchiseName);
  syncFranchiseFields();
  function resetPoster(){posterData='';preview.removeAttribute('src');preview.style.display='none';uploadTitle.textContent='＋ Upload movie poster';uploadHint.textContent='JPG, PNG, WEBP, GIF · Max 8 MB';}
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
      const selection=franchiseAlphabet.value;
      if(selection.startsWith('existing|')){
        franchiseCode=selection.split('|')[1];
        const match=franchiseGroups((name.charAt(0)||'').toLocaleUpperCase()).groups.find(group=>group.code===franchiseCode);
        franchiseName=match?.name||'';
      }else if(selection.startsWith('new|')){
        franchiseCode=selection.split('|')[1];
        franchiseName=newFranchiseName.value.trim();
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
