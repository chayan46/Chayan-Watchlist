import {addMovie} from './data-management.js';
export function initInputSection(){
  const form=document.getElementById('movieForm');
  const fileInput=document.getElementById('posterFile');
  const preview=document.getElementById('posterPreview');
  const uploadTitle=document.getElementById('uploadTitle');
  const uploadHint=document.getElementById('uploadHint');
  const movieType=document.getElementById('movieType');
  const franchiseNameField=document.getElementById('franchiseNameField');
  const franchiseNameInput=document.getElementById('franchiseName');
  let posterData='';
  function syncFranchiseField(){
    const isFranchise=movieType.value==='franchise';
    franchiseNameField.hidden=!isFranchise;
    franchiseNameInput.required=isFranchise;
    if(!isFranchise)franchiseNameInput.value='';
  }
  movieType.addEventListener('change',syncFranchiseField);
  syncFranchiseField();
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
    const name=document.getElementById('movieName').value.trim();
    const year=Number(document.getElementById('movieYear').value);
    const type=movieType.value;
    const franchiseName=type==='franchise'?franchiseNameInput.value.trim():'';
    if(!['Bengali','Hindi','English'].includes(language)||!name||!Number.isInteger(year)||year<1888||year>2100||!['franchise','non-franchise'].includes(type)||(type==='franchise'&&!franchiseName)){
      alert('Please enter a movie language, title, valid release year and the required franchise details.');return;
    }
    const movie={id:Date.now().toString(36)+Math.random().toString(36).slice(2,7),language,name,year,poster:posterData,movieType:type,franchiseName};
    if(!addMovie(movie))return;
    form.reset();resetPoster();syncFranchiseField();
    const search=document.getElementById('search');if(search){search.value='';search.dispatchEvent(new Event('input'));}
  });
}
