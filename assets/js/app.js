import {initInputSection} from './input-section.js';
import {initExcelInterface} from './excel-interface.js?v=20261106-movie-entry-gated-v1';
initInputSection();
initExcelInterface();

const workspace=document.getElementById('mainWorkspace');
const sidebarToggle=document.getElementById('toggleMovieSidebar');
const sidebarLabel=sidebarToggle?.querySelector('.toggle-label');
sidebarToggle?.addEventListener('click',()=>{
  const collapsed=workspace.classList.toggle('sidebar-collapsed');
  sidebarToggle.setAttribute('aria-expanded',String(!collapsed));
  sidebarToggle.setAttribute('aria-label',collapsed?'Show movie details sidebar':'Hide movie details sidebar');
  if(sidebarLabel)sidebarLabel.textContent=collapsed?'Show movie details':'Hide movie details';
  const icon=sidebarToggle.querySelector('.toggle-icon');
  if(icon)icon.textContent=collapsed?'›':'‹';
});

const movieEntryPanel=document.getElementById('movieDetailsPanel');
function setMovieEntryAvailability(enabled){
  if(!movieEntryPanel)return;
  movieEntryPanel.inert=!enabled;
  movieEntryPanel.classList.toggle('movie-entry-locked',!enabled);
  movieEntryPanel.setAttribute('aria-disabled',String(!enabled));
}
setMovieEntryAvailability(false);
document.addEventListener('movie-entry-availability',event=>{
  setMovieEntryAvailability(Boolean(event.detail&&event.detail.enabled));
});
