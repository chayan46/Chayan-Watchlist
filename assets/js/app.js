import {initInputSection} from './input-section.js?v=20261106-paste-only-v1';
import {initExcelInterface} from './excel-interface.js?v=20261106-excel-image-index-fix-v1';
import {initInputSection as initWebseriesInputSection} from './input-section-webseries.js?v=20261106-independent-webseries-v1';
import {initExcelInterface as initWebseriesExcelInterface} from './excel-interface-webseries.js?v=20261106-excel-image-index-fix-v1';
initInputSection();
initExcelInterface();
initWebseriesInputSection();
initWebseriesExcelInterface();

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
const webseriesWorkspace=document.getElementById('webseriesWorkspace');
const webseriesSidebarToggle=document.getElementById('webseriesToggleWebseriesSidebar');
const webseriesSidebarLabel=webseriesSidebarToggle?.querySelector('.toggle-label');
webseriesSidebarToggle?.addEventListener('click',()=>{
  const collapsed=webseriesWorkspace.classList.toggle('sidebar-collapsed');
  webseriesSidebarToggle.setAttribute('aria-expanded',String(!collapsed));
  webseriesSidebarToggle.setAttribute('aria-label',collapsed?'Show webseries details sidebar':'Hide webseries details sidebar');
  if(webseriesSidebarLabel)webseriesSidebarLabel.textContent=collapsed?'Show webseries details':'Hide webseries details';
  const icon=webseriesSidebarToggle.querySelector('.toggle-icon');
  if(icon)icon.textContent=collapsed?'›':'‹';
});


const movieEntryPanel=document.getElementById('movieDetailsPanel');
const webseriesEntryPanel=document.getElementById('webseriesDetailsPanel');
function setMovieEntryAvailability(enabled){
  if(!movieEntryPanel)return;
  movieEntryPanel.inert=!enabled;
  movieEntryPanel.classList.toggle('movie-entry-locked',!enabled);
  movieEntryPanel.setAttribute('aria-disabled',String(!enabled));
}
setMovieEntryAvailability(false);
function setWebseriesEntryAvailability(enabled){if(!webseriesEntryPanel)return;webseriesEntryPanel.inert=!enabled;webseriesEntryPanel.classList.toggle('webseries-entry-locked',!enabled);webseriesEntryPanel.setAttribute('aria-disabled',String(!enabled));}
setWebseriesEntryAvailability(false);
document.addEventListener('webseries-entry-availability',event=>setWebseriesEntryAvailability(Boolean(event.detail&&event.detail.enabled)));
document.addEventListener('movie-entry-availability',event=>{
  setMovieEntryAvailability(Boolean(event.detail&&event.detail.enabled));
});


// Top-level collection tabs keep the existing movie workspace intact.
const collectionTabs=[
  {button:document.getElementById('movieCollectionTab'),panel:document.getElementById('movieCollectionPanel')},
  {button:document.getElementById('webseriesCollectionTab'),panel:document.getElementById('webseriesCollectionPanel')}
];
collectionTabs.forEach((tab,index)=>{
  tab.button?.addEventListener('click',()=>{
    collectionTabs.forEach((item,itemIndex)=>{
      const active=itemIndex===index;
      item.button?.classList.toggle('active',active);
      item.button?.setAttribute('aria-selected',String(active));
      if(item.panel)item.panel.hidden=!active;
    });
  });
});

// A scoped reset carries its collection through the reload; normal visits still default to Movies.
const requestedCollection=new URLSearchParams(window.location.search).get('activeCollection');
const initialCollectionIndex=requestedCollection==='webseries'?1:0;
collectionTabs.forEach((item,index)=>{
  const active=index===initialCollectionIndex;
  item.button?.classList.toggle('active',active);
  item.button?.setAttribute('aria-selected',String(active));
  if(item.panel)item.panel.hidden=!active;
});
