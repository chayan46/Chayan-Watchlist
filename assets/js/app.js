import {initInputSection} from './input-section.js';
import {initExcelInterface} from './excel-interface.js?v=20261106-excel-remove-entry-v1';
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
