const STORAGE_KEY='chayan-watchlist-webseries-v1';
const LANGUAGES=['Bengali','Hindi','English'];
let webseries=loadWebseries();
const subscribers=new Set();
function loadWebseries(){
  try{
    const value=JSON.parse(localStorage.getItem(STORAGE_KEY)||'[]');
    return Array.isArray(value)?value.map((m,i)=>({
      id:String(m.id||('legacy-'+i)),
      language:LANGUAGES.includes(m.language)?m.language:(LANGUAGES.find(l=>String(m.language||'').toLowerCase().includes(l.toLowerCase()))||'English'),
      name:String(m.name||m.title||m.webseriesName||'Untitled webseries'),
      year:Number(m.year||m.releaseYear||m.webseriesYear||'')||'Not specified',
      poster:typeof m.poster==='string'?m.poster:(typeof m.image==='string'?m.image:''),
      webseriesType:m.webseriesType==='franchise'?'franchise':'non-franchise',
      franchiseName:typeof m.franchiseName==='string'?m.franchiseName.trim():'',
      franchiseCode:typeof m.franchiseCode==='string'?m.franchiseCode.trim().toLocaleUpperCase():''
    })): [];
  }catch(error){console.error('Could not read saved watchlist:',error);return [];}
}
function persist(next){
  try{localStorage.setItem(STORAGE_KEY,JSON.stringify(next));return true;}
  catch(error){console.error('Could not save watchlist:',error);alert('The browser could not save this entry. Storage may be full; try a smaller poster image.');return false;}
}
function publish(){const snapshot=getWebseries();subscribers.forEach(fn=>fn(snapshot));}
export function getWebseries(){return webseries.map(m=>({...m}));}
export function subscribeToWebseries(fn){subscribers.add(fn);return()=>subscribers.delete(fn);}
export function addWebseries(item){const next=[{...item},...webseries];if(!persist(next))return false;webseries=next;publish();return true;}
export function deleteWebseries(id){const next=webseries.filter(m=>m.id!==id);if(next.length===webseries.length||!persist(next))return false;webseries=next;publish();return true;}
window.addEventListener('storage',event=>{if(event.key===STORAGE_KEY){webseries=loadWebseries();publish();}});
