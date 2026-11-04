const STORAGE_KEY='chayan-watchlist-v1';
const LANGUAGES=['Bengali','Hindi','English'];
let movies=loadMovies();
const subscribers=new Set();
function loadMovies(){
  try{
    const value=JSON.parse(localStorage.getItem(STORAGE_KEY)||'[]');
    return Array.isArray(value)?value.map((m,i)=>({
      id:String(m.id||('legacy-'+i)),
      language:LANGUAGES.includes(m.language)?m.language:(LANGUAGES.find(l=>String(m.language||'').toLowerCase().includes(l.toLowerCase()))||'English'),
      name:String(m.name||m.title||m.movieName||'Untitled movie'),
      year:Number(m.year||m.releaseYear||m.movieYear||'')||'Not specified',
      poster:typeof m.poster==='string'?m.poster:(typeof m.image==='string'?m.image:'')
    })): [];
  }catch(error){console.error('Could not read saved watchlist:',error);return [];}
}
function persist(next){
  try{localStorage.setItem(STORAGE_KEY,JSON.stringify(next));return true;}
  catch(error){console.error('Could not save watchlist:',error);alert('The browser could not save this entry. Storage may be full; try a smaller poster image.');return false;}
}
function publish(){const snapshot=getMovies();subscribers.forEach(fn=>fn(snapshot));}
export function getMovies(){return movies.map(m=>({...m}));}
export function subscribeToMovies(fn){subscribers.add(fn);return()=>subscribers.delete(fn);}
export function addMovie(movie){const next=[{...movie},...movies];if(!persist(next))return false;movies=next;publish();return true;}
export function deleteMovie(id){const next=movies.filter(m=>m.id!==id);if(next.length===movies.length||!persist(next))return false;movies=next;publish();return true;}
window.addEventListener('storage',event=>{if(event.key===STORAGE_KEY){movies=loadMovies();publish();}});
