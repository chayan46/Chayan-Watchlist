import {getMovies,subscribeToMovies,deleteMovie} from './data-management.js';
const LANGUAGES=['Bengali','Hindi','English'];
const $=id=>document.getElementById(id);
function node(tag,cls,text){const n=document.createElement(tag);if(cls)n.className=cls;if(text!==undefined)n.textContent=text;return n;}
function render(movies){
  const counts=Object.fromEntries(LANGUAGES.map(l=>[l,movies.filter(m=>m.language===l).length]));
  $('total').textContent=String(movies.length);$('bnCard').textContent=String(counts.Bengali);$('hiCard').textContent=String(counts.Hindi);$('enCard').textContent=String(counts.English);
  const query=$('search').value.trim().toLocaleLowerCase();const fragment=document.createDocumentFragment();
  LANGUAGES.forEach((language,index)=>{
    const section=node('section','section'),heading=node('div','sectionhead'),title=node('div','sectiontitle'),dot=node('span','dot');
    dot.style.background=['#168b7e','#c68a36','#536bd0'][index];
    title.append(dot,node('h3','',language+' Movies'),node('span','count',counts[language]+' TOTAL'));heading.append(title);
    const list=node('div','movies');
    const matching=movies.filter(m=>m.language===language&&(!query||(m.name+' '+m.year+' '+m.language).toLocaleLowerCase().includes(query)));
    if(!matching.length)list.append(node('div','empty',query?'No matching '+language.toLowerCase()+' movies.':'No '+language.toLowerCase()+' movies added yet. Add one using the form.'));
    else matching.forEach(movie=>{
      const record=node('article','movie text-record'),remove=node('button','delete','×');remove.type='button';remove.dataset.remove=movie.id;remove.setAttribute('aria-label','Delete '+movie.name);remove.title='Remove movie';
      const body=node('div','moviebody'),serial=node('span','record-number',String(movies.indexOf(movie)+1).padStart(2,'0')),details=node('div','record-content');
      details.append(node('span','detail-label','Movie title'),node('h4','',movie.name));
      const meta=node('p','record-meta');meta.append(node('span','','Release year: '+String(movie.year)),node('span','record-separator','|'),node('span','','Language: '+movie.language));
      details.append(meta);body.append(serial,details);record.append(remove,body);list.append(record);
    });
    section.append(heading,list);fragment.append(section);
  });
  $('sections').replaceChildren(fragment);
}
export function initReadingViewSection(){
  const search=$('search');search.addEventListener('input',()=>render(getMovies()));
  $('clearSearch').addEventListener('click',()=>{search.value='';render(getMovies());});
  $('sections').addEventListener('click',event=>{const button=event.target.closest('[data-remove]');if(!button)return;if(!confirm('Remove this movie from your watchlist?'))return;deleteMovie(button.dataset.remove);});
  subscribeToMovies(render);render(getMovies());
}
