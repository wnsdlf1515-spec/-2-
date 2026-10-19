/* Original Web Audio synthesis. No recordings, external music or sound assets. */
(function(root){'use strict';
const KEY='daon-stage6-preferences-v1',defaults={version:1,muted:false,volume:.25,reducedMotion:!!root.matchMedia?.('(prefers-reduced-motion: reduce)').matches,skipEffects:false,tutorialDone:false,tutorialStep:'region',tutorialProject:null};
let prefs={...defaults},persisted=true,ctx=null,master=null,voices=new Set(),played=0,blocked=0,unlocked=false;
function clean(o){const p={...defaults};for(const k of ['muted','reducedMotion','skipEffects','tutorialDone'])if(typeof o?.[k]==='boolean')p[k]=o[k];if(Number.isFinite(o?.volume))p.volume=Math.max(0,Math.min(1,o.volume));if(['region','policy','confirm','next','result','done'].includes(o?.tutorialStep))p.tutorialStep=o.tutorialStep;if(o?.tutorialProject===null||typeof o?.tutorialProject==='string'&&o.tutorialProject.length<80)p.tutorialProject=o.tutorialProject;return p;}
try{const raw=root.localStorage.getItem(KEY);if(raw)prefs=clean(JSON.parse(raw));}catch{persisted=false;}
function update(patch){prefs=clean({...prefs,...patch});try{root.localStorage.setItem(KEY,JSON.stringify(prefs));persisted=true;}catch{persisted=false;}if(master&&ctx)master.gain.setValueAtTime(prefs.muted?0:prefs.volume,ctx.currentTime);if(prefs.muted)for(const v of voices){try{v.stop();}catch{}}return get();}
function get(){return {...prefs,persisted};}
function unlock(){// Called only by UI input handlers; no autoplay on page load.
 if(root.navigator?.userActivation&&!root.navigator.userActivation.isActive&&!unlocked)return false;
 try{const AC=root.AudioContext||root.webkitAudioContext;if(!AC)return false;if(!ctx){ctx=new AC();master=ctx.createGain();master.gain.value=prefs.muted?0:prefs.volume;master.connect(ctx.destination);}if(ctx.state==='suspended')ctx.resume().catch(()=>{});unlocked=true;return true;}catch{return false;}}
const cues={select:[[440,0,.04]],confirm:[[392,0,.07],[523.25,.06,.09]],complete:[[392,0,.09],[493.88,.08,.09],[587.33,.16,.12]],warning:[[293.66,0,.08],[261.63,.12,.1]]};
function play(name){if(prefs.muted||prefs.volume<=0||!unlocked||!ctx||ctx.state!=='running'||root.document.hidden){blocked++;return false;}const now=ctx.currentTime;for(const[f,delay,duration]of cues[name]||cues.select){const o=ctx.createOscillator(),g=ctx.createGain();o.type='triangle';o.frequency.value=f;g.gain.setValueAtTime(0,now+delay);g.gain.linearRampToValueAtTime(.16,now+delay+.007);g.gain.exponentialRampToValueAtTime(.001,now+delay+duration);o.connect(g);g.connect(master);voices.add(o);o.onended=()=>{voices.delete(o);o.disconnect();g.disconnect();};o.start(now+delay);o.stop(now+delay+duration+.01);}played++;return true;}
root.document.addEventListener('visibilitychange',()=>{if(root.document.hidden){for(const v of voices){try{v.stop();}catch{}}}});
root.CitySettings=Object.freeze({get,update,unlock,play,key:KEY,getSoundStats:()=>({created:!!ctx,unlocked,state:ctx?.state||'not-created',played,blocked,activeVoices:voices.size,muted:prefs.muted,volume:prefs.volume})});
})(window);
