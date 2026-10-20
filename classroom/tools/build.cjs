const fs=require('node:fs'),path=require('node:path');const root=path.resolve(__dirname,'..'),game=path.join(root,'../city-stage6');
let html=fs.readFileSync(path.join(game,'index.html'),'utf8');html=html.replace(/<link[^>]*href="src\/style.css"[^>]*>/,()=>'<style>'+fs.readFileSync(path.join(game,'src/style.css'),'utf8')+'</style>');
for(const name of ['engine.js','city-map.js','settings.js','app.js']){let s=fs.readFileSync(path.join(game,'src',name),'utf8');if(name==='app.js'){
 const replace=(a,b)=>{if(!s.includes(a))throw Error('Game bridge anchor changed: '+a);s=s.replace(a,b)};
 replace("try{const raw=localStorage.getItem(KEY);", "try{const raw=null;");
 replace('function save(){try{','function save(){if(window.ClassroomBridge){window.ClassroomBridge.submit(state);return true}try{');
 replace('const a=el.dataset.action;if(busy||!el.isConnected)return;',`const a=el.dataset.action;if(busy||!el.isConnected)return;if(['start','new','retry','import','confirm-import'].includes(a)){toast('수업 모드에서는 교사가 정한 도시를 운영합니다.');return}if(['apply','toggle-confirm','next','resolve'].includes(a)&&!window.ClassroomBridge.allowed()){toast('수업 시작 또는 서버 연결을 기다려 주세요.');return}`);
 replace('window.DaonGame={getState:',`window.DaonGame={setClassroomState:raw=>{const restored=E.restore(raw);if(!restored)throw Error('Invalid server state');if(JSON.stringify(state)===JSON.stringify(restored))return;state=restored;close();render();paintTour()},getState:`);
 }
 html=html.replace('<script src="src/'+name+'"></script>',()=>'<script>'+s.replace(/<\/script/gi,'<\\/script')+'</script>');}
html=html.replace('브라우저의 이 사이트 저장 공간을 사용합니다.', '수업 중 선택은 서버에 저장됩니다. 이 메뉴의 JSON은 개인 보관용이며 수업방에 다시 불러올 수 없습니다.').replace('계정 동기화는 없으며 브라우저 데이터 삭제·다른 경로로 파일 열기·비공개 창 종료 시 이어하기가 제한될 수 있습니다.', '브라우저 데이터 삭제 시 조 참가 권한을 잃을 수 있습니다. 서버 기록은 7일간 보관됩니다.').replace('현재 진행 저장 완료 · 매 선택과 결산 후 자동 저장됩니다.', '수업 서버와 동기화 요청 · 상단의 연결 및 기록 상태를 확인하세요.');
const bridge=`<script>
(()=>{let enabled=false,pending=false,acknowledgedActions='[]';window.ClassroomBridge={allowed:()=>enabled&&!pending,submit:s=>{if(JSON.stringify(s.actions)===acknowledgedActions)return;pending=true;parent.postMessage({kind:'daon-choice',actions:s.actions},location.origin)}};window.addEventListener('message',e=>{if(e.origin!==location.origin||e.source!==parent)return;const m=e.data;if(m.kind==='daon-session'){enabled=m.enabled;pending=false;if(m.state){acknowledgedActions=JSON.stringify(m.state.actions);DaonGame.setClassroomState(m.state)}}if(m.kind==='daon-lock'){enabled=false;pending=false}});window.addEventListener('load',()=>parent.postMessage({kind:'daon-ready'},location.origin));})();
</script>`;
html=html.replace('</head>','<style>[data-action=new],[data-action=retry],[data-action=import],[data-action=confirm-import]{display:none!important}</style></head>').replace('지역별 변화와 최종 결말을 확인하고 다른 전략에 도전하세요.', '지역별 변화와 최종 결말을 확인하고 다른 조와 비교하세요.');
html=html.replace('<script>',bridge+'<script>');fs.writeFileSync(path.join(root,'public/game.html'),html);console.log('Classroom game built');
