from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import html,json
root=Path(__file__).parent
meta=[('일    시','____. __. __. __교시'),('장    소','__학년 __반 교실'),('단    원','5. 미래와 지속가능한 삶\n인구 문제와 해결 방안'),('대    상','__학년 __반 학생'),('지도교사','____________')]
goals=['인구 변화와 지역별 생활 여건을 근거로 도시의 주요 문제를 설명할 수 있다.','정책 선택 전후의 지표를 비교하여 정책의 효과와 제한 조건을 설명할 수 있다.','인구 증가와 생활 여건 개선을 구분하여 지속가능한 도시 운영 방안을 제안할 수 있다.']
rows=[
['도입\n5분','문제 인식\n학습 목표 확인','• “인구가 줄면 반드시 살기 나빠질까?”를 제시한다.\n• 저출생·고령화·인구 이동과 지역 격차를 연결해 전시 내용을 확인한다.\n• 오늘의 목표와 역할을 안내한다.','• 인구 감소와 삶의 질에 대한 생각을 말한다.\n• 자연적 증감과 사회적 증감을 구분한다.\n• 운영자·분석가 역할을 나눈다.','사전 학습 자료\n도시 화면','인구수만으로 도시의 성패를 판단하지 않도록 한다.'],
['전개\n40분','게임 안내·\n문제 발견\n7분','• 참가 코드로 조별 접속을 안내한다.\n• 같은 도시·목표·시드를 제시하고 지역 선택→정책 검토→시행→결산을 짧게 시연한다.\n• 예산/연간 수지, 공급/실제 이용의 차이를 강조한다.','• 조 이름으로 참가하고 목표를 확인한다.\n• 생활권별 주택·일자리·서비스 문제를 찾는다.\n• 해결할 문제와 첫 정책의 이유를 합의한다.','교사 PC 서버\n조별 기기\n게임 안내표','조당 기기 1대. 초기 조건을 통일하고 최적 정책 순서를 미리 제시하지 않는다.'],
['','정책 선택·\n도시 운영\n23분','• 수업을 시작하고 조별 진행을 살핀다.\n• “효과가 나타나려면 어떤 조건이 필요할까?”를 발문한다.\n• 4턴 무렵 역할 교대를 안내한다.\n• 중요한 선택 2회만 기록하게 한다.','• 비용·기간·유지비를 비교해 정책과 대상을 선택한다.\n• 시간 진행과 사건 대응 후 지도·지표 변화를 확인한다.\n• 전후 수치와 예상과 다른 이유를 기록한다.\n• 4턴 이후 역할을 바꾼다.','게임\n선택 기록 활동지\n실시간 현황','목표는 8턴 운영. 진행이 느린 조는 현재 턴을 기록한다. 정책 수치는 교육용 가정이며 특정 집단을 문제로 단정하지 않는다.'],
['','결과 비교·\n원인 설명\n10분','• 필요 시 전체 일시정지한다.\n• 차트에서 같은 지표·같은 턴의 2~3개 조를 비교한다.\n• “공급 개선 때문인가, 수요 변화 때문인가?”를 묻는다.\n• 인구 감소 중 생활 여건 개선도 성과로 다룬다.','• 초기 조건과 비교 턴을 확인한다.\n• 정책 선택→조건 변화→결과의 순서로 설명한다.\n• 다른 조와 차이를 만든 투자 위치·순서를 찾는다.\n• 개선점과 남은 문제를 함께 말한다.','교사 집계 차트\n비교 활동지','다른 턴의 현재 수치를 단순 비교하지 않는다. 목표 달성·인구 증가량만으로 평가하지 않는다.'],
['정리·\n차시 예고\n5분','개념 정리\n형성 평가\n차시 안내','• 주거·일자리·서비스·재정의 연결을 정리한다.\n• 게임이 충분히 표현하지 못한 현실 조건을 묻는다.\n• 개인 정리 문장을 확인하고 다음 탐구를 안내한다.','• 개선 근거와 정책의 한계를 한 문장씩 작성한다.\n• 일·생활 균형, 돌봄 분담, 세대 간 책임 등 현실 조건을 제시한다.\n• 다음 탐구 질문을 적는다.','활동지\n개별 정리 문항','차시 예고: 실제 지역 인구 자료와 정책 사례를 비교하고 모형의 한계를 보완하는 대안을 탐구한다.']
]
notes='수업 운영 기준: 50분, 학생 약 20명, 2인 1조(약 10개 도시). 수업 전 학생 기기 1대로 접속·정책·차트 갱신을 확인한다. 접속 장애 시 교사 시연 화면과 준비된 비교 자료를 활용한 모둠 토의로 전환한다.'
assessment='관찰·활동지 평가: 문제 발견의 근거, 비용과 효과 조건의 고려, 전후 수치에 근거한 인과 설명, 현실의 지속가능성에 대한 제안. 게임 승패 자체는 평가 점수로 사용하지 않는다.'
check='확인 사항: 단원 내 차시 번호는 수업 진도에 맞게 기입. 첨부 교과서 PDF는 전송 용량 제한으로 원문 미대조이며, 소단원명은 제공 사진의 “인구 문제와 해결 방안”을 본시 주제로 사용함. 단원 성취기준 코드는 확인되지 않아 임의로 기재하지 않음.'
def esc(t):return html.escape(t).replace('\n','<br>')
css='''@page{size:A4;margin:0}*{box-sizing:border-box}body{margin:0;font-family:"Noto Sans CJK KR",sans-serif;color:#111;font-size:10pt;line-height:1.5}.page{width:210mm;height:297mm;padding:15mm 13mm;position:relative;break-after:page}.page:last-child{break-after:auto}.cover{border:1px solid black;height:264mm;margin:0 4mm;padding:19mm 13mm}.cover h1{text-align:center;font-size:28pt;font-weight:500;border:3px double #111;padding:13px;margin:0 0 44mm}.meta{width:100%;border-collapse:collapse;font-size:12pt}.meta th{width:24%;font-weight:500}.meta th,.meta td{border:1px solid #111;padding:10px;text-align:center}.meta td{white-space:normal}h2{text-align:center;font-size:18pt;margin:0 0 10px}table{border-collapse:collapse;width:100%;table-layout:fixed}th,td{border:1px solid #333;padding:5px;vertical-align:top}th{background:#f0f0f0;text-align:center}.intro{font-size:9pt;margin-bottom:8px}.intro th{width:13%}.process{font-size:8.3pt;line-height:1.48}.process td{overflow-wrap:break-word}.process .stage{text-align:center;font-weight:600}.note{font-size:8pt;margin:8px 0;line-height:1.5}.footer{position:absolute;bottom:8mm;left:13mm;right:13mm;text-align:center;font-size:8pt}.confirm{font-size:8pt;color:#444;margin-top:15mm}'''
cover='<section class="page"><div class="cover"><h1>(통합사회2) 공개수업<br>지도안</h1><table class="meta">'+''.join(f'<tr><th>{esc(k)}</th><td>{esc(v)}</td></tr>' for k,v in meta)+'</table></div><div class="footer">1</div></section>'
intro='<h2>본시 교수-학습 과정안</h2><table class="intro"><tr><th>대단원</th><td colspan="3">5. 미래와 지속가능한 삶</td></tr><tr><th>소단원</th><td>인구 문제와 해결 방안<br>(도시 운영 모형으로 정책 탐구)</td><th>차시</th><td>적용·탐구 차시<br>단원 내 순서 확인 필요</td></tr><tr><th>학습 목표</th><td colspan="3">'+''.join(f'{i}. {g}<br>' for i,g in enumerate(goals,1))+'</td></tr></table>'
t='<table class="process"><colgroup><col style="width:9%"><col style="width:12%"><col style="width:26%"><col style="width:26%"><col style="width:10%"><col style="width:17%"></colgroup><thead><tr><th rowspan="2">학습<br>단계</th><th rowspan="2">수업내용</th><th colspan="2">교수학습활동</th><th rowspan="2">자료</th><th rowspan="2">유의점</th></tr><tr><th>교수활동</th><th>학습활동</th></tr></thead><tbody>'
for i,row in enumerate(rows):
 t+='<tr>'
 for j,c in enumerate(row):
  if j==0 and i in [2,3]:continue
  t+=f'<td'+(' rowspan="3"' if i==1 and j==0 else '')+(' class="stage"' if j==0 else '')+'>'+esc(c)+'</td>'
 t+='</tr>'
t+='</tbody></table>'
body='<section class="page">'+intro+t+f'<p class="note"><b>운영·대응</b> {notes}</p><p class="note"><b>평가</b> {assessment}</p><p class="note"><b>작성 확인</b> {check}</p><div class="footer">2</div></section>'
(root/'수업공개_지도안_작성본.html').write_text('<!doctype html><html lang="ko"><meta charset="utf-8"><title>통합사회2 공개수업 지도안</title><style>'+css+'</style><body>'+cover+body+'</body></html>')
# Editable DOCX, same template fields and merged header/phase cells.
doc=Document();sec=doc.sections[0];sec.page_width=Cm(21);sec.page_height=Cm(29.7);sec.top_margin=Cm(1.5);sec.bottom_margin=Cm(1.5);sec.left_margin=sec.right_margin=Cm(1.3)
normal=doc.styles['Normal'];normal.font.name='맑은 고딕';normal.font.size=Pt(9);normal.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'맑은 고딕');normal.paragraph_format.space_after=Pt(2)
p=doc.add_paragraph();p.paragraph_format.space_before=Pt(45);p.paragraph_format.space_after=Pt(110);p.alignment=1;r=p.add_run('(통합사회2) 공개수업\n지도안');r.font.size=Pt(28)
tb=doc.add_table(rows=0,cols=2);tb.style='Table Grid';tb.columns[0].width=Cm(3.3);tb.columns[1].width=Cm(13)
for k,v in meta:
 cells=tb.add_row().cells;cells[0].text=k;cells[1].text=v
 for c in cells:
  for p in c.paragraphs:p.alignment=1;p.paragraph_format.space_before=Pt(8);p.paragraph_format.space_after=Pt(8)
doc.add_page_break();p=doc.add_paragraph('본시 교수-학습 과정안');p.alignment=1;p.runs[0].font.size=Pt(17)
tb=doc.add_table(rows=3,cols=4);tb.style='Table Grid';tb.cell(0,0).text='대단원';tb.cell(0,1).merge(tb.cell(0,3)).text='5. 미래와 지속가능한 삶';tb.cell(1,0).text='소단원';tb.cell(1,1).text='인구 문제와 해결 방안\n(도시 운영 모형으로 정책 탐구)';tb.cell(1,2).text='차시';tb.cell(1,3).text='적용·탐구 차시\n단원 내 순서 확인 필요';tb.cell(2,0).text='학습 목표';tb.cell(2,1).merge(tb.cell(2,3)).text='\n'.join(f'{i}. {g}' for i,g in enumerate(goals,1))
doc.add_paragraph('')
tb=doc.add_table(rows=2,cols=6);tb.style='Table Grid';tb.autofit=False
widths=[1.6,2.2,4.8,4.8,1.8,3.2]
for c,w in zip(tb.columns,widths):c.width=Cm(w)
for j,name in enumerate(['학습 단계','수업내용','교수활동','학습활동','자료','유의점']):tb.cell(1,j).text=name
for j in [0,1,4,5]:tb.cell(0,j).merge(tb.cell(1,j)).text=['학습 단계','수업내용','','','자료','유의점'][j]
tb.cell(0,2).merge(tb.cell(0,3)).text='교수학습활동'
for row in rows:
 cells=tb.add_row().cells
 for c,text in zip(cells,row):
  c.text=text
  for p in c.paragraphs:
   p.paragraph_format.space_after=Pt(0)
   for r in p.runs:r.font.size=Pt(8)
tb.cell(3,0).merge(tb.cell(5,0)).text='전개\n40분'
for row in tb.rows:
 trPr=row._tr.get_or_add_trPr();trPr.append(OxmlElement('w:cantSplit'))
for label,text in [('운영·대응',notes),('평가',assessment),('작성 확인',check)]:
 p=doc.add_paragraph(label+' — '+text)
 for r in p.runs:r.font.size=Pt(8)
doc.save(root/'수업공개_지도안_작성본.docx')
plain='(통합사회2) 공개수업 지도안\n\n'+'\n'.join(k+': '+v for k,v in meta)+'\n\n본시 교수-학습 과정안\n대단원:5. 미래와 지속가능한 삶\n소단원:인구 문제와 해결 방안\n차시:적용·탐구 차시(단원 내 순서 확인 필요)\n\n학습 목표\n'+'\n'.join(goals)+'\n\n'
for row in rows:plain+='\n'.join(k+': '+v for k,v in zip(['학습 단계','수업내용','교수활동','학습활동','자료','유의점'],row))+'\n\n'
plain+='운영·대응: '+notes+'\n평가: '+assessment+'\n작성 확인: '+check
(root/'한글_붙여넣기용_본문.txt').write_text(plain,encoding='utf-8-sig')
(root/'README.txt').write_text('첨부 양식의 항목에 맞춰 작성한 공개 다운로드용 지도안입니다. 일시·장소·대상·지도교사 칸은 직접 기입하세요. 원본 HWP를 직접 수정한 파일은 아닙니다.\nDOCX를 한글에서 열고 표·쪽 나눔을 확인한 뒤 HWP로 다른 이름 저장할 수 있습니다. PDF는 확인·인쇄용이고 TXT는 원본 HWP에 붙여넣기용입니다.\n50분 수업으로 구성했으며 단원 내 차시 번호와 실제 진도만 확인해 주세요. 차시 예고는 후속 탐구 제안이므로 실제 수업 계획에 맞게 조정하세요.\n수업 내용은 유지하고 개인 식별 정보를 빈칸으로 처리했습니다. 교과서 원문 대조는 용량 제한으로 미완료입니다.\n',encoding='utf-8-sig')
print('Created private lesson plan DOCX,HTML,TXT')
