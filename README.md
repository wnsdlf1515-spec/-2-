# 다온시 — 여덟 해의 선택

3개 초기 도시 조건에서 주거·일자리·생활서비스를 운영하는 8턴 전략 시뮬레이션입니다. 교육용 모형이며 실제 정책 효과를 예측하지 않습니다.

## 실행 파일

- [게임 HTML 다운로드](https://github.com/wnsdlf1515-spec/-2-/raw/refs/heads/main/index.html)
- [소스·검증 기록 ZIP](https://github.com/wnsdlf1515-spec/-2-/raw/refs/heads/main/city-stage6/release/daon-city-stage6-rc1.zip)
- [조작·저장 안내](city-stage6/README.md)
- [실제 검증 범위](city-stage6/docs/validation.md)

## 웹 배포

GitHub 저장소 Settings → Pages → Build and deployment에서 Source를 Deploy from a branch, Branch를 main, 폴더를 /(root)로 지정하고 Save합니다. 루트 index.html과 .nojekyll은 준비되어 있습니다. 배포 완료 후 Pages 설정에 표시되는 주소로 실행하세요. 예상 주소는 https://wnsdlf1515-spec.github.io/-2-/ 이지만 활성화/접속 확인 전에는 작동하는 링크로 간주하지 않습니다.

현재 연결에서는 Pages API 설정 권한이 거부되어 호스팅 활성화를 확인하지 못했습니다. 파일 게시와 호스팅 활성화는 구분됩니다.

## 저장과 업데이트

브라우저의 사이트별 localStorage에 저장하며 계정 동기화가 없습니다. 기존 로컬 파일이나 다른 사이트의 저장은 웹 주소로 자동 이전되지 않으므로 JSON을 내보낸 뒤 새 사이트에서 불러오세요. 5·6단계 호환 형식만 지원하며 다른 구버전은 해당 실행 파일을 사용하세요. 손상된 기록을 자동 삭제하지 않습니다.

도트는 절차적 래스터 자산이고 효과음은 직접 합성합니다. 외부 이미지·음원·CDN·개인정보 수집 서비스가 없습니다. 실제 스피커 청취, Firefox/Safari, 장시간 학교 PC 시연은 남아 있습니다.

## 조별 도시 + 교사 실시간 집계 (내부망 수업)

- [수업 서버 ZIP 다운로드](https://github.com/wnsdlf1515-spec/-2-/raw/refs/heads/main/classroom/release/daon-classroom.zip)
- [교사 PC 시작 방법](classroom/README.md)
- [실제 교사 집계 화면](classroom/artifacts/teacher-final.png)

교사 PC에 Node.js24 이상을 설치하고 ZIP 전체를 푼 뒤 `classroom/수업서버-시작.cmd`를 실행합니다. 표시되는 내부망 주소와 수업 참가 코드를 학생들에게 공유합니다. 20명은2인1조/10개 도시를 권장합니다. 학생 기기에는 설치가 필요 없습니다. 기존 GitHub Pages 개인용 게임은 그대로이며 실시간 수업은 교사 PC가 제공하는 별도 주소에서 실행됩니다.
