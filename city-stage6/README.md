# 다온시 — 6단계 배포 후보 RC1

외부 공개하지 않은 검토용 버전. 5단계의 3개 시나리오·6개 정책·6개 조건부 사건·8년 규칙을 유지하고 안내, 조작, 저장, 연출을 보완했다. engine.js는 5단계와 동일하다.

## 실행

`release/daon-city-stage6-rc1.html`은 CSS·JS·도트 생성 코드가 포함된 단일 HTML이다. 별도 음원·이미지·CDN 요청은 없다. 검증은 로컬 HTTP에서 수행했다. 파일을 직접 여는 방식의 저장 동작은 브라우저별 차이가 있으므로 HTTP 실행을 권장한다.

```
cd /workspace/-2-/city-stage6
npm run build
npm start
```

서버의 `/release/daon-city-stage6-rc1.html`을 브라우저에서 연다. 이 클라우드 채팅에는 외부 접속 가능한 미리보기 URL을 발행하지 않았다. 로컬 서버는 환경 재시작 때 다시 시작해야 한다.

## 처음 플레이

1. 기본 시나리오와 목표를 확인하고 도시 운영 시작.
2. 안내에 따라 기존 주거 생활권을 선택하고 지역 일자리 기반 개선을 검토한다.
3. 투자비·완공 시점·유지비를 확인해 시행한 다음, 다음 해로 진행한다.
4. 결산과 달라진 사업장, 실제 취업 수치를 확인한다. 이후 주거·서비스 병목에 맞춰 투자한다.
5. 도시의 제안에 대응하고 8년을 마친다. 결말에서 같은 조건 재도전 또는 새로운 조건을 선택한다.

안내는 건너뛰거나 도움말에서 다시 볼 수 있다. 별도 보상이나 완화 규칙은 없다. 드래그로 지도 이동, 1×/2×/3× 정수 배율, 시설 클릭 또는 아래 구역 버튼을 사용한다. 설정에서 합성 효과음·음소거·움직임 줄이기·연출 생략을 조정한다.

## 저장

브라우저 localStorage의 `daon-stage5-restored-v1`에 매 선택·결산을 자동 저장한다. 같은 출처의 5단계 데이터와 호환된다. 5단계와 6단계를 같은 출처에서 동시에 열면 동일 저장 슬롯을 사용하므로 한 버전만 열고 업데이트 전 JSON으로 보관한다. 설정은 `daon-stage6-preferences-v1`에 별도 저장된다.

계정·기기 동기화가 없다. 사이트 데이터 삭제, 비공개 창 종료, 다른 출처로 이동하면 이어하기가 제한된다. 저장 메뉴에서 JSON 내보내기·불러오기가 가능하다. 손상/미지원 자동 저장은 자동 삭제하지 않으며 시작 화면에서 원본을 내려받을 수 있다. 새 도전 시작은 안내 후 현재 슬롯을 덮어쓴다. 저장 공간 부족 시 경고와 JSON 내보내기를 제공한다. 장기 보존은 내보낸 파일을 사용한다.

## 개발·검증

Node>=20, Python3. 앱 실행/빌드는 외부 패키지가 필요 없다. 브라우저 검증은 Playwright1.62.1과 `/usr/bin/chromium`을 사용한다. 없는 개발 환경에서는 `npm install`로 개발 의존성을 설치한다.

```
npm test
npm run test:strategies
# npm start를 별도 세션에서 실행한 후
npm run test:browser
npm run test:acceptance
npm run test:interaction
npm run test:storage
npm run test:ui
npm run package
```

범위·측정값·미확인은 [docs/validation.md](docs/validation.md), 자산은 [docs/assets.md](docs/assets.md)를 참고한다. `artifacts/`에 실제 화면과 JSON 결과가 있다. 공개 배포/호스팅/도메인 설정은 수행하지 않았다.
