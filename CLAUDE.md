# 프로젝트 규칙 (CLAUDE.md)

## 수정 원칙
- 지시된 범위만 수정. 지시 외 코드 변경 금지.
- 전면 재작성 시 기존 함수 목록과 비교하여 누락 없는지 확인 필수.

## 코드 작성 규칙
- 이벤트 핸들러는 `<button onclick="...">` 사용. `<div onclick>` 금지 (GAS Caja 제거 정책).
- **clasp push 전 반드시 `node --check Code.js` 실행하여 JS 파싱 오류 확인.**
- index.html 스크립트 블록은 `new Function(code)` 또는 동등한 방법으로 문법 검증 후 push.

## 배포 흐름
- `index.html` 수정 → `git push` → GitHub Pages 자동 반영 (GAS 재배포 불필요)
- `staff.html` / `admin.html` 수정 → `clasp push` + GAS 재배포 필요
- `Code.js`만 수정 → `clasp push` + **GAS 재배포 필요** (doPost 등 웹앱 엔드포인트 변경 포함)

## GAS 재배포 후 필수: 배포 URL 자동 동기화
GAS에서 **새 배포**를 생성하면 배포 ID가 바뀐다. 기존 배포를 ✏️ 수정할 때는 ID가 유지된다.

재배포 후 반드시 다음 절차 수행:
1. `clasp deployments` 실행
2. 출력 마지막 줄의 배포 ID 추출 (예: `AKfycbwp...`)
3. `index.html`의 `GAS_POST_URL` / `GAS_CONFIG_URL`과 비교
4. 다르면 두 변수 모두 새 ID로 교체 후 `git push`

**이 단계를 건너뛰면 doPost가 구 버전 엔드포인트를 호출해 시트/메일이 무응답 상태가 된다.**

## clasp push / git push
- Claude가 직접 수행.
- push 전 `node --check Code.js` 파싱 검증 필수.

## 트리거 현황 (2026-04-14 기준)
- `warmup`: 30분 간격, 콜드스타트 방지 (행사 후 자동 삭제)
- `sendRemindMails`: 2026-05-07 09:00 KST 1회 발송
- `syncFirestoreToSheets` / `autoSyncFirebaseToSheets`: 폐기됨 (Firebase IAM 403, doPost 방식으로 대체)
- 트리거 초기화 필요 시 `fixAllTriggers()` 실행
