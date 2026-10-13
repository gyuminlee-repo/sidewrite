# 외부 추적 replace 끝 삭제의 판정 증거 조사

## 결론과 범위

- 기준: 2026-10-13 `main`의 `34a2a564e55bc54a860fe2639130fde825aac1d5`.
- 분기: `fix/external-insertion-evidence-20261013`.
- 현재 저장하는 정보로는 저자의 외부 추적 replace와 무관한 추적 삽입·삭제를 구별할 수 없다.
- P2의 `accepted` 오판은 합성 입력으로 재현했다. 실제 Word에서 재현한 결과는 아니다.
- 판정 코드를 바꾸지 않는다. 조사 문서와 합성 재현 스크립트만 추가한다. P2는 해결되지 않았다.
- PR #32가 막은 `accepted` → `pending` 회귀를 피하려면 추적 run의 출처와 도구 삭제 직전 범위를 연결하는 증거가 먼저 필요하다.
- 별도 P3 작업의 문단 번호 정수 조건과 follow-through 귀속 변경은 포함하지 않는다.

가정은 저장소의 기존 회귀 예제와 같다. Word가 추적 삽입 안의 글자를 지우면 그 글자가 별도 추적 삭제로 남지 않을 수 있다. 이 조사에서는 교량 점검 문장으로 그 입력 형태를 구성했다. 실제 문서나 계정은 사용하지 않았다.

## 현재 들어오는 정보

아래 위치는 모두 위 기준 커밋의 행 번호다.

| 경로                                          | 보존하는 정보                                                                                            | 이번 판정에 부족한 정보                                     |
| --------------------------------------------- | -------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| `taskpane/shared/tools-word.js:2399-2407`     | 문단 `index`, `id`, `text`, `reviewed_text`, `w_ins`, `w_del`, `tracked`                                 | 문단 OOXML 원문은 스냅샷에 보내지 않음                      |
| `taskpane/shared/review-view.js:196-256`      | run의 `kind`, `text`, `start`, `marks`                                                                   | `w:author`, `w:date`, `w:id`를 파싱하거나 반환하지 않음     |
| `taskpane/shared/review-view.js:244-249`      | 인접한 같은 종류의 run을 합친 텍스트와 표식 개수                                                         | 작성자·시각·원래 추적 요소 경계가 다른 run도 병합될 수 있음 |
| `daemon/revision-classify.mjs:174-208`        | 정규화된 본문 위의 `kind`, `start`, `end` 및 불완전 여부                                                 | 원래 run의 출처나 replace 쌍 식별 정보 없음                 |
| `daemon/revision-store.mjs:461-470`           | 외부 항목의 `id`, `detected_at`, `kind`, `paragraph_index`, `old_index`, `before`, `after`, `lost_facts` | 문단 고유 ID·글자 범위·추적 run·작성자·실제 편집 시각 없음  |
| `daemon/revision-store.mjs:527-545,1421-1428` | 마지막 새로고침의 문단 텍스트·문단 ID 배열·읽기 상태                                                     | 이전 추적 run과 OOXML을 저장하지 않음                       |

`start`는 주석 앵커 제거와 공백 축약을 거친 문단 좌표다. Word의 원시 range 좌표라고 해석할 수 없다. `placedRuns`는 현재 본문에서 run 텍스트를 다시 찾고 시작 위치와 가까운 후보를 선택한다.

`outsideChanges`는 문단 단위 차이를 만든다(`daemon/doc-diff.mjs:174-209`). `kind: "modified"`는 특정한 추적 replace라는 뜻이 아니다. 여러 차례의 독립 편집이나 수락·거절도 같은 문단 차이를 남길 수 있다.

`snapshotTexts`는 가능한 경우 `reviewed_text`를 쓴다(`daemon/doc-diff.mjs:24-28`). 이 값에는 추적 삭제의 원문과 삽입·삭제의 인접 관계가 보존되지 않는다. `reviewed_text`가 없으면 원문 텍스트를 쓰지만 run의 출처는 여전히 없다.

`addExternal`의 `detected_at`은 발견 시각이다(`daemon/revision-store.mjs:1433-1441`). 실제 Word 편집 시각이나 도구 삭제 전후 순서를 증명하지 않는다. 외부 항목은 최신 100개만 남고 사용자가 지울 수도 있다. 현재 `classifyRevision`은 `comments[].changes`만 형제 목록으로 전달한다(`daemon/revision-classify.mjs:424-432`). 외부 항목을 그대로 그 목록에 넣을 수 없다.

## 합성 재현과 대조군

문단 ID는 `P4`이고 문단 번호는 `4`다. 도구가 본 삭제 전 문장은 다음과 같다.

```text
Crews installed new steel supports old piers beside the road.
```

도구의 삭제는 `find: " supports"`, `text: ""`다. 삭제 후 원문 텍스트는 다음과 같다.

```text
Crews installed new steel old piers beside the road.
```

현재 문단의 추적 run은 `trackedRuns`의 실제 반환값으로 구성했다.

```json
[
  { "kind": "ins", "text": "new steel ", "start": 16, "marks": 1 },
  { "kind": "del", "text": "old piers", "start": 26, "marks": 1 }
]
```

아래 두 종류의 합성 OOXML을 `trackedRuns`에 넣었다.

- 같은 작성자·시각: 삽입 `w:id="10"`, 삭제 `w:id="11"`, 작성자 `Bridge editor`, 시각 `2026-10-12T09:00:00Z`.
- 다른 작성자·시각: 삽입 `w:id="23"`, 작성자 `Catalog editor`, 시각 `2026-10-12T10:00:00Z`. 삭제 `w:id="45"`, 작성자 `Road editor`, 시각 `2026-10-12T11:00:00Z`.

두 XML에서 글자와 요소 순서는 같다. 파서가 속성을 버리므로 두 스냅샷은 `assert.deepEqual`로 같다. 이 검사는 메타데이터가 소실된다는 증거다. 같은 작성자·시각이라는 이유만으로 실제 replace가 입증된다는 뜻은 아니다.

| 경우                                                                                                                   | 기대 또는 유지해야 할 상태     | 현재 결과               |
| ---------------------------------------------------------------------------------------------------------------------- | ------------------------------ | ----------------------- |
| 저자의 외부 replace가 `old piers`를 `new steel supports `로 바꾼 뒤 도구가 삽입 끝을 삭제했다고 가정. 기록된 형제 없음 | `pending`                      | `accepted`: 알려진 한계 |
| 독립 삽입과 독립 삭제 사이의 일반 글자 ` supports`를 삭제하고 수락했다고 가정                                          | `accepted`                     | `accepted`              |
| 같은 문단에 `find: "old piers"`, `text: "new steel supports "`인 기록된 replace 형제 존재                              | `pending`                      | `pending`               |
| 외부 before/after 항목만 추가                                                                                          | 기존 증거 부족을 해소하지 못함 | 분류 결과 변화 없음     |

첫 두 경우는 현재 분류기에 같은 입력을 줄 수 있다. 현재 스냅샷의 인접성이나 같은 문단이라는 사실만으로 첫 경우를 `pending`으로 바꾸면 둘째 경우도 바뀐다. 외부 항목의 문단 전체 before/after를 추적 replace 기록으로 해석해도 개별 편집의 관계를 증명하지 못한다.

재현에서는 `classifyChange`, `classifyRevision`, `trackedRuns`, `countRevisionMarks`, `outsideChanges`, `addExternal`을 직접 호출했다. [재현 스크립트](external-insertion-evidence-2026-10-13/reproduce.mjs)는 메모리의 합성 데이터만 사용한다. 제품 코드나 기존 테스트는 바꾸지 않았다. 기존 테스트가 의도적으로 지키는 `accepted` 결과를 바꾸는 실패 테스트도 남기지 않았다.

저장소 루트에서 기존 의존성이 준비된 상태로 다시 실행할 수 있다. Word나 `scripts/e2e/*`는 호출하지 않는다.

```sh
node docs/review/external-insertion-evidence-2026-10-13/reproduce.mjs
```

## `outcome_unknown`을 형제 후보로 넣을지

미해결 항목을 무조건 형제 증거로 넣지 않는다.

`Attempt`는 Word 호출 전에 저장하는 예상 변경이다(`daemon/revision-store.mjs:115-125,1226-1242`). `outcome_unknown`은 Word가 적용했다고 확인했다는 상태가 아니다. 응답 소실 또는 기록 실패로 적용 여부를 모르는 상태다. 요청의 `find`와 `text`가 현재 삽입·삭제와 비슷하다는 사실만으로 그 시도가 해당 run을 만들었다고 단정할 수 없다.

합성 검사는 다음 결과를 확인했다.

1. 새 글자 중 `supports`가 사라졌고 근거 주석도 없는 replace 시도는 미해결 상태로 남는다.
2. 이 시도를 적용된 형제처럼 직접 전달하면 독립 삽입·삭제 대조군도 `pending`으로 바뀐다. 단순한 후보 확대는 PR #32의 회귀를 되살릴 수 있다.
3. 같은 시도에 현재 복구 조건을 만족하는 근거 주석을 제공하면 `resolveUnknownOverrides`가 변경으로 기록한다. 뒤의 `classifyRevision`에서 삭제가 `pending`으로 판정된다.

현재 함수는 각 시도를 판정할 때 먼저 복구된 `next.comments[].changes`를 사용한다(`daemon/review-service.mjs:175-191`). 새로고침은 복구 뒤에 전체 분류를 수행한다(`daemon/review-service.mjs:1495-1499`). 이 조사에서 확인한 주석 증거가 있는 경우는 이미 그 경로로 처리된다.

앞서 판정한 시도를 나중에 복구한 형제로 다시 판정하는 고정점 처리 여부는 별도의 복구 설계 문제다. 그런 처리를 추가해도 근거가 전혀 없는 외부 replace의 출처가 생기지는 않는다. 이번 문서에서는 복구 순서나 후보 규칙을 바꾸지 않는다.

## 안전한 수정에 필요한 추가 증거

다음은 후속 설계와 실제 Word 검증의 요구사항이다. 새 스키마나 수집 코드는 이번 PR에 넣지 않는다.

1. 작업창이 원래 추적 요소의 ID·작성자·시각·종류·텍스트·요소 경계를 보존해야 한다. 작성자나 시각이 다르다는 이유로 정보를 버리지 않도록 현재 병합 규칙도 검토해야 한다.
2. 도구 삭제 직전의 해당 문단 ID와 삽입 run 범위를 확보해야 한다. 지운 글자가 그 삽입 run의 끝에 실제로 속했다는 관계가 필요하다. 현재 스냅샷만 보면 사라진 글자가 추적 삽입이었는지 일반 글자였는지 알 수 없다.
3. 외부 replace의 삽입·삭제를 연결하는 검증된 관계가 필요하다. 작성자·시각이 같다는 사실이나 연속된 ID만으로 같은 작업이라고 가정하지 않는다. 같은 저자의 독립 편집도 대조군에 포함해야 한다.
4. 원시 좌표와 공백 정규화 좌표의 변환 및 문단 이동을 처리해야 한다. 이전/현재 스냅샷 버전과 문단 고유 ID를 묶어 오래된 외부 기록이 다른 문단의 증거가 되지 않게 해야 한다.
5. 외부 기록을 판정에 쓰려면 문단 전체 차이와 검증된 추적 변경 증거를 구별해야 한다. 근거가 없는 기존 기록을 새 형제 증거로 승격하면 안 된다.
6. 실제 Word에서 외부 replace 끝 삭제와 무관한 삽입·삭제 사이의 수락을 비교해야 한다. 분할 run·공백·주석 앵커·같은 작성자의 연속 편집·문단 이동도 확인해야 한다.

이 정보를 확보한 뒤에도 최소 수정의 수용 기준은 세 가지다. 외부 replace 끝 삭제는 `pending`이어야 한다. 독립 삽입·삭제 사이의 수락은 `accepted`여야 한다. 기록된 형제 replace가 있는 기존 경우는 계속 `pending`이어야 한다.

## 검증

- Node `v24.19.0`, npm `11.9.0`.
- 설치 명령 없이 기존 별도 체크아웃의 의존성을 이 작업 폴더로 복사했다. `prettier 3.8.3`, `zod 4.4.3`, `ws 8.22.0`, Agent SDK `0.3.285`, Anthropic SDK `0.93.0`, MCP SDK `1.31.0`을 확인했다.
- 합성 증거 검사의 모든 assertion이 통과했다. 저자 replace의 현재 `accepted` 결과는 해결 성공이 아니라 한계 재현이다.
- `tests/revision-classify-dropped-text.test.mjs`는 변경하지 않았다. 기준 커밋과 작업 파일의 Git blob hash는 모두 `c0a14107204b9306589c5cd51ac888224a35ed8b`다.

실행한 기존 테스트:

```sh
node --test --test-reporter=tap \
  tests/revision-classify-dropped-text.test.mjs \
  tests/revision-classify.test.mjs \
  tests/revision-classify-stacked.test.mjs \
  tests/revision-store.test.mjs \
  tests/review-view.test.mjs \
  tests/edit-attempt.test.mjs
```

```text
# tests 137
# pass 137
# fail 0
```

전체 `npm test`는 SKIP이다. 제품 코드를 바꾸지 않는 보고서 PR이며 전체 테스트에는 실행이 금지된 `scripts/e2e/make_parts_fixture.py` 등의 호출이 포함된다. 금지된 보조 스크립트를 우회 실행하지 않았다. 일부 테스트 통과를 전체 통과로 기록하지 않는다.

문서와 재현 스크립트의 `npx --no-install prettier --check` 및 재현 스크립트의 `node --check`는 통과했다.

Word·리플레이·실제 원고 검증은 수행하지 않았다. 별도 P3 브랜치를 합치지 않았고 제품 코드의 통합 검증도 이 문서의 범위가 아니다. 같은 판정 파일을 수정하는 P3 작업이 먼저 머지되면 이 문서의 기준 행 번호는 다시 확인해야 한다.

## 재현 출력

합성 입력 검사의 실제 출력이다. 상태 문자열은 테스트한 API의 반환값을 그대로 적었다.

```text
Equal classifier inputs after author/date/revision-ID differences are discarded: true
Author-replace tail deletion without recorded sibling: accepted (known limitation; desired pending)
Independent ins/del with accepted deletion: accepted (control)
Recorded same-paragraph replace sibling: pending (control)
Tracked runs: [{"kind":"ins","text":"new steel ","start":16,"marks":1},{"kind":"del","text":"old piers","start":26,"marks":1}]
External before/after entry changes classifier result: false
External fields: kind, paragraph_index, old_index, before, after, lost_facts, id, detected_at
Unknown replace without full new text or rationale comment: unresolved
Blindly treating that attempted replace as applied sibling changes the independent control to: pending (unsafe)
Unknown replace with matching rationale comment: resolved; subsequent classification pending
All synthetic evidence assertions passed.
```
