# 원고 옆에 생기는 파일

`<원고 이름>` 은 확장자를 뺀 원고 파일 이름이다. 예를 들어 `~/papers/draft.docx` 의 리비전 파일은 `~/papers/draft.revision.json` 이다.

## 원고 옆

| 파일                                    | 쓰는 곳                          | 내용                                                                                                                                |
| --------------------------------------- | -------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `<원고 이름>.revision.json`             | Review, Response, figures 페이지 | 심사 의견, 리뷰 흐름이 만든 모든 추적된 변경(근거 주석 id, 이전과 이후 문장), 금지 표현이 막은 편집, 지시 장부, 이 논문의 프로필 값 |
| `<원고 이름>.decisions.jsonl`           | Review                           | 모델 결정마다 한 줄. 입력 텍스트 대신 SHA-256 지문, 모델, 짧은 출력, 사람이 채울 빈 `expected` 칸. 언제 지워도 된다                 |
| `<원고 이름>.comments_backup.json`      | Delete ALL comments              | 지운 주석과 답글의 백업. 이미 있으면 시각이 붙은 이름                                                                               |
| `<원고 이름>.ledger.json`               | Export ledger                    | 장부 규칙만                                                                                                                         |
| `<원고 이름>.response.docx`             | Export letter                    | 응답서. 이름이 있거나 열려 있으면 `-2`, `-3`                                                                                        |
| `<원고 이름>.response.model.json`       | Export letter                    | 쓴 그대로의 응답서 문단                                                                                                             |
| `<원고 이름>.response.manuscript.json`  | Export letter                    | 인용문을 대조한 원고 문단                                                                                                           |
| `<원고 이름>.response.R1.txt`, `.E.txt` | Export plain text per reviewer   | 심사자별 텍스트, 편집자는 `E`                                                                                                       |
| `<원고 이름>_marked.docx`               | Marked copy                      | 표시본                                                                                                                              |
| `<원고 이름>_anonymized.docx`           | Anonymized copy                  | 익명본. 검사 실패 시 `.INCOMPLETE.docx`                                                                                             |
| `figures/submission/`                   | figures 페이지 Adopt             | 채택한 파일과 `FILE_SOURCES.md`                                                                                                     |
| `figures/.renders/`                     | figures 페이지                   | 비교 그림. 내용 해시 이름                                                                                                           |
| `journal_profiles/<이름>.json`          | 직접 둔다                        | 이 폴더의 저널 프로필                                                                                                               |
| `submission.json`                       | 직접 둔다                        | 프로필 이름과 override 의 대체값                                                                                                    |

모든 쓰기는 같은 폴더의 임시 파일에 쓴 뒤 이름을 바꾼다. 파일마다 쓰기를 순서대로 처리한다.

보충 파일을 묶으면 각 문서가 자기 `<이름>.revision.json` 을 가진다. 연결은 두 파일 모두에 적힌다(보충 파일의 `part_of`, 원고의 `parts`).

## 워크스페이스와 사용자 폴더

| 위치                                             | 내용                                            |
| ------------------------------------------------ | ----------------------------------------------- |
| 워크스페이스 `CLAUDE.md` 의 `CONTEXT-FILES` 블록 | Context files 항목. 경로만                      |
| `~/.claude/office-addins/daemon.log`             | daemon 로그. 채팅 메시지 평문 포함              |
| `~/.claude/office-addins/sessions.json`          | 워크스페이스와 세션 기록                        |
| `~/.claude/office-addins/provider-keys.json`     | provider API key (권한 0600)                    |
| `~/.claude/projects/<hash>/<session_id>.jsonl`   | Agent SDK 대화 기록. 터미널 Claude Code 와 같다 |

`uninstall.sh` 는 `~/.claude/office-addins` 의 파일과 원고 옆 파일을 지우지 않고 남긴다. 필요 없으면 직접 지운다.
