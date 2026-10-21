# 트러블슈팅

자주 생기는 문제와 해결 절차.

| 증상                                                                          | 해결                                                             |
| ----------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| 작업창 연결 표시가 `Disconnected` 상태로 다시 시도 중이다                     | [Daemon not running](#daemon-not-running)                        |
| `Sign-in required` 배너                                                       | [Sign-in required](#sign-in-required)                            |
| 설정한 MCP 서버가 에이전트에 안 보인다                                        | [MCP server not visible](#mcp-server-not-visible)                |
| 워크스페이스가 엉뚱한 폴더다                                                  | [Wrong workspace](#wrong-workspace)                              |
| 추가 기능 자동 설치가 안 됐다                                                 | [Add-in install](#add-in-install)                                |
| 갤러리의 추가 기능 아이콘이 비어 있다 (macOS)                                 | [Blank add-in icon](#blank-add-in-icon)                          |
| Context files 변경이 반영되지 않는다                                          | [Context files not applied](#context-files-not-applied)          |
| 응답서 줄 번호를 못 찾는다                                                    | [Line numbers](#line-numbers)                                    |
| Word 창을 내린 뒤 변경에 근거 주석이 없다                                     | [Word 에서 결정](../revision/04-decide.md)                       |
| `Blocked by ...` 로 편집이 막혔다                                             | [금지 표현에 막힌 편집](../revision/02-ledger.md)                |
| `Not exported: ...`                                                           | [Export](../response/03-export.md)                               |
| 제출 점검의 copy 행이 `This document has changes that are not saved...`       | [Before submission](../submission/before-submission.md)          |
| 보충 파일 행이 `... changed since Sidewrite last read it`                     | [Supplementary files](../revision/06-supplements.md)             |
| `Save the document to a local folder to use the tools.`                       | 문서를 로컬 폴더에 저장한다. [제출 준비](../submission/index.md) |
| `Waiting for the daemon to confirm the stored provider choice...` 가 계속된다 | [Providers](../settings/providers.md)                            |

추가 문제는 [GitHub Issues](https://github.com/gyuminlee-repo/sidewrite/issues) 에 등록한다.

## Daemon not running

daemon 이 돌지 않거나 죽었다.

1. 트레이 메뉴 → **Open logs** 로 로그를 연다(`~/.claude/office-addins/daemon.log`).
2. 트레이 앱은 daemon 을 최대 3번까지 자동으로 다시 띄운다. 그 뒤에는 멈춘다.
3. 상주 daemon 을 설치했다면 `bash scripts/install/daemon-service.sh status` 와 `node scripts/install/doctor.mjs` 로 상태를 본다.

트레이 앱과 상주 daemon 을 동시에 띄우지 않는다. 둘 다 포트 47833/47834 를 쓴다.

## Sign-in required

Agent SDK 가 인증하지 못했다. 둘 중 하나를 한다.

- 터미널에서 `claude` 를 실행해 로그인한다.
- 셸에 `ANTHROPIC_API_KEY` 를 설정하고 트레이 앱을 다시 띄운다. daemon 은 시작할 때 환경을 읽는다.

둘 중 하나를 마친 뒤 트레이 앱을 끄고 다시 연다.

## MCP server not visible

사용자 MCP 서버는 기본으로 꺼져 있다.

1. `CC_OFFICE_LOAD_USER_MCP=1` 을 설정하고 트레이 앱을 다시 띄운다.
2. daemon 시작 로그에 `Loaded N MCP server(s) from ~/.claude.json: …` 가 보이는지 확인한다.
3. 처음 연결이 실패한 서버는 SDK 가 조용히 뺀다. 서버가 살아난 뒤 트레이 메뉴 → **Restart daemon** 을 누른다.

`~/.claude.json` 은 daemon 시작 때 한 번만 읽는다. 바꾸면 daemon 을 다시 띄운다.

## Wrong workspace

워크스페이스는 열린 문서의 폴더를 자동으로 따른다. 엉뚱하다면 전에 다른 폴더를 직접 골랐거나 문서가 클라우드(OneDrive/SharePoint)에 있다. Setup 탭의 **Change workspace** 로 지정한다([Workspace and Context files](../settings/workspace.md)).

## Add-in install

Word/Excel 에서 **Insert → Add-ins → Manage My Add-ins → Upload My Add-in** 을 고르고 `manifests/` 의 파일을 지정한다. 이 메뉴가 없으면 아래를 따른다.

- macOS: `manifests/word.xml` 을 `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` 에 복사한다. 폴더가 없으면 만든다. Excel 도 같다.
- Windows: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` 아래에 `REG_SZ` 값을 더한다. 값 이름과 데이터 모두 manifest 의 전체 경로다. 값을 넣은 뒤 Office 를 다시 연다.

로컬 경로의 Trusted Add-in Catalog 는 Office 가 무시한다. 실제 UNC 네트워크 공유가 필요하다.

## Blank add-in icon

겉모습만의 문제다. macOS Office 는 갤러리 타일에 `http://127.0.0.1` 아이콘을 자주 가져오지 않는다. 추가 기능 자체는 동작한다.

## Context files not applied

context 항목을 저장하면 에이전트 세션이 다시 시작된다. 대화는 `resume` 으로 이어지고 새 `CLAUDE.md` 가 로드된다. 변경이 보이지 않으면 위쪽 막대의 워크스페이스 칩이 맞는 폴더인지 확인한다.

## Line numbers

`No line numbers were found in the PDF. ...` 이 뜨면 Word 에서 Layout > Line Numbers > Continuous 를 켜고 **Update line numbers** 를 다시 누른다. PyMuPDF 가 없으면 줄 번호를 읽지 못한다(`pip install pymupdf`). 원고 PDF 가 너무 크면 `The PDF of the manuscript is too large to send to the daemon (n MB).` 가 뜬다.
