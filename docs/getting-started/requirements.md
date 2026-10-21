# 요구 사항

앱을 띄우기 전에 아래를 갖춘다.

| 항목                     | 확인 방법                                                                                                                                            |
| ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| macOS 또는 Windows 10/11 | 다른 플랫폼에서는 manifest 가 로드되지 않는다                                                                                                        |
| Node.js 18 이상          | `node --version`                                                                                                                                     |
| Microsoft Word 와 Excel  | Microsoft 365, Office 2019, 2021. Word 는 WordApi 1.4 이상, Excel 은 ExcelApi 1.4 이상이 필요하다                                                    |
| 로그인한 Claude Code     | 터미널에서 `claude` 를 한 번 실행해 로그인한다. 대신 `ANTHROPIC_API_KEY=sk-ant-...` 를 실행 전에 export 해도 된다. API key 가 있으면 그쪽이 우선이다 |
| Git                      | 저장소를 받을 때 쓴다                                                                                                                                |

Claude Code 가 없으면 아무것도 연결되지 않는다. 먼저 [Claude Code 를 설치](https://docs.claude.com/en/docs/claude-code/overview)한다.

## 원고 리비전 기능에 더 필요한 것

| 기능                                 | 필요한 것                                                |
| ------------------------------------ | -------------------------------------------------------- |
| 응답서 쓰기, 읽기, 검증              | daemon 이 도는 컴퓨터의 Python 3. 표준 라이브러리만 쓴다 |
| 응답서의 줄 번호                     | PyMuPDF (`pip install pymupdf`)                          |
| Figures and other files 의 비교 그림 | PyMuPDF                                                  |
| Marked copy                          | Python 패키지 lxml                                       |

PyMuPDF 가 없으면 줄 번호를 읽지 못한다. 줄 번호 검사가 실패하고 응답서는 "Export anyway" 로만 줄 번호 없이 저장된다. 그림 기록과 채택은 PyMuPDF 없이도 된다.

Python 실행 파일은 기본으로 `python3` 이다. 다른 것을 쓰려면 `CC_OFFICE_PYTHON` 을 지정한다([환경 변수](../reference/environment.md)).

## 다음

→ [설치](install.md)
