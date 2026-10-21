# 처음 열기

## 동작

1. 트레이 앱을 먼저 띄운다. 상주 daemon 을 설치했다면 이미 떠 있다. daemon 이 없으면 Word/Excel 이 연결하지 못한다.
2. 문서나 통합 문서를 열고 추가 기능 패널을 연다.
3. 위쪽 막대의 연결 표시가 `Connecting…` 에서 `Connected` 로 바뀐다. 채팅 세션이 돌고 있으면 `Connected · <모델>` 처럼 모델 이름이 붙는다.
4. 연결 표시 옆 칩은 현재 워크스페이스 폴더다. 누르면 워크스페이스를 바꾼다.

![Word 오른쪽에 열린 Sidewrite 작업창, Review 탭](../images/guide/first-run-word-window.png)

## 워크스페이스

워크스페이스는 열린 문서가 있는 폴더로 자동으로 정해진다. Claude 는 그 폴더 안의 파일을 필요할 때 읽는다. 워크스페이스는 어느 `CLAUDE.md` 와 대화 기록을 쓸지를 정한다. 다른 폴더로 바꾸려면 Setup 탭의 **Change workspace** 를 누른다. 다른 폴더의 문서를 열 때까지 그 선택이 유지된다. 자세한 내용은 [Workspace and Context files](../settings/workspace.md) 에 있다.

## Word 작업창의 탭

| 탭       | 하는 일                                                                          | 안내                                |
| -------- | -------------------------------------------------------------------------------- | ----------------------------------- |
| Review   | 심사 의견 붙여넣기, 의견별 변경 제안, 지시 장부, 제출 전 점검                    | [원고 리비전](../revision/index.md) |
| Response | 응답서 초안, 줄 번호, 검사, 내보내기                                             | [응답서](../response/index.md)      |
| Tools    | 선택 영역 번역과 질문, 저널 지침, 서식과 분량 검사, marked copy, anonymized copy | [제출 준비](../submission/index.md) |
| Chat     | 문서에 대해 자유롭게 묻고 시키는 대화                                            | [Chat](../chat/index.md)            |
| Presets  | 한 번 눌러 보내는 프롬프트                                                       | [Presets](../chat/presets.md)       |
| Setup    | Providers, Workspace, Context files, Preferences                                 | [설정](../settings/providers.md)    |

Excel 작업창에는 Chat, Presets, Setup 세 탭만 있다([Excel](../chat/excel.md)).

## 모델 줄

탭 아래 한 줄은 작업 종류마다 마지막으로 돈 모델을 보여 준다. 예를 들면 `Edits: Claude · Opus 5.5   Writing: Claude · Opus 5.5   Judgments: Claude · Haiku 4.5` 다. 메뉴에서 고른 값이 아니라 provider 가 보고한 모델이다. 요청한 모델과 다른 모델이 돌았으면 `asked for ..., ran on ...` 이라고 쓴다. 이 줄을 누르면 Setup 이 열린다. 자세한 내용은 [Providers](../settings/providers.md) 에 있다.

## 다음

→ [원고 리비전 개요](../revision/index.md)
