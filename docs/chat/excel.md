# Excel

Excel 작업창(_Sidewrite for Excel_)에는 **Chat**, **Presets**, **Setup** 세 탭이 있다. 리비전, 응답서, Tools 탭은 없다.

## Chat

Word 와 같은 채팅이다. 모델 선택은 입력 칸의 메뉴 하나뿐이다. Setup 에 Providers 절은 없다.

Excel 도구는 23개다. A1 표기와 2차원 배열을 쓴다. 범위 읽기와 쓰기, 수식 쓰기, 서식, 찾기, 행과 열 삽입과 삭제, 시트 관리, 정렬, 자동 필터, 표, 차트, 열 너비와 행 높이를 다룬다. 수식은 계산된 숫자가 아니라 수식으로 쓴다.

## 요청 예

- _"Add a Total row with `=SUM()` formulas and read back the revenue total."_
- _"Fill this sheet from the assumptions in the planning deck (added as a context file)."_

## Setup

| 절            | 내용                                                                 |
| ------------- | -------------------------------------------------------------------- |
| Workspace     | 열린 통합 문서의 폴더. **Change workspace** 로 바꾼다                |
| Context files | 워크스페이스 밖의 자료                                               |
| Preferences   | **Show diagnostics in chat** 만 있다. Track changes 는 Word 전용이다 |

[Workspace and Context files](../settings/workspace.md) 와 [Preferences](../settings/preferences.md) 는 Word 와 같다.

## 다음

→ [Providers](../settings/providers.md)
