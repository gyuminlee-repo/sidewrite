# Workspace and Context files

## Workspace

워크스페이스는 Claude 가 일하는 폴더다. 열린 문서의 폴더를 자동으로 따른다. Claude 는 그 안의 파일을 필요할 때 읽는다. 워크스페이스는 어느 `CLAUDE.md` 와 대화 기록을 쓸지 정한다.

### 동작

1. Setup 탭의 **Workspace** 절에 현재 폴더가 보인다.
2. 바꾸려면 **Change workspace** 를 누르고 폴더를 고른다. 위쪽 막대의 워크스페이스 칩을 눌러도 된다.
3. 다른 폴더의 문서를 열 때까지 그 선택이 유지된다.

문서가 워크스페이스 밖에 있으면 경고가 뜬다. 그 상태에서는 Claude 가 문서 옆 파일을 읽지 못한다.

### 클라우드 문서 (OneDrive / SharePoint)

자동 감지는 Office 가 알려 주는 문서 경로에 기댄다. 로컬 폴더와 Google Drive for Desktop 은 일반 파일 경로를 준다. OneDrive 와 SharePoint 문서는 `https://…sharepoint.com/…` 주소를 준다. 동기화된 로컬 경로를 돌려주는 API 는 없다.

동기화된 사본은 디스크의 OneDrive/SharePoint 폴더에 있다. **Change workspace** 로 그 로컬 폴더를 직접 지정한다. 지정하고 나면 모두 정상으로 동작한다.

### WSL 의 daemon 과 Windows 의 Office

daemon 이 WSL 안에서 돌고 Office 가 Windows 에서 돌면 Windows 경로를 WSL 경로로 바꾼다(`CC_OFFICE_WSL_PATHS`). 드라이브 루트나 Windows 프로필 폴더(`C:\Users\<you>`) 바로 아래 문서는 자동 워크스페이스를 받지 않는다. 다른 UNC 공유는 닿지 않는다.

## Context files

워크스페이스 **밖**의 자료를 Claude 에 알려 준다. 노트, 이전 초안, 사양서처럼 디스크 어디에 있든 된다. 워크스페이스 안의 파일은 추가할 필요가 없다.

### 입력

| 항목        | 값                                                                  | 필수 |
| ----------- | ------------------------------------------------------------------- | ---- |
| Path        | 파일이나 폴더 경로. **Choose file…** / **Choose folder…** 로 고른다 | 필수 |
| Description | 그 안에 무엇이 있는지 한 줄. Claude 가 언제 볼지 판단한다           | 선택 |

### 동작

1. **+ Add folder** 또는 **+ Add file** 을 누른다.
2. 경로와 설명을 넣고 **Add** 를 누른다.
3. 항목은 워크스페이스 `CLAUDE.md` 의 `CONTEXT-FILES` 블록에 경로만 저장된다. 워크스페이스 안 항목은 상대 경로가 된다.
4. 저장하면 에이전트 세션이 다시 시작된다. 대화는 `resume` 으로 이어진다.

예: `~/notes/` 폴더에 "실험 노트, 날짜별 md" 라는 설명을 붙인다.

## 다음

→ [Preferences](preferences.md)
