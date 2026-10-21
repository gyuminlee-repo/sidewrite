# 설치

설치 경로는 두 가지다. 트레이 앱을 쓰는 기본 경로와 Windows + WSL2 에서 daemon 을 상주시키는 경로다. 두 경로를 동시에 쓰지 않는다. 둘 다 포트 47833/47834 를 쓴다.

## 트레이 앱 (macOS, Windows)

```bash
git clone https://github.com/gyuminlee-repo/sidewrite.git
cd sidewrite
npm install
npm start
```

### 동작

1. 트레이 아이콘이 뜬다. macOS 에서는 메뉴 막대, Windows 에서는 시스템 트레이다.
2. 처음 실행하면 Word 와 Excel 에 추가 기능을 설치할지 묻는다. _Install_ 을 누른다.
3. Word/Excel 을 완전히 종료했다가 다시 연다. 그래야 Office 가 manifest 를 읽는다.
4. 추가 기능을 연다.
   - macOS: Insert → Add-ins → **Developer Add-ins** → _Sidewrite_ 또는 _Sidewrite for Excel_.
   - Windows: `WEF\Developer` 레지스트리 키로 자동 사이드로드된다. **Insert → Add-ins** 또는 **Home → Add-ins** 에서 찾는다.

!!! note "호스트당 한 벌"
manifest 의 add-in `<Id>` GUID 가 고정이라 Word(또는 Excel) 하나에는 한 벌만 사이드로드된다. 클론 두 개를 나란히 쓰려면 한쪽 `manifests/word.xml` / `excel.xml` 의 `<Id>` 를 새 UUID 로 바꾼다.

daemon 로그를 터미널에서 보려면 `npm start` 대신 `npm run dev` 를 쓴다. Electron 셸 없이 daemon 만 뜬다.

### 이전 Draftspect 설치에서 올리기

Sidewrite 설치를 다시 하거나 트레이 앱에서 _Install_ 을 고른 뒤 Office 를 완전히 다시 연다. 설치기는 이전에 배포된 manifest 와 그에 맞는 `WEF\Developer` 값만 지운다. 손으로 고친 manifest, 서비스, 로그, 인증 정보, 사용자 파일은 남긴다.

## Windows + WSL2: 터미널 없이 상주하는 daemon

WSL2(systemd 사용) 안의 체크아웃을 같은 PC 의 Word 에서 쓰는 경우다. 설치하고 나면 Windows 로그온부터 daemon 이 돌고 Word 에 추가 기능이 보인다. 터미널도 트레이 앱도 필요 없다.

```bash
npm install
bash scripts/install/install.sh      # 설치. 다시 실행해도 이미 있는 것은 건너뛴다
node scripts/install/doctor.mjs      # 무엇이 갖춰졌고 무엇이 빠졌는지, 고치는 명령
bash scripts/install/daemon-service.sh logs     # daemon 로그 따라 보기
bash scripts/install/daemon-service.sh status
bash scripts/install/uninstall.sh    # 전부 되돌리기
```

### `install.sh` 가 하는 일

모두 현재 사용자 기준이고 관리자 권한이 필요 없다.

| 단계                     | 하는 일                                                                                                                                                                      |
| ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1/4 daemon service       | 이 체크아웃의 daemon 을 systemd 사용자 서비스로 등록한다. 실패하면 다시 띄우고 로그인 없이도 돈다(linger). WSL 홈의 Claude 로그인을 쓰므로 먼저 `claude` 로 한 번 로그인한다 |
| 2/4 Word add-in          | Word manifest 를 `%APPDATA%\Sidewrite\manifests` 에 두고 `WEF\Developer` 레지스트리 키에 값을 넣는다. 트레이 앱의 _Install_ 과 같다. 끝나면 Word 를 다시 연다                |
| 3/4 WSL at Windows logon | Windows 로그온 때 WSL 을 띄우는 예약 작업을 만든다. Word 를 열 때 서비스가 이미 떠 있게 한다                                                                                 |
| 4/4 doctor               | 갖춰진 것과 빠진 것을 보고한다. 종료 코드는 doctor 가 찾은 문제 수다                                                                                                         |

설치기는 패키지를 설치하지 않고 환경 변수를 설정하지 않는다. 두 PowerShell 스크립트는 그 프로세스에서만 `-ExecutionPolicy Bypass` 로 실행된다. 컴퓨터의 정책은 바뀌지 않는다.

체크아웃을 옮기거나 Node 버전을 바꾼 뒤에는 `install.sh` 를 다시 실행한다. 서비스가 둘 다에 묶여 있다.

### 되돌리기

`uninstall.sh` 는 역순으로 로그온 작업, Word 등록, systemd 서비스를 지운다. daemon 로그, `~/.claude/office-addins` 의 다른 파일(키, bridge token), 원고 옆의 파일은 일부러 남긴다.

## 자동 설치가 안 될 때

[트러블슈팅](../troubleshooting/index.md#add-in-install) 의 수동 등록 절차를 따른다.

## 다음

→ [처음 열기](first-run.md)
