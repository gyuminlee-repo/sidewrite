# 환경 변수

daemon 은 시작할 때 환경 변수를 읽는다. 트레이 앱(또는 `npm run dev`)을 띄우는 셸에 설정하고 다시 띄운다.

## 세션과 Claude Code

| 변수                          | 기본값       | 효과                                                                                                                                                                                                                |
| ----------------------------- | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ANTHROPIC_API_KEY`           | 없음         | 있으면 키체인 OAuth 대신 이 key 로 인증한다                                                                                                                                                                         |
| `CC_OFFICE_SETTING_SOURCES`   | `project`    | 세션이 읽는 Claude Code 설정 출처. `project`: 워크스페이스의 `CLAUDE.md` 와 `.claude/settings.json` 만. `all`: SDK 기본값(user, project, local). 빈 문자열: 없음. 그 밖에는 `user`, `project`, `local` 의 쉼표 목록 |
| `CC_OFFICE_LOAD_USER_MCP`     | 꺼짐         | `1` 이면 `~/.claude.json` 의 `mcpServers` 를 모든 세션에 합친다. 이름이 겹치면 추가 기능의 `office` 서버가 이긴다                                                                                                   |
| `CC_OFFICE_CLAUDE_PATH`       | SDK 내장 CLI | SDK 가 돌릴 로컬 `claude` 실행 파일의 절대 경로                                                                                                                                                                     |
| `CC_OFFICE_DEBUG`             | 꺼짐         | `1` 이면 `[diag]` 로거를 켠다                                                                                                                                                                                       |
| `CC_OFFICE_MODELS_TIMEOUT_MS` | `20000`      | 계정 모델 목록을 기다리는 시간(ms)                                                                                                                                                                                  |

예전 동작(사용자 수준 설정 전체와 `~/.claude.json` 의 MCP 서버)으로 돌아가려면:

```bash
CC_OFFICE_SETTING_SOURCES=all CC_OFFICE_LOAD_USER_MCP=1 npm start
```

## WSL

| 변수                       | 기본값 | 효과                                                                                                                               |
| -------------------------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------- |
| `CC_OFFICE_WSL_PATHS`      | 자동   | WSL 안의 daemon 이 Windows 경로를 WSL 경로로 바꾼다. `1` 켬, `0` 끔. 설정하지 않으면 Linux 커널 릴리스에 Microsoft 가 있을 때 켠다 |
| `CC_OFFICE_WSL_MOUNT_ROOT` | `/mnt` | WSL 안에서 Windows 드라이브가 마운트된 곳                                                                                          |

## 응답서와 프로필

| 변수                             | 기본값                    | 효과                                                                                                                |
| -------------------------------- | ------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| `CC_OFFICE_PROFILE_DIRS`         | 없음                      | 프로필 폴더 추가. `:` 로 구분                                                                                       |
| `CC_OFFICE_PROFILE`              | `generic`                 | 프로필을 정하지 않은 논문의 프로필 이름                                                                             |
| `CC_OFFICE_RATIONALE_LANG`       | `English`                 | 논문과 프로필이 정하지 않았을 때 근거 주석의 언어                                                                   |
| `CC_OFFICE_RESPONSE_MODEL`       | 가장 좋은 모델            | Writing provider 가 Claude 이고 Setup 에서 모델을 고르지 않았을 때 쓰는 Claude 모델. 다른 provider 는 바뀌지 않는다 |
| `CC_OFFICE_PYTHON`               | `python3`                 | 응답서와 문서 도구에 쓰는 Python 실행 파일                                                                          |
| `CC_OFFICE_CONFIG_DIR`           | `~/.claude/office-addins` | 사용자 프로필 폴더의 루트(`<dir>/journal_profiles`)                                                                 |
| `CC_OFFICE_REVIEW_PROMPT_FORBID` | 켬                        | `0` 이면 리뷰 턴 프롬프트에 금지 표현 목록을 넣지 않는다                                                            |

## Providers

| 변수                  | 효과                                                                    |
| --------------------- | ----------------------------------------------------------------------- |
| `CC_OFFICE_KEY_<ID>`  | provider API key. 예: `CC_OFFICE_KEY_OPENROUTER`. key 파일보다 우선한다 |
| `CC_OFFICE_CODEX_BIN` | Codex CLI 실행 파일 경로                                                |
| `CC_OFFICE_AGY_BIN`   | Antigravity CLI(agy) 실행 파일 경로                                     |

## 포트

WebSocket 은 47833, HTTP 는 47834 다. 이 포트를 바꾸는 환경 변수는 없다.
