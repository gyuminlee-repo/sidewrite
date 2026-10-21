# Changelog

## 1.0.0 (2026-10-20)

첫 릴리스. Sidewrite 는 Word 에서 원고 리비전을 진행한다. 심사 의견, 근거 주석이 하나씩 붙은 추적된 변경, 응답서, 그림 기록, 제출 점검을 로컬 Claude Agent SDK daemon 위에서 돌린다.

### 이 릴리스에서 확인한 것

Windows Word(Office 16)에서 WSL 로 구동해 이 릴리스의 코드로 확인했다.

- `npm test`: 2672 pass, 0 fail, 0 skipped. Python unittest 20개 파일 통과.
- 무인 수락 실행의 단계 s2, s25, s3, s4, cite, parts, resident 가 모두 exit 0.
- 실제 저널 리비전 한 건의 재실행: 심사 의견 24개, 턴 24개, 실패한 턴 0, 모든 변경에 근거 주석, 응답 초안 24개.

macOS 는 이 릴리스에서 다시 확인하지 않았다.

### main 의 마지막 변경 이후

- Propose 와 Propose all 은 다른 리뷰 파일 요청과 같은 시간(120 s)만큼 daemon 의 답을 기다린다. 작업창이 포기한 턴은 시작하지 않는다 (#53).

### 알려진 점

- 추가 기능 manifest 는 `<Version>0.1.2.0</Version>` 을 유지한다. 기존 설치가 제자리에서 올라가게 하려는 것이다.

[GitHub 릴리스 페이지](https://github.com/gyuminlee-repo/sidewrite/releases/tag/v1.0.0)
