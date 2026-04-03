오늘 작업한 내용을 CLAUDE.md에 반영하고, GitHub에 commit & push합니다.

## 실행 순서

### 1단계 — CLAUDE.md 구현 현황 갱신

`git status`와 `git diff --stat HEAD`로 변경 파일 목록을 파악한 뒤,
`CLAUDE.md`의 **"구현 현황"** 섹션을 실제 상태에 맞게 수정한다.

- 새로 생성된 파일 → "완료된 파일" 표에 추가
- 삭제된 파일 → 표에서 제거
- "미구현 (다음 단계)" 목록도 실제 상태 반영
- 날짜는 오늘 날짜(`YYYY-MM-DD`)로 업데이트

### 2단계 — 변경 내용 요약 및 커밋 메시지 초안 출력

변경 파일을 분석해서 아래 형식으로 커밋 메시지 초안을 **먼저 출력**한다.
(실제 커밋은 아직 하지 않는다)

```
<type>: <한 줄 요약>

<변경된 주요 내용 3줄 이내>
```

type: `feat` / `fix` / `docs` / `refactor` / `chore`

### 3단계 — 사용자 확인 후 진행

"위 내용으로 commit & push 진행할까요?" 라고 묻는다.
사용자가 **확인**하면 4~5단계를 실행한다.
사용자가 **수정**을 요청하면 메시지를 수정한 뒤 다시 확인한다.

### 4단계 — git add & commit

`.env`, `data/input/` 파일은 제외하고 스테이징 후 커밋한다.

```bash
git add --all -- ':!.env' ':!data/input/*'
git commit -m "<확인된 메시지>

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```

### 5단계 — push

```bash
git push origin master
```

push 완료 후 결과를 한 줄로 요약한다.
