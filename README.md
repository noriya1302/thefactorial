# 팩토리얼 모의고사 (thefactorial.web.app)

- `public/index.html` : 사이트 전체 (한 파일)
- `firestore.rules` : Firestore 보안 규칙
- `main` 브랜치에 올라가면 GitHub Actions가 자동으로 https://thefactorial.web.app 에 배포합니다.

## 자동 배포 연결 (처음 한 번)
사이트 폴더에서 `firebase init hosting:github` 실행 → 저장소 `noriya1302/thefactorial` 입력.
GitHub 저장소에 `FIREBASE_SERVICE_ACCOUNT_FACTORIAL_E1571` 비밀값이 생기면 연결 완료.

## 보안 규칙을 바꿨을 때
자동 배포는 사이트만 올립니다. 규칙은 `firebase deploy --only firestore:rules` 로 따로 올립니다.

## 주문 단계(stage)
- 답지 주문: 0 접수 → 1 입금 확인 → 2 PDF 발송
- THEFACTORIAL: 0 접수 → 1 입금 안내 → 2 입금 확인 → 3 제작 중 → 4 PDF 발송
