# TripCoupon Partner Manager

TripCoupon 제휴사 목록을 CSV로 업로드해 이전 MASTER와 비교하는 브라우저 기반 관리 도구입니다.

- Dashboard
- Partners
- Import & Compare
- Change History
- Existing / New / Discontinued / Changed 자동 분류

## 데이터 저장 방식
현재 MVP는 서버 없이 브라우저 localStorage에 데이터를 저장합니다. 따라서 공유 URL은 동일하지만 각 사용자의 브라우저 데이터는 서로 공유되지 않습니다. 운영 단계에서 공용 Google Sheet/DB 연동을 추가할 수 있습니다.

## 보안
검색엔진 색인을 막는 meta 태그가 포함되어 있지만 URL 자체는 공개될 수 있습니다. 민감정보 운영 전에는 인증 기능을 추가하세요.
