## OpenAPI CALL 데이터 추출 도구

OpenAPI CALL은 공공데이터포털의 OpenAPI를 누구나 쉽게 조회·저장할 수 있도록 만든 GUI 프로그램입니다.  
복잡한 요청 과정이 어려운 사용자를 위해 테이블 형태로 결과를 쉽게 받을 수 있도록 개발되었습니다.

### 주요 기능

- URL 요청 관리 : API 엔드포인트 URL을 입력하고 파라미터 설정을 통한 GET 요청
- 파라미터 관리 : 동적으로 파라미터 추가 및 삭제 가능
- 응답 처리 : JSON 및 XML 응답 자동 파싱
- 결과 시각화 : 데이터를 테이블 형식으로 표시
- 데이터 저장 : 결과를 CSV 파일로 저장 가능

### 개발 환경

- 언어 : Python 3.9.21
- 운영체제 : Windows 10
- 개발도구 : Visual Studio Code
- 가상환경 : Anaconda 가상환경


### 사용 방법
[Releases](https://github.com/sparky1543/openapi-call/releases)에서 실행 파일 다운로드 후 사용

![스크린샷1](https://github.com/user-attachments/assets/e74bdf81-97f1-4ed9-a488-6df4cd318265)
![스크린샷2](https://github.com/user-attachments/assets/1bdaa1aa-7cf7-4c2e-a2f4-fa12b7b53162)
![스크린샷3](https://github.com/user-attachments/assets/86096426-4527-436b-852e-d058de10720e)

[OpenAPI CALL 매뉴얼](https://github.com/sparky1543/openapi-call/blob/main/OpenAPI%20CALL%20%EB%A7%A4%EB%89%B4%EC%96%BC.pdf) 파일에서 더 자세한 사용 방법을 확인하실 수 있습니다.

#### 1. URL 및 파라미터 설정

- **요청 URL** : API 엔드포인트 URL 입력 (전체 URL 또는 기본 URL만 입력 가능)
- **요청변수 설정** : 체크박스 선택 시 URL의 파라미터가 자동으로 분리되어 표시
- **요청변수** : 서비스키, 페이지 번호 등 필요한 파라미터 입력  

#### 2. API 호출 및 결과 확인

- **조회하기** : 버튼 클릭으로 API 호출 실행
- **상태표시** : 준비/진행 중/완료/오류 등 현재 상태 표시
- **결과 데이터** : 테이블 형식으로 응답 데이터 확인

#### 3. 데이터 저장 및 활용

- **저장하기** : 조회된 데이터를 CSV 파일로 저장
- 저장 위치 지정 후 파일명 입력

#### 참고사항

- Encoding된 서비스키는 요청변수 설정 체크 시 자동으로 Decoding 형태로 변환
- 오류 발생 시 상세 메시지를 통해 문제 원인 확인 가능

### 문의 및 피드백

버그 제보 및 기능 개선 요청은 **Issues**에 남겨주세요.
