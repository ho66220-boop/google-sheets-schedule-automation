# Refactoring Plan

현재 공개용 Apps Script는 `apps-script/Code.gs` 단일 파일로 관리됩니다. 이번 포트폴리오 보강에서는 동작 안정성을 우선해 코드 구조를 실제로 변경하지 않았습니다.

## 리팩토링을 보류한 이유

Google Apps Script는 여러 `.gs` 파일을 같은 전역 스코프로 합쳐 실행합니다. 파일을 나누는 것 자체는 가능하지만, 상수 위치 변경이나 함수명 변경이 발생하면 메뉴 실행, 출석 반영, 날짜 처리 등 기존 동작에 영향을 줄 수 있습니다.

이 저장소는 실제 운영 사례를 포트폴리오용으로 정리한 것이므로, 첫 공개 버전에서는 기능 변경 없이 문서화와 비식별 처리를 우선했습니다.

## 향후 권장 구조

```txt
apps-script/
├─ main.gs
├─ config.gs
├─ parser.gs
├─ scheduleGenerator.gs
├─ attendanceSync.gs
├─ conflictLogger.gs
└─ utils.gs
```

## 파일별 역할

| 파일 | 역할 |
|---|---|
| main.gs | 메뉴 생성, 실행 엔트리포인트 |
| config.gs | 시트명, 교실 목록, 보호값, 컬럼 위치 등 설정 |
| parser.gs | 수강기호와 수기 일정 코드 파싱 |
| scheduleGenerator.gs | 교실별 시간표 생성 |
| attendanceSync.gs | 출석 시트 자동 반영 |
| conflictLogger.gs | 충돌 기록 시트 생성 및 기록 |
| utils.gs | 날짜, 문자열, 학번, 좌석, 교시 정규화 유틸 |

## 리팩토링 시 주의사항

- 함수 이름과 인자는 변경하지 않습니다.
- Apps Script 메뉴에서 직접 호출되는 함수는 전역 함수로 유지합니다.
- 상수명을 바꾸는 경우 모든 참조를 함께 수정해야 합니다.
- 리팩토링 후 실제 Google Sheets 환경에서 메뉴 실행 테스트가 필요합니다.
- 기능 변경과 파일 분리는 별도 커밋으로 분리하는 것이 좋습니다.

## 권장 진행 순서

1. 현재 `Code.gs` 백업
2. 함수 목록과 호출 관계 정리
3. 상수와 유틸 함수부터 분리
4. 파서, 시간표 생성, 출석 반영 순서로 분리
5. Google Apps Script 편집기에서 전체 저장
6. 메뉴 실행 테스트
7. 시간표 생성, 충돌 기록, 출석 반영 테스트
