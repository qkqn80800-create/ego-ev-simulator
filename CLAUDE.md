# EV 충전기 시공 시뮬레이터

## 프로젝트 개요
EV 충전기 설치 비용을 시뮬레이션하는 웹 앱. 로그인 없이 사용, 데이터는 브라우저 localStorage에만 저장.

**로컬 경로:** `C:\ev_simulation_app`  
**GitHub:** `https://github.com/qkqn80800-create/ego-ev-simulator` (브랜치: `main`)  
**개발 서버:** `http://localhost:5175`

---

## 아키텍처

- **프레임워크:** React + TypeScript + Vite
- **메인 파일:** `frontend/src/SimDashboard.tsx` — UI 전체 (모든 컴포넌트 포함)
- **계산 엔진:** `frontend/src/engine.ts`
- **타입 정의:** `frontend/src/types.ts`
- **데이터 저장:** 브라우저 `localStorage` (`sim_global_defaults` 키)
- **서버 없음** — 순수 정적 SPA

---

## 주요 기능

### 시뮬레이션 파라미터
- 충전기 구성 (완속 7kW, 급속 50/100/240kW, 수량)
- 비용 설정 (공사비, 부대비용)
- 기간·성장률 (운영 기간, 연간 성장률)
- 고객 직접 납부 항목

### 추가 발생 비용 (ExtraCostsModal)
아코디언 UI (5개 항목, 기본 접힘):
1. 한전 시설부담금 (kepco)
2. 사용전검사 (safety)
3. 전기안전관리대행비 (elec_safety)
4. 보험료 (insurance)
5. 기타 안내 (extra_notice)

### 담당자 관리
- 기본 설정값 모달 → "⑤ 담당자 관리" 탭
- 업체별 담당자 등록 구조 (`CompanyEntry` 타입)
- 등록된 담당자만 "담당자 확인" 필드 통과 가능
- 담당자 미등록 시 누구든 입력 가능 (하위 호환)

### 기본 설정값 (GlobalUpdateModal)
탭 구조: ① 충전기 구성 / ② 비용 설정 / ③ 기간·성장률 / ④ 고객 직접 납부 / ⑤ 담당자 관리  
localStorage의 `sim_global_defaults` 키에 저장.

---

## 개발 서버 실행

```bash
cd C:\ev_simulation_app\frontend
npm run dev
```
→ `http://localhost:5175` 에서 실행

---

## 빌드 (배포용 정적 파일 생성)

```bash
cd C:\ev_simulation_app\frontend
npm run build
```
→ `frontend/dist/` 에 정적 파일 생성

---

## 중요 규칙

### React Hooks 규칙
- 훅은 반드시 함수 컴포넌트의 최상위에서 호출
- IIFE나 콜백 내부에서 useState/useEffect 사용 금지
- 복잡한 모달 UI는 별도 컴포넌트로 분리할 것 (예: `ExtraCostsModal`)

### Git 작업
- 코드 변경 후 `git add + commit + push` 까지 실행
- 브랜치: `main` (master 아님)

### 배포
- **사용자가 명시적으로 요청할 때만 배포**
- 오라클 서버 연결 방안은 아직 검토 중 (보안 방식 미결정)

---

## 타입 구조 (주요)

```typescript
type CompanyEntry = { id: string; name: string; managers: string[] }

type GlobalDefaults = Partial<SimParams> & {
  vkw_presets?: { label: string; kwh: number }[]
  custom_charger_types?: ChargerTypeEntry[]
  companies?: CompanyEntry[]
}
```

---

## 오라클 서버 연동 (미결)
- 보안 방식(IP 제한 / Basic Auth / 이고 충전관리 연동) 검토 중
- 결정 전까지 로컬 개발만 진행
