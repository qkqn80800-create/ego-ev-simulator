# EV 충전기 시공 시뮬레이터

## 프로젝트 개요
EV 충전기 설치 비용을 시뮬레이션하는 웹 앱. 로그인 없이 사용, 데이터는 브라우저 localStorage에만 저장.

**로컬 경로:** `C:\ev_simulation_app`  
**GitHub:** `https://github.com/qkqn80800-create/ego-ev-simulator` (브랜치: `main`)  
**개발 서버:** `http://localhost:5175`  
**배포 주소:** `https://ego-ev-sim.surge.sh` (surge.sh · 공개 · 로그인 없음)

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

## 빌드 · 배포

배포처는 **surge.sh 단독**이다. 오라클 서버에는 올리지 않는다.

```bash
cd C:\ev_simulation_app\frontend
npm run build
cp dist/index.html dist/200.html     # surge SPA fallback (vercel.json 은 surge 가 읽지 않음)
npx surge dist ego-ev-sim.surge.sh
```

- 인증은 `~/.netrc` 의 `machine surge.surge.sh` 항목 사용 — 별도 로그인 불필요
- `npm run build` 는 `tsc -b` 를 먼저 돌린다. **타입 오류가 하나라도 있으면 빌드가 통째로 멈추고,**
  **그러면 배포본이 조용히 구버전으로 남는다.** (2026-09-14 실제 발생: TS2322 로 손익분기 kWh 탭이 미배포)
- 배포 후에는 실제 URL 의 번들을 받아 새 기능이 들어갔는지 확인할 것.

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
- 배포처는 surge.sh 뿐 — 오라클 서버에는 배포하지 않는다

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

## 운영 방식 (확정)
- **surge.sh 단독 배포.** 오라클 서버 연동은 하지 않기로 결정했다.
- 로그인·인증 없음. 주소를 아는 사람은 누구나 접속 가능 (고객에게 링크 전달 가능).
- 입력값은 서버에 저장되지 않고 브라우저 `localStorage`(`sim_global_defaults`)에만 남는다.
  → PC 가 바뀌면 기본 설정값도 없다.
- 사내 안내서: https://claude.ai/code/artifact/935594f6-2e02-4c9f-aac0-8d08a43fc234
