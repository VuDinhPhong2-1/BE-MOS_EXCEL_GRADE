# Frontend Analytics Spec

This file documents frontend-facing analytics response shapes and UI mapping notes. Keep canonical API contract details in `../API_CONTRACT.md` and the full endpoint inventory in `../API_ENDPOINTS_DETAILED.md`.

## Base

- Base URL: configured by `FRONTEND/src/config/api.ts`
- Protected requests: use `authFetch`
- Auth: `Authorization: Bearer <accessToken>`
- Roles: `Teacher` or `Admin`
- Permission claim: `grades.view`

## 1) Class Overview

### Endpoint

`GET /api/analytics/class/{classId}/overview`

### Response

```json
{
  "classId": "67b0e1e2e3f4a5b6c7d8e9f0",
  "totalAttempts": 120,
  "totalStudents": 38,
  "averagePercentage": 67.45,
  "passRate": 74.17,
  "warningRate": 12.5
}
```

### TypeScript

```ts
export interface ClassAnalyticsOverviewResponse {
  classId: string;
  totalAttempts: number;
  totalStudents: number;
  averagePercentage: number; // 0..100
  passRate: number; // 0..100
  warningRate: number; // 0..100
}
```

### Chart mapping example

```ts
export function mapOverviewToGaugeData(d: ClassAnalyticsOverviewResponse) {
  return [
    { label: "Average", value: d.averagePercentage },
    { label: "Pass Rate", value: d.passRate },
    { label: "Warning Rate", value: d.warningRate }
  ];
}
```

## 2) Weak Tasks

### Endpoint

`GET /api/analytics/class/{classId}/weak-tasks?projectEndpoint=excel/project09&top=10`

- `projectEndpoint` is optional.
- `top` is optional; backend default is `10`.
- Use canonical project endpoints such as `excel/project09` when filtering. Legacy values such as `project09` may appear in old persisted attempts and should be handled defensively in UI labels.

### Response

```json
[
  {
    "taskId": "P09T3",
    "taskName": "Task 3",
    "attemptCount": 80,
    "failedCount": 49,
    "failedRate": 61.25
  },
  {
    "taskId": "P09T5",
    "taskName": "Task 5",
    "attemptCount": 80,
    "failedCount": 40,
    "failedRate": 50
  }
]
```

### TypeScript

```ts
export interface WeakTaskResponse {
  taskId: string;
  taskName: string;
  attemptCount: number;
  failedCount: number;
  failedRate: number; // 0..100
}
```

### Chart mapping example

```ts
export function mapWeakTasksToBarChart(rows: WeakTaskResponse[]) {
  return rows.map((r) => ({
    x: r.taskId,
    y: r.failedRate,
    label: r.taskName,
    attempts: r.attemptCount,
    failed: r.failedCount
  }));
}
```

## 3) Project Performance

### Endpoint

`GET /api/analytics/class/{classId}/project-performance`

### Response

```json
[
  {
    "projectEndpoint": "excel/project09",
    "attemptCount": 80,
    "averagePercentage": 72.5,
    "passRate": 81.25
  },
  {
    "projectEndpoint": "excel/project10",
    "attemptCount": 40,
    "averagePercentage": 61.2,
    "passRate": 62.5
  }
]
```

### TypeScript

```ts
export interface ProjectPerformanceResponse {
  projectEndpoint: string;
  attemptCount: number;
  averagePercentage: number; // 0..100
  passRate: number; // 0..100
}
```

### Chart mapping example

```ts
export function mapProjectPerformanceToCombo(rows: ProjectPerformanceResponse[]) {
  return rows.map((r) => ({
    x: r.projectEndpoint,
    avgLine: r.averagePercentage,
    passBar: r.passRate,
    attempts: r.attemptCount
  }));
}
```

## API Client Example (`authFetch`)

```ts
import { authFetch } from "../context/AuthContext";
import { API_BASE_URL } from "../config/api";

async function getJson<T>(path: string): Promise<T> {
  const response = await authFetch(`${API_BASE_URL}${path}`);
  if (!response.ok) {
    throw new Error(`Analytics request failed: ${response.status}`);
  }

  return response.json() as Promise<T>;
}

export const analyticsApi = {
  getClassOverview: (classId: string) =>
    getJson<ClassAnalyticsOverviewResponse>(`/analytics/class/${classId}/overview`),

  getWeakTasks: (classId: string, projectEndpoint?: string, top = 10) => {
    const params = new URLSearchParams({ top: String(top) });
    if (projectEndpoint) params.set("projectEndpoint", projectEndpoint);

    return getJson<WeakTaskResponse[]>(
      `/analytics/class/${classId}/weak-tasks?${params.toString()}`
    );
  },

  getProjectPerformance: (classId: string) =>
    getJson<ProjectPerformanceResponse[]>(
      `/analytics/class/${classId}/project-performance`
    )
};
```

Adapt imports to the actual service/context locations in `FRONTEND/src`; do not introduce a parallel Axios client for protected API calls.

## Suggested UI Blocks

1. KPI cards: `averagePercentage`, `passRate`, `warningRate`, `totalAttempts`
2. Bar chart: top weak tasks (`failedRate`)
3. Combo chart: project performance (`averagePercentage` line + `passRate` bar)
4. Filter: `classId`, `projectEndpoint`, `top`

## Data Collection Requirement

Analytics depends on persisted grading attempts. Frontend grading flows that should feed analytics need to include context fields with grading requests and/or save scores through the score persistence workflow:

- `classId`
- `assignmentId`
- `studentId`

Example canonical direct grading route:

`POST /api/grading/excel/project09` (`multipart/form-data`)

Legacy Excel aliases such as `/api/grading/project09` remain supported for backward compatibility, but new docs and UI code should prefer canonical endpoint names.