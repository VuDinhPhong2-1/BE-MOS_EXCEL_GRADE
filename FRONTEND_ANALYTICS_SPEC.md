# Frontend Analytics Spec

## Base

- Base URL: `/api`
- Auth: `Authorization: Bearer <accessToken>`
- Role: `Teacher` hoặc `Admin`
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

## 2) Weak Tasks (Top câu hay sai)

### Endpoint

`GET /api/analytics/class/{classId}/weak-tasks?projectEndpoint=project09&top=10`

- `projectEndpoint` optional
- `top` optional, default backend là `10`

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
    "projectEndpoint": "project09",
    "attemptCount": 80,
    "averagePercentage": 72.5,
    "passRate": 81.25
  },
  {
    "projectEndpoint": "project10",
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

## API Client Example (Axios)

```ts
import axios from "axios";

export const api = axios.create({ baseURL: "/api" });

api.interceptors.request.use((config) => {
  const token = localStorage.getItem("accessToken");
  if (token) config.headers.Authorization = `Bearer ${token}`;
  return config;
});

export const analyticsApi = {
  getClassOverview: (classId: string) =>
    api.get<ClassAnalyticsOverviewResponse>(`/analytics/class/${classId}/overview`),

  getWeakTasks: (classId: string, projectEndpoint?: string, top = 10) =>
    api.get<WeakTaskResponse[]>(`/analytics/class/${classId}/weak-tasks`, {
      params: { projectEndpoint, top }
    }),

  getProjectPerformance: (classId: string) =>
    api.get<ProjectPerformanceResponse[]>(`/analytics/class/${classId}/project-performance`)
};
```

## Suggested UI Blocks

1. KPI cards: `averagePercentage`, `passRate`, `warningRate`, `totalAttempts`
2. Bar chart: Top weak tasks (`failedRate`)
3. Combo chart: Project performance (`averagePercentage` line + `passRate` bar)
4. Filter: `classId`, `projectEndpoint`, `top`

## Data Collection Requirement

Để analytics có dữ liệu đúng theo lớp, frontend khi gọi:

`POST /api/grading/project09` (`multipart/form-data`)

ngoài `studentFile`, cần gửi thêm:

- `classId`
- `assignmentId`
- `studentId`
