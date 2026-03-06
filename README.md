# MOS.ExcelGrading

## Run API on VS Code (HTTPS)

### 1) What `dotnet restore` does

```powershell
dotnet restore
```

- Restores NuGet dependencies from all `.csproj` files.
- Must run after clone, branch switch, or package changes.

### 2) What `dotnet run --launch-profile https` does

```powershell
dotnet run --project .\MOS.ExcelGrading.API\MOS.ExcelGrading.API.csproj --launch-profile https
```

- Runs the API project explicitly.
- Uses the `https` profile from `launchSettings.json`.
- Default listening URLs:
  - `https://localhost:7223`
  - `http://localhost:5293`

## Quick start

```powershell
cd D:\All_Project\MOS.ExcelGrading
dotnet restore
dotnet build MOS.ExcelGrading.sln
dotnet run --project .\MOS.ExcelGrading.API\MOS.ExcelGrading.API.csproj --launch-profile https
```

## Required environment variables

`appsettings.json` no longer stores DB credentials. Set MongoDB connection string from environment:

```powershell
$env:MongoDbSettings__ConnectionString="your-mongodb-connection-string"
```

Optional if you need to override defaults:

```powershell
$env:MongoDbSettings__DatabaseName="MOS"
```

Select backend mode (`local` or `deploy`):

```powershell
$env:AppMode="local"
```

## Health check

- `https://localhost:7223/api/grading/health`

## Analytics APIs (Class insights)

After grading, backend now stores task-level snapshots for analytics.  
Grading endpoints now use only one upload file:

- `POST /api/grading/project01` (`multipart/form-data`)
- `POST /api/grading/project02` (`multipart/form-data`)
- `POST /api/grading/project03` (`multipart/form-data`) - Task 1-5 auto, Task 6 manual note
- `POST /api/grading/project04` (`multipart/form-data`)
- `POST /api/grading/project09` (`multipart/form-data`)

Required form field:

- `studentFile` (`.xlsx` or `.xlsm`)

Optional form fields:

- `classId`
- `assignmentId`
- `studentId`

New analytics endpoints:

- `GET /api/analytics/class/{classId}/overview`
- `GET /api/analytics/class/{classId}/weak-tasks?projectEndpoint=project09&top=10`
- `GET /api/analytics/class/{classId}/project-performance`

## Frontend config example (Vite)

```env
VITE_API_TARGET=local
VITE_API_LOCAL_URL=https://localhost:7223
VITE_API_DEPLOY_URL=https://be-mos-excel-grade.onrender.com
```

## Common issues

1. `Failed to fetch`
- Frontend calls `https://localhost:7223` but API was started with `http` profile.
- Start API with `--launch-profile https`.

2. HTTPS certificate issue

```powershell
dotnet dev-certs https --trust
```

3. Missing SDK

```powershell
dotnet --info
```

This project targets `net9.0`, so .NET SDK 9 is required.
