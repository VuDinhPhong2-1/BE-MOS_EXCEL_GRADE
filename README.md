# MOS Project Backend

ASP.NET Core Web API backend for MOS Project.

## Stack

- .NET 9
- ASP.NET Core Web API
- MongoDB + GridFS
- JWT authentication
- EPPlus for Excel grading
- Google Sheets API integration
- Optional Redis cache with in-memory fallback

## Directory

Run backend commands from:

```powershell
cd BACKEND
```

## Restore and build

```powershell
dotnet restore
dotnet build MOS.ExcelGrading.sln
```

## Required configuration

`appsettings.json` should not contain real credentials. Prefer environment variables or local user secrets.

Required values:

```powershell
$env:AppMode="local" # use "deploy" on Render production
$env:MongoDbSettings__ConnectionString="mongodb://localhost:27017"
$env:MongoDbSettings__DatabaseName="MOS"
$env:JwtSettings__SecretKey="<set-a-strong-secret>"
$env:JwtSettings__Issuer="MOS.ExcelGrading"
$env:JwtSettings__Audience="MOS.ExcelGrading.Users"
```

Optional values:

```powershell
$env:GoogleAuth__ClientId="<google-web-client-id>"
$env:GoogleSheets__Enabled="true"
$env:GoogleSheets__ServiceAccountJson="<service-account-json>"
$env:GoogleSheets__ServiceAccountJsonPath="<path-to-service-account-json>"
$env:Redis__Enabled="false"
$env:Redis__ConnectionString="<redis-connection-string>"
```

Do not commit real secrets, service account JSON, tokens, or production connection strings.

Current production domains:

- Frontend: `https://mos-grader-app.info.vn`
- Backend: `https://api.mos-grader-app.info.vn`

## Run API locally

```powershell
dotnet run --project .\MOS.ExcelGrading.API\MOS.ExcelGrading.API.csproj --launch-profile https
```

Default local URLs:

- `https://localhost:7223`
- `http://localhost:5293`

## Health check

```powershell
curl.exe https://localhost:7223/api/grading/health
```

## API documentation

Keep this README backend-focused and avoid duplicating full endpoint inventories here.

Canonical API docs:

- `../API_CONTRACT.md` - stable request/response contracts and compatibility rules
- `../API_ENDPOINTS_DETAILED.md` - controller endpoint inventory
- `../API_QUICK_REFERENCE.md` - ready-to-run examples

Important grading route compatibility:

- Excel canonical routes: `/api/grading/excel/project01..16,18,20,22`
- Excel legacy aliases: `/api/grading/projectXX`
- Word dynamic route: `/api/grading/word/{projectCode}`

## Common issues

### `Failed to fetch` from frontend

- Start the API with the `https` launch profile.
- Confirm frontend env points to `https://localhost:7223`.
- Trust the local development HTTPS certificate if needed.

```powershell
dotnet dev-certs https --trust
```

### Missing .NET SDK

```powershell
dotnet --info
```

This backend targets `net9.0`, so .NET SDK 9 is required.

### Empty analytics

Analytics depends on persisted grading attempts. Direct grading and score-save flows can differ; see `../BUGS_AND_PATCHES.md` before changing persistence behavior.

## Verification

```powershell
dotnet build MOS.ExcelGrading.sln
```
