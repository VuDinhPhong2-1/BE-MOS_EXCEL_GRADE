FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src

COPY . .
RUN dotnet restore MOS.ExcelGrading.sln
RUN dotnet publish MOS.ExcelGrading.API/MOS.ExcelGrading.API.csproj -c Release -o /app/publish

FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS final
WORKDIR /app
COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://+:${PORT}

CMD ["dotnet", "MOS.ExcelGrading.API.dll"]
