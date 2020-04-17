FROM mcr.microsoft.com/dotnet/framework/sdk:4.8 AS build
WORKDIR /app

#Copy required components
COPY msg.Testframework_cs/*.csproj ./msg.Testframework_cs/
COPY msg.Testframework_cs/testdata/TestSet/Scenario_UnitTest/ExecuteFile.xlsx ./msg.Testframework_cs/testdata/TestSet/Scenario_UnitTest/
COPY msg.Testframework_cs/testdata/TestSet/Scenario_UnitTest/DummyTest.xlsx ./msg.Testframework_cs/testdata/TestSet/Scenario_UnitTest/
COPY . .

#Build msgfw2.0
WORKDIR /app/msg.Testframework_cs
RUN dotnet build msg.Testframework_cs.csproj --source /app/OfflineNugetPackages

FROM build AS testrunner
WORKDIR /app/msg.Testframework_cs
ENTRYPOINT ["dotnet", "test", "--no-build", "--logger:trx"]
#ENTRYPOINT ["dotnet", "test", "--logger:trx"]

FROM build AS test
WORKDIR /app/msg.Testframework_cs
RUN dotnet test --no-build --logger:trx