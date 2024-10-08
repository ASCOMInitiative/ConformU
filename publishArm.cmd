echo *** Setup environment

call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd
cd J:\ConformU

echo *** Build application
MSBuild ConformU.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Restore 
cd
MSBuild ConformU.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
echo *** Completed Build

echo *** Publishing Windows ARM 32bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-arm32 --framework net8.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/ConformUArm32/
echo *** Completed 32bit publish

echo *** Publishing Windows ARM 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-arm64 --framework net8.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/ConformUArm64/
echo *** Completed 64bit publish

echo *** Builds complete
pause