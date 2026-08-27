echo *** Setup environment

rmdir /s /q "publish"
mkdir publish

call "C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd
cd J:\ConformU

echo *** Build application
MSBuild ConformU.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Restore 
cd
MSBuild ConformU.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
echo *** Completed Build

echo *** Publishing Linux ARM32
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r linux-arm --framework net10.0 --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true /p:UseAppHost=true /p:PublishReadyToRun=true -o ./publish/LinuxArm32/
rem bsdtar -cJf publish/conformu.linux-arm32.needsexec.tar.xz -C ConformU\bin\Debug\net10.0\linux-arm\publish\ *
echo *** Completed Linux ARM32

echo *** Publishing Linux ARM64
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r linux-arm64 --framework net10.0 --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true /p:UseAppHost=true /p:PublishReadyToRun=true -o ./publish/LinuxArm64/
rem bsdtar -cJf publish/conformu.linux-arm64.needsexec.tar.xz -C ConformU\bin\Debug\net10.0\linux-arm64\publish\ *
echo *** Completed Linux ARM64

echo *** Completed 64bit publish

echo *** Builds complete
pause