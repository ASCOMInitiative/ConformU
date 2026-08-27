echo *** Setup environment

rmdir /s /q "publish"
mkdir publish

call "C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd
cd J:\ConformU

echo *** Publishing Linux X64
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r linux-x64 --framework net10.0 --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true /p:UseAppHost=true /p:PublishReadyToRun=true
bsdtar -cJf publish/conformu.linux-x64.needsexec.tar.xz -C "ConformU\bin\Any CPU\Debug\net10.0\linux-x64\publish" .
echo *** Completed Linux X64

echo *** Builds complete
pause