echo *** Setup environment

rmdir /s /q "publish"
mkdir publish

call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd
cd J:\ConformU

echo *** Build application
MSBuild "J:\ConformU\ConformU.sln" /p:Configuration=Debug /p:Platform="Any CPU" /t:Restore 
cd
MSBuild "J:\ConformU\ConformU.sln" /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
echo *** Completed Build

echo *** Publishing Windows 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x64 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=false -o ./publish/ConformU64
echo ***Completed 64bit publish

echo *** Publishing Windows 32bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x86 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=false -o ./publish/ConformU86
echo *** Completed 32bit publish

echo *** Creating Windows installer
cd WindowsInstaller
"C:\Program Files (x86)\Inno Script Studio\isstudio.exe" -compile "J:\ConformU\WindowsInstaller\ConformU.iss"
cd ..

echo *** Builds complete

pause