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
dotnet publish -c Debug /p:Platform="Any CPU" -r win-x64 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false -o ./publish/ConformU64
echo ***Completed 64bit publish

echo *** Signing Windows 64bit
rem signtool sign /a /as /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\ConformU64\*.dll"
rem signtool sign /a /as /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\ConformU64\*.exe"
echo ***Completed 64bit signing

echo *** Publishing Windows 32bit
dotnet publish -c Debug /p:Platform="Any CPU" -r win-x86 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false -o ./publish/ConformU86
echo *** Completed 32bit publish

echo *** Signing Windows 32bit
rem signtool sign /a /as /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\ConformU86\*.dll"
rem signtool sign /a /as /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\ConformU86\*.exe"
echo *** Completed 32bit signing

echo *** Creating Windows installer
cd setup
"C:\Program Files (x86)\Inno Script Studio\isstudio.exe" -compile "J:\ConformU\Setup\ConformU.iss"
cd ..

echo *** Publishing MacOS Intel silicon
dotnet publish -c Debug -p:Platform="Any CPU" -r osx-x64 --framework net7.0 --self-contained true -o ./publish/conformu.macos-x64 -p:PublishSingleFile=true -p:PublishReadyToRunShowWarnings=true
bsdtar -cJf publish/conformu.macos-x64.tar.xz -C publish\conformu.macos-x64\ *
echo *** Completed MacOS Intel silicon

echo *** Publishing MacOS Apple silicon
dotnet publish -c Debug -p:Platform="Any CPU" -r osx-arm64 --framework net7.0 --self-contained true -o ./publish/conformu.macos-arm64 -p:PublishSingleFile=true 
bsdtar -cJf publish/conformu.macos-arm64.tar.xz -C publish\conformu.macos-arm64\ *
echo *** Completed MacOS Apple silicon

echo *** Publishing Linux ARM32
dotnet publish -c Debug /p:Platform="Any CPU" -r linux-arm --framework net7.0 --self-contained true -o ./publish/conformu.linux-arm32 
bsdtar -cJf publish/conformu.linux-arm32.needsexec.tar.xz -C publish\conformu.linux-arm32 *
echo *** Completed Linux ARM32

echo *** Publishing Linux ARM64
dotnet publish -c Debug /p:Platform="Any CPU" -r linux-arm64 --framework net7.0 --self-contained true -o ./publish/conformu.linux-arm64
bsdtar -cJf publish/conformu.linux-arm64.needsexec.tar.xz -C publish\conformu.linux-arm64 *
echo *** Completed Linux ARM64

echo *** Publishing Linux X64
dotnet publish -c Debug /p:Platform="Any CPU" -r linux-x64 --framework net7.0 --self-contained true -o ./publish/conformu.linux-x64
bsdtar -cJf publish/conformu.linux-x64.needsexec.tar.xz -C publish\conformu.linux-x64\ *
echo *** Completed Linux X64

echo *** Builds complete

pause