echo *** Setup environment

rmdir /s /q "publish"
mkdir publish

call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd
cd J:\ConformU

echo *** Build application
MSBuild ConformU.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Restore 
cd
MSBuild ConformU.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
echo *** Completed Build

echo *** Publishing MacOS Intel silicon
dotnet publish -c Debug -p:Platform="Any CPU" -r osx-x64 --framework net7.0 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRunShowWarnings=true
bsdtar -cJf publish/conformu.macos-x64.tar.xz -C ConformU\bin\Debug\net7.0\osx-x64\publish\ *
echo *** Completed MacOS Intel silicon

echo *** Publishing MacOS Apple silicon
dotnet publish -c Debug -p:Platform="Any CPU" -r osx-arm64 --framework net7.0 --self-contained true -p:PublishSingleFile=true 
bsdtar -cJf publish/conformu.macos-arm64.tar.xz -C ConformU\bin\Debug\net7.0\osx-arm64\publish\ *
echo *** Completed MacOS Apple silicon

echo *** Publishing Linux ARM32
dotnet publish -c Debug /p:Platform="Any CPU" -r linux-arm --framework net7.0 --self-contained true /p:PublishSingleFile=true 
bsdtar -cJf publish/conformu.linux-arm32.needsexec.tar.xz -C ConformU\bin\Debug\net7.0\linux-arm\publish\ *
echo *** Completed Linux ARM32

echo *** Publishing Linux ARM64
dotnet publish -c Debug /p:Platform="Any CPU" -r linux-arm64 --framework net7.0 --self-contained true /p:PublishSingleFile=true
bsdtar -cJf publish/conformu.linux-arm64.needsexec.tar.xz -C ConformU\bin\Debug\net7.0\linux-arm64\publish\ *
echo *** Completed Linux ARM64

echo *** Publishing Linux X64
dotnet publish -c Debug /p:Platform="Any CPU" -r linux-x64 --framework net7.0 --self-contained true /p:PublishSingleFile=true
bsdtar -cJf publish/conformu.linux-x64.needsexec.tar.xz -C ConformU\bin\Debug\net7.0\linux-x64\publish\ *
echo *** Completed Linux X64

echo *** Publishing Windows 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x64 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/ConformU64/
echo *** Completed 64bit publish

echo *** Publishing Windows 32bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x86 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/ConformU86/
echo *** Completed 32bit publish

editbin /LARGEADDRESSAWARE ./publish/ConformU86/Conformu.exe
echo *** Completed setting large address aware flag on 32bit EXE

echo *** Creating Windows installer
cd WindowsInstaller
"C:\Program Files (x86)\Inno Script Studio\isstudio.exe" -compile "conformu.iss"
cd ..

echo *** Builds complete
pause