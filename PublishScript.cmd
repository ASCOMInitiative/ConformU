@echo off
echo *** Setup environment

rmdir /s /q "publish"
mkdir publish
mkdir publish\Temp

if defined __VCVARSALL_HOST_ARCH (
    echo __VCVARSALL_HOST_ARCH is set to "%__VCVARSALL_HOST_ARCH%"
) else (
    echo __VCVARSALL_HOST_ARCH is NOT set, setting environment
    call "C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
)

Rem Build ConformU so that the InformationVersion,txt file is up to date.
dotnet build ConformU/ConformU.csproj -c Debug -p:Platform="Any CPU" --framework net10.0

Rem Extract the full build number from the InformationVersion.txt file.
set /p FULL_VERSION=<"ConformU\bin\Any CPU\Debug\net10.0\InformationVersion.txt"

Rem Extract the version number from the full build number
for /f "tokens=1 delims=+" %%a in ("%FULL_VERSION%") do (
    set BUILD_VERSION=%%a
)

echo Version number: %BUILD_VERSION%, Build number: %FULL_VERSION%

cd
cd J:\ConformU

echo *** Publishing Linux builds
wsl -e bash -lc "/mnt/j/ConformU/build.bash %BUILD_VERSION%" 

echo *** Publishing MacOS Intel silicon
dotnet publish ConformU/ConformU.csproj -c Debug -p:Platform="Any CPU" -r osx-x64 --framework net10.0 --self-contained true /p:PublishTrimmed=false -p:PublishSingleFile=true -p:UseAppHost=true -o ./publish/Temp/ConformUOsxX64/
echo *** Completed MacOS Intel silicon

echo *** Publishing MacOS Apple silicon
dotnet publish ConformU/ConformU.csproj -c Debug -p:Platform="Any CPU" -r osx-arm64 --framework net10.0 --self-contained true /p:PublishTrimmed=false -p:PublishSingleFile=true -p:UseAppHost=true -o ./publish/Temp/ConformUOsxArm64/
echo *** Completed MacOS Apple silicon

echo *** Copying MacOS Intel silicon
scp -r J:\ConformU\publish\Temp\ConformUOsxX64\* petersimpson@macmini:/Users/petersimpson/builds/ConformU/source_files/ConformU.macos-x64

echo *** Copying MacOS Apple silicon
scp -r J:\ConformU\publish\Temp\ConformUOsxArm64\* petersimpson@macmini:/Users/petersimpson/builds/ConformU/source_files/ConformU.macos-arm64

echo *** Running make script
ssh -t petersimpson@macmini "sudo /Users/petersimpson/Builds/ConformU/makeconformu %BUILD_VERSION%"

echo *** Copying published DMG file
scp petersimpson@macmini:/Users/petersimpson/Builds/ConformU/dmg_file/* J:\ConformU\publish

echo *** Publishing Windows ARM 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-arm64 --framework net10.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/Temp/ConformUArm64/
echo *** Completed Windows ARM 64bit publish

echo *** Publishing Windows Intel 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x64 --framework net10.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/Temp/ConformUx64/
echo *** Completed Windows Intel 64bit publish

rem The Intel 32bit version serves on ARM64 as well because .NET doesn't support publishing 32bit Windows-Arm executables
echo *** Publishing Windows Intel 32bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x86 --framework net10.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/Temp/ConformUx86/
echo *** Completed Windows Intel 32bit publish

echo *** Creating Windows installer
cd J:\ConformU\Setup
"C:\Program Files (x86)\Inno Setup 6\iscc.exe" "ConformU.iss"
cd ..

echo *** Builds complete