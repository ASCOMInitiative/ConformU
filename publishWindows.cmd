echo *** Setup environment

rmdir /s /q "publish"
mkdir publish

call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd
cd J:\ConformU

echo *** Build application
MSBuild "ConformU.sln" /p:Configuration=Debug /p:Platform="Any CPU" /t:Restore 
cd
MSBuild "ConformU.sln" /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
echo *** Completed Build

echo *** Publishing Windows ARM 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-arm64 --framework net8.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=false -o ./publish/ConformUArm64/
echo *** Completed Windows ARM 64bit publish

echo *** Publishing Windows Intel 64bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x64   --framework net8.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/ConformUx64/
echo *** Completed Windows Intel 64bit publish

rem The Intel 32bit version serves on ARM64 as well because .NET doesn't support publishing 32bit Windows-Arm executables
echo *** Publishing Windows Intel 32bit
dotnet publish ConformU/ConformU.csproj -c Debug /p:Platform="Any CPU" -r win-x86   --framework net8.0-windows --self-contained true /p:PublishTrimmed=false /p:PublishSingleFile=true -o ./publish/ConformUx86/
echo *** Completed Windows Intel 32bit publish

editbin /LARGEADDRESSAWARE .\publish\ConformUx86\conformu.exe
echo *** Completed setting large address aware flag on 32bit EXE

echo *** Creating Windows installer
cd Setup
"C:\Program Files (x86)\Inno Script Studio\isstudio.exe" -compile "conformu.iss"
cd ..

echo *** Builds complete

pause