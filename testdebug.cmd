rem call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64

MSBuild "j:\ConformU\ConformU.sln" /p:Configuration=Debug /p:Platform="Any CPU" /t:Restore 
MSBuild "j:\ConformU\ConformU.sln" /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
echo ***Setup environment
rmdir /s /q "publish"
mkdir publish

cd publish
echo ***Completed Build
cd
dotnet publish "..\conformu.sln"  -c Debug /p:Platform="Any CPU" -r win-x86 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false -o ./ConformU86
echo ***Completed 32bit publish
signtool sign /a /as /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "ConformU86\*.dll"
signtool sign /a /as /fd SHA256 /tr http://time.certum.pl /td SHA256 "ConformU86\*.exe"
echo ***Completed 32bit signing


dotnet publish "..\conformu.sln"  -c Debug /p:Platform="Any CPU" -r win-x64 --framework net7.0-windows --self-contained true /p:PublishTrimmed=false -o ./ConformU64
echo ***Completed 64bit publish
signtool sign /a /as /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "ConformU64\*.dll"
signtool sign /a /as /fd SHA256 /tr http://time.certum.pl /td SHA256 "ConformU64\*.exe"
echo ***Completed 64bit signing


cd ..\setup

"C:\Program Files (x86)\Inno Script Studio\isstudio.exe" -compile "J:\ConformU\Setup\ConformU.iss"

cd ..