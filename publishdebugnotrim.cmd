call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64

mkdir publish

dotnet publish -c Debug /p:Platform="Any CPU" -r win-x64 --framework net6.0-windows --self-contained true /p:PublishTrimmed=false  -o ./publish/conformu.windows-x64
signtool sign /a /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-x64\*.dll"
signtool sign /a /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-x64\*.exe"

7z a publish/conformu-windows-x64.zip publish\conformu.windows-x64\ 
rmdir /s /q "publish/conformu.windows-x64"

dotnet publish -c Debug /p:Platform="Any CPU" -r linux-arm --framework net6.0 --self-contained true -o ./publish/conformu.linux-arm32
bsdtar -cJf publish/conformu.linux-arm32.tar.xz -C publish\conformu.linux-arm32 *
rmdir /s /q "publish/conformu.linux-arm32"

dotnet publish -c Debug /p:Platform="Any CPU" -r linux-arm64 --framework net6.0 --self-contained true -o ./publish/conformu.linux-aarch64
bsdtar -cJf publish/conformu.linux-aarch64.tar.xz -C publish\conformu.linux-aarch64 *
rmdir /s /q "publish/conformu.linux-aarch64"

dotnet publish -c Debug /p:Platform="Any CPU" -r linux-x64 --framework net6.0 --self-contained true -o ./publish/conformu.linux-x64
bsdtar -cJf publish/conformu.linux-x64.tar.xz -C publish\conformu.linux-x64\ *
rmdir /s /q "publish/conformu.linux-x64"

dotnet publish -c Debug /p:Platform="Any CPU" -r win-x86 --framework net6.0-windows --self-contained true /p:PublishTrimmed=false -o ./publish/conformu.windows-x86
signtool sign /a /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-x86\*.dll"
signtool sign /a /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-x86\*.exe"
7z a publish/conformu-windows-x86.zip publish\conformu.windows-x86\
rmdir /s /q "publish/conformu.windows-x86"

dotnet publish -c Debug /p:Platform="Any CPU" -r win-x86 --framework net6.0-windows --self-contained false /p:PublishTrimmed=false -o ./publish/conformu.windows-DotNet86
signtool sign /a /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-DotNet86\*.dll"
signtool sign /a /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-DotNet86\*.exe"
7z a publish/conformu-windows-DotNet86.zip publish\conformu.windows-DotNet86\
rmdir /s /q "publish/conformu.windows-DotNet86"

dotnet publish -c Debug /p:Platform="Any CPU" -r win-x64 --framework net6.0-windows --self-contained false /p:PublishTrimmed=false -o ./publish/conformu.windows-DotNet64
signtool sign /a /v /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-DotNet64\*.dll"
signtool sign /a /fd SHA256 /tr http://time.certum.pl /td SHA256 "publish\conformu.windows-DotNet64\*.exe"
7z a publish/conformu-windows-DotNet64.zip publish\conformu.windows-DotNet64\
rmdir /s /q "publish/conformu.windows-DotNet64"

dotnet publish -c Debug /p:Platform="Any CPU" -r osx-x64 --framework net6.0 --self-contained true /p:AppImage=true -o ./publish/conformu.macos-x64
bsdtar -cJf publish/conformu.macos-x64.tar.xz -C publish\conformu.macos-x64\ *
rmdir /s /q "publish/conformu.macos-x64"

echo "Builds complete"

pause