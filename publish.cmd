vcvarsall.bat x64
mkdir publish

dotnet publish -c Release /p:Platform="Any CPU" -r linux-arm64 --framework net5.0 --self-contained true /p:PublishTrimmed=true /p:PublishSingleFile=true -o ./bin/conformu.linux-aarch64
bsdtar -cJf publish/conformu.linux-aarch64.tar.xz -C bin\conformu.linux-aarch64 *

dotnet publish -c Release /p:Platform="Any CPU" -r linux-x64 --framework net5.0 --self-contained true /p:PublishTrimmed=true /p:PublishSingleFile=true -o ./bin/conformu.linux-x64
bsdtar -cJf publish/conformu.linux-x64.tar.xz -C bin\conformu.linux-x64\ *

dotnet publish -c Release /p:Platform="Any CPU" -r win-x64 --framework net5.0-windows --self-contained true /p:PublishTrimmed=true /p:PublishSingleFile=true -o ./bin/conformu.windows-x64
7z a publish/conformu-windows-x64.zip bin\conformu.windows-x64\ 

dotnet publish -c Release /p:Platform="Any CPU" -r win-x86 --framework net5.0-windows --self-contained true /p:PublishTrimmed=true /p:PublishSingleFile=true -o ./bin/conformu.windows-x86
7z a publish/conformu-windows-x86.zip bin\conformu.windows-x86\

dotnet publish -c Release /p:Platform="Any CPU" -r osx-x64 --framework net5.0 --self-contained true /p:PublishTrimmed=true /p:PublishSingleFile=true -o ./bin/conformu.macos-x64
bsdtar -cJf publish/conformu.macos-x64.tar.xz -C bin\conformu.macos-x64\ *

echo "Note, these builds are not Ready to Run so they will run slower"
echo "Builds complete"

pause