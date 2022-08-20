vcvarsall.bat x64
mkdir publish

rmdir /s /q "publish/conformu.macos-x64"
dotnet publish -c Debug -p:Platform="Any CPU" -r osx-x64 --framework net6.0 --self-contained true -o ./publish/conformu.macos-x64 -p:PublishSingleFile=true -p:PublishReadyToRunShowWarnings=true
bsdtar -cJf publish/conformu.macos-x64.tar.xz -C publish\conformu.macos-x64\ *

rmdir /s /q "publish/conformu.macos-arm64"
dotnet publish -c Debug -p:Platform="Any CPU" -r osx-arm64 --framework net6.0 --self-contained true -o ./publish/conformu.macos-arm64 -p:PublishSingleFile=true 
bsdtar -cJf publish/conformu.macos-arm64.tar.xz -C publish\conformu.macos-arm64\ *

echo "Builds complete"

pause