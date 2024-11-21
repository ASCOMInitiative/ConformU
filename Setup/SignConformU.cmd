@echo on
call "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat" -startdir=none -arch=x64 -host_arch=x64
rem @echo off
rem @echo ConformU Requested to sign %1
rem call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
rem start "Sign with SHA256" /wait /b "signtool" sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1

"signtool" sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1

exit 0