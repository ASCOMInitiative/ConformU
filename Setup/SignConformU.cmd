rem @echo off
rem @echo ConformU Requested to sign %1
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
start "Sign with SHA256" /wait /b "signtool" sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1