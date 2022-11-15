@echo on
@echo Setting up variables
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64

echo Signing installer %1

REM The delays provide time for the executable to be released by the previous process ready for the next to proceed
echo Starting 1 second wait
timeout /T 1
echo Finished wait
start "Sign with SHA256" /wait /b "signtool" sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1

rem pause

