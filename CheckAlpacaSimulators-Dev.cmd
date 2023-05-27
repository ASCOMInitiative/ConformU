@echo off
echo %CD%\logs
ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/camera/0" --logfile logs\%1.camera.txt
echo.
echo Return code: %errorlevel%
echo.
set CameraRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/covercalibrator/0"
echo.
echo Return code: %errorlevel%
echo.
set CoverCalibratorRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/dome/0"
echo.
echo Return code: %errorlevel%
echo.
set DomeRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/filterwheel/0"
echo.
echo Return code: %errorlevel%
echo.
set FilterWheelRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/focuser/0"
echo.
echo Return code: %errorlevel%
echo.
set FocuserRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/observingconditions/0"
echo.
echo Return code: %errorlevel%
echo.
set ObservingConditionsRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/rotator/0"
echo.
echo Return code: %errorlevel%
echo.
set RotatorRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/safetymonitor/0"
echo.
echo Return code: %errorlevel%
echo.
set SafetyMonitorRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/switch/0"
echo.
echo Return code: %errorlevel%
echo.
set SwitchRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe %1 "http://%2/api/v1/telescope/0"
echo.
echo Return code: %errorlevel%
echo.
set TelescopeRC=%errorlevel%

echo Camera issues: %CameraRC%
echo CoverCalibrator issues: %CoverCalibratorRC%
echo Dome issues: %DomeRC%
echo FilterWheel issues: %FilterWheelRC%
echo Focuser issues: %FocuserRC%
echo ObservingConditions issues: %ObservingConditionsRC%
echo Rotator issues: %RotatorRC%
echo SafetyMonitor issues: %SafetyMonitorRC%
echo Switch issues: %SwitchRC%
echo Telescope issues: %TelescopeRC%
echo.

pause