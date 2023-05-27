@echo off
ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.camera
echo.
echo Return code: %errorlevel%
echo.
set CameraRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.covercalibrator
echo.
echo Return code: %errorlevel%
echo.
set CoverCalibratorRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.dome
echo.
echo Return code: %errorlevel%
echo.
set DomeRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.filterwheel
echo.
echo Return code: %errorlevel%
echo.
set FilterWheelRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.focuser
echo.
echo Return code: %errorlevel%
echo.
set FocuserRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.observingconditions
echo.
echo Return code: %errorlevel%
echo.
set ObservingConditionsRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.rotator
echo.
echo Return code: %errorlevel%
echo.
set RotatorRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.safetymonitor
echo.
echo Return code: %errorlevel%
echo.
set SafetyMonitorRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.switch
echo.
echo Return code: %errorlevel%
echo.
set SwitchRC=%errorlevel%

rem ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.telescope
echo.
echo Return code: %errorlevel%
echo.
set TelescopeRC=%errorlevel%

ConformU\bin\Debug\net7.0-windows\conformu.exe conformance ascom.simulator.video
echo.
echo Return code: %errorlevel%
echo.
set VideoRC=%errorlevel%

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
echo Video issues: %VideoRC%
echo.

pause

