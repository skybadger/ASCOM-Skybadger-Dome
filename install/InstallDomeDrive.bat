xcopy ..\..\dome "h:\Program Files\Common Files\ASCOM\Dome\source\skybadger" /S /E /I

"h:\Program Files\Common Files\ASCOM\Dome\skybadgerdome.exe" /unregserver
copy "skybadgerdome.exe" "h:\Program Files\Common Files\ASCOM\Dome\"
copy "Vbhlp32.dll" "h:\Program Files\Common Files\ASCOM\Dome\"
copy "usbi2cio.dll" "h:\Program Files\Common Files\ASCOM\Dome\"
copy "astro32.dll" "h:\Program Files\Common Files\ASCOM\Dome\"
"h:\Program Files\Common Files\ASCOM\Dome\skybadgerdome.exe" /regserver

ASCOMDomeReg.vbs

pause
