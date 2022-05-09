@echo off
cd source

pyinstaller.exe ^
    --console ^
    --noconfirm ^
    --icon %cd%\icon.ico ^
	--workpath ..\build ^
    --specpath ..\build ^
	--distpath ..\distribution ^
	--add-binary %cd%\icon.ico;. ^
	--noupx ^
    .\CSV_to_PDF.py

cd ..
copy .\OtherFiles\Example.ini .\distribution\.

rem Die ist 34 MB groß und wird anscheined für die Anwendung nicht benötigt
del .\distribution\CSV_to_PDF\*gfortran-win_amd64.dll

pause
