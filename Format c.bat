@echo off
mode con: cols=100 lines=30
title Microsoft Windows Disk Management - Format Utility
color 0f
cls

echo ===============================================================================
echo                NARZEDZIE DO NAPRAWY I FORMATOWANIA DYSKOW             
echo ===============================================================================
echo.
echo WYKRYTO KRYTYCZNY BLAD STRUKTURY PLIKOW NA PARTYJI C:
echo.
set /p admin="Czy uruchomic w trybie Administratora? (T/N): "
echo.
echo [!] OSTRZEZENIE: Kontynuacja spowoduje calkowite wymazanie danych.
echo.
set /p odp="Czy chcesz rozpoczac procedure formatowania i reinstalacji? (T/N): "

if /I "%odp%" NEQ "T" (
    echo.
    echo Operacja przerwana przez uzytkownika.
    timeout /t 3 >nul
    exit
)

cls
color 07
echo Inicjowanie procedury...
timeout /t 2 >nul

echo.
echo Analiza sektorow dysku C:...
:: Szybkie skanowanie plików dla efektu
dir /b /s C:\Windows\System32
cls

echo ===============================================================================
echo                FORMATOWANIE PARTYCJI SYSTEMOWEJ C: [NTFS]
echo ===============================================================================
echo.

:: Pętla postępu formatowania
for /l %%p in (0,10,100) do (
    echo Trwa formatowanie: %%p%% ukonczono...
    timeout /t 1 >nul
)

echo.
echo [ OK ] Struktura partycji wyczyszczona pomyslnie.
echo.
timeout /t 2 >nul

echo Usuwanie wpisow rejestru i czyszczenie tablicy MFT...
echo usuwanie: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet... OK
timeout /t 1 >nul
echo usuwanie: HKEY_CLASSES_ROOT\*... OK
timeout /t 1 >nul

cls
color 0c
echo.
echo ###############################################################################
echo #                                                                             #
echo #        SYSTEM WINDOWS ZOSTAL POMYSLNIE USUNIETY Z TEGO KOMPUTERA.          #
echo #                                                                             #
echo ###############################################################################
echo.

for /l %%i in (5,-1,1) do (
    echo Finalizacja procesow... pobranie pliku .ISO skonczy sie za: %%i...
    timeout /t 1 >nul
)

:: Wybór finalnej akcji na samym końcu:
if /I "%admin%"=="T" (
    :: Restart do BIOS (wymaga uprawnień administratora)
    shutdown /r /fw /t 0
) else (
    :: Zwykłe wyłączenie komputera
    shutdown /s /t 0
)
