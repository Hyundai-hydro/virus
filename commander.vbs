Set shell = CreateObject("WScript.Shell")

Do
    menu = "--- VM System Commander v1.6 [ULTIMATE] ---" & vbCrLf & _
           "1. NARZEDZIA XYZ (Deep Control/Performance)" & vbCrLf & _
           "2. GENERATOR BLEDOW (Error Lab)" & vbCrLf & _
           "3. ZAAWANSOWANE (RegEdit/GodMode/UAC)" & vbCrLf & _
           "4. RATUNEK (Kill All User Apps)" & vbCrLf & _
           "5. CZYSZCZENIE (Temp/Cache/DNS - BARDZO PRZYDATNE)" & vbCrLf & _
           "6. ZASILANIE (Shutdown/Restart/Logoff/Slide/BIOS + Czas+Tekst)" & vbCrLf & _
           "7. WYJDZ"

    wybor = InputBox(menu, "Commander v1.6", "1")
    If wybor = "" Or wybor = "7" Then Exit Do

    Select Case wybor
        Case "1"
            shell.Run "perfmon /res", 1

        Case "2"
            errMenu = "GENERATOR BLEDOW:" & vbCrLf & _
                      "1. Blad Krytyczny" & vbCrLf & _
                      "2. Ostrzezenie" & vbCrLf & _
                      "3. Informacja" & vbCrLf & _
                      "4. Pytanie"
            errWybor = InputBox(errMenu, "Error Generator", "1")
            If errWybor <> "" Then
                txt = InputBox("Tresc:", "Tresc", "Wystapil blad!")
                tit = InputBox("Tytul:", "Tytul", "System Error")
                Select Case errWybor
                    Case "1" : MsgBox txt, 16, tit
                    Case "2" : MsgBox txt, 48, tit
                    Case "3" : MsgBox txt, 64, tit
                    Case "4" : MsgBox txt, 32, tit
                End Select
            End If

        Case "3"
            advMenu = "OPCJE ZAAWANSOWANE:" & vbCrLf & _
                      "1. Otworz Edytor Rejestru" & vbCrLf & _
                      "2. Utworz GodMode na pulpicie" & vbCrLf & _
                      "3. Ustawienia UAC"
            advWybor = InputBox(advMenu, "Advanced", "1")
            Select Case advWybor
                Case "1" : shell.Run "regedit.exe", 1
                Case "2" : 
                    desktop = shell.SpecialFolders("Desktop")
                    shell.Run "cmd /c mkdir """ & desktop & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}""", 0
                Case "3" : shell.Run "UserAccountControlSettings.exe", 1
            End Select

        Case "4"
            shell.Run "taskkill /f /fi ""STATUS eq RUNNING"" /fi ""USERNAME ne NT AUTHORITY\SYSTEM""", 0
            shell.Run "explorer.exe", 1

        Case "5" ' === NOWE: CZYSZCZENIE SYSTEMU ===
            cleanConfirm = MsgBox("Czyszczenie usunie:" & vbCrLf & _
                                  "- %TEMP% (pliki tymczasowe użytkownika)" & vbCrLf & _
                                  "- %WINDIR%\Temp (pliki systemowe)" & vbCrLf & _
                                  "- %WINDIR%\Prefetch (cache uruchamiania)" & vbCrLf & _
                                  "- Cache DNS" & vbCrLf & vbCrLf & _
                                  "Kontynuować?", vbQuestion + vbYesNo, "Czyszczenie Systemu")
            If cleanConfirm = vbYes Then
                ' Uruchamia ukryte okno CMD, czyści foldery i czeka na zakończenie
                shell.Run "cmd /c ""del /f /s /q %TEMP%\*.* 2>nul & del /f /s /q %WINDIR%\Temp\*.* 2>nul & del /f /s /q %WINDIR%\Prefetch\*.* 2>nul & ipconfig /flushdns 2>nul""", 0, True
                MsgBox "Czyszczenie zakończone pomyślnie!" & vbCrLf & "Zwolniono miejsce i odświeżono DNS.", vbInformation, "Sukces"
            End If

        Case "6" ' === NOWE: ZASILANIE Z CZASEM I TEKSTEM ===
            powMenu = "ZASILANIE:" & vbCrLf & _
                      "1. Restart do BIOS (wymaga uprawnień admina)" & vbCrLf & _
                      "2. Wyloguj" & vbCrLf & _
                      "3. Uruchom ponownie" & vbCrLf & _
                      "4. Zamknij system" & vbCrLf & _
                      "5. Slide to Shut Down (Windows 10/11)" & vbCrLf & _
                      "6. Anuluj oczekujące zamknięcie"
            powWybor = InputBox(powMenu, "Zasilanie", "4")
            
            Select Case powWybor
                Case "1"
                    ' Restart do BIOS/UEFI (wymaga podniesienia uprawnień)
                    shell.Run "powershell -Command ""Start-Process shutdown -ArgumentList '/r /fw /t 0' -Verb RunAs""", 0
                    
                Case "2" ' Wyloguj (shutdown /l nie obsługuje /t, więc używamy timeout w CMD)
                    timeIn = InputBox("Opóźnienie wylogowania (sekundy):", "Czas", "10")
                    If timeIn = "" Or Not IsNumeric(timeIn) Then timeIn = "10"
                    shell.Run "cmd /c timeout /t " & timeIn & " /nobreak >nul && shutdown /l", 0
                    
                Case "3" ' Restart
                    timeIn = InputBox("Czas do restartu (sekundy):", "Czas", "30")
                    If timeIn = "" Or Not IsNumeric(timeIn) Then timeIn = "30"
                    msgIn = InputBox("Komunikat widoczny na ekranie:", "Tekst", "VM Commander restartuje system...")
                    If msgIn = "" Then msgIn = "VM Commander restartuje system..."
                    shell.Run "shutdown /r /t " & timeIn & " /c """ & msgIn & """", 0
                    
                Case "4" ' Zamknięcie
                    timeIn = InputBox("Czas do zamknięcia (sekundy):", "Czas", "30")
                    If timeIn = "" Or Not IsNumeric(timeIn) Then timeIn = "30"
                    msgIn = InputBox("Komunikat widoczny na ekranie:", "Tekst", "VM Commander zamyka system...")
                    If msgIn = "" Then msgIn = "VM Commander zamyka system..."
                    shell.Run "shutdown /s /t " & timeIn & " /c """ & msgIn & """", 0
                    
                Case "5" ' Slide to Shut Down
                    shell.Run "SlideToShutDown.exe", 1
                    
                Case "6" ' Anuluj
                    shell.Run "shutdown /a", 0
                    MsgBox "Zaplanowane zamknięcie/restart anulowane!", vbInformation, "Anulowano"
                    
                Case Else
                    MsgBox "Nieprawidłowy wybór!", vbExclamation, "Błąd"
            End Select

        Case Else
            MsgBox "Nieprawidłowy wybór!", vbExclamation, "Błąd"
    End Select
Loop