Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Dim TITLE : TITLE = "VM COMMANDER v7.0 [ASCII PL]"
Dim SEP : SEP = String(48, "=")
Randomize

Do
    mainMenu = SEP & vbCrLf & _
               "        " & TITLE & vbCrLf & SEP & vbCrLf & _
               "ZAKLADKI GLOWNE:" & vbCrLf & _
               "[1] SYSTEM    [2] SIEC       [3] ZASILANIE" & vbCrLf & _
               "[4] NARZEDZIA [5] ZABAWA     [6] WYJDZ" & vbCrLf & SEP & vbCrLf & _
               "WYBIERZ ZAKLADKE (1-6):"

    tabWybor = InputBox(mainMenu, TITLE, "1")
    If tabWybor = "" Or tabWybor = "6" Then Exit Do

    Select Case tabWybor
        Case "1" : ShowSystemMenu
        Case "2" : ShowNetworkMenu
        Case "3" : ShowPowerMenu
        Case "4" : ShowToolsMenu
        Case "5" : ShowFunMenu
        Case Else : MsgBox "Nieprawidlowy wybor zakladki.", vbExclamation, "Blad"
    End Select
Loop
MsgBox "Do zobaczenia!", vbInformation, "Wyjscie"
WScript.Quit

' ===== ZAKLADKI =====

Sub ShowSystemMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: SYSTEM" & vbCrLf & SEP & vbCrLf & _
               "1. Menedzer Zadan" & vbCrLf & _
               "2. Tryb Gracza (Zabija tlo)" & vbCrLf & _
               "3. Oczyszczanie Systemu" & vbCrLf & _
               "4. Monitor Zasobow (CPU/RAM/Dysk)" & vbCrLf & _
               "5. Informacje o Sprzecie" & vbCrLf & _
               "6. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "System", "1")
        If wyb = "" Or wyb = "6" Then Exit Do
        Select Case wyb
            Case "1" : shell.Run "taskmgr", 1
            Case "2" : RunGamerMode
            Case "3" : RunCleaner
            Case "4" : RunResourceMonitor
            Case "5" : RunHardwareInfo
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub ShowNetworkMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: SIEC" & vbCrLf & SEP & vbCrLf & _
               "1. Narzedzia Sieciowe (IP/Ping/Reset)" & vbCrLf & _
               "2. Pokaz Hasla WiFi" & vbCrLf & _
               "3. Menadzer Defendera" & vbCrLf & _
               "4. Test Opoznienia Sieci (20x ping)" & vbCrLf & _
               "5. Sprawdz Otwarte Porty" & vbCrLf & _
               "6. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Siec", "1")
        If wyb = "" Or wyb = "6" Then Exit Do
        Select Case wyb
            Case "1" : RunNetwork
            Case "2" : RunWiFiPass
            Case "3" : RunDefender
            Case "4" : shell.Run "cmd /k ping -n 20 1.1.1.1", 1
            Case "5" : shell.Run "cmd /k netstat -ano | findstr LISTENING", 1
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub ShowPowerMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: ZASILANIE" & vbCrLf & SEP & vbCrLf & _
               "1. Restart do BIOS/UEFI" & vbCrLf & _
               "2. Wyloguj" & vbCrLf & _
               "3. Restart" & vbCrLf & _
               "4. Zamknij" & vbCrLf & _
               "5. Slide to Shut Down" & vbCrLf & _
               "6. Anuluj Akcje" & vbCrLf & _
               "7. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Zasilanie", "4")
        If wyb = "" Or wyb = "7" Then Exit Do
        Select Case wyb
            Case "1" : shell.Run "powershell -Command ""Start-Process shutdown -ArgumentList '/r /fw /t 0' -Verb RunAs""", 0
            Case "2"
                t = InputBox("Opoznienie (s):", "Czas", "10")
                If t = "" Or Not IsNumeric(t) Then t = "10"
                shell.Run "cmd /c timeout /t " & t & " /nobreak >nul && shutdown /l", 0
            Case "3"
                t = InputBox("Opoznienie (s):", "Czas", "30")
                If t = "" Or Not IsNumeric(t) Then t = "30"
                m = InputBox("Tekst:", "Komunikat", "Restartuje...")
                If m = "" Then m = "Restartuje..."
                shell.Run "shutdown /r /t " & t & " /c """ & m & """", 0
            Case "4"
                t = InputBox("Opoznienie (s):", "Czas", "30")
                If t = "" Or Not IsNumeric(t) Then t = "30"
                m = InputBox("Tekst:", "Komunikat", "Zamykam...")
                If m = "" Then m = "Zamykam..."
                shell.Run "shutdown /s /t " & t & " /c """ & m & """", 0
            Case "5" : shell.Run "SlideToShutDown.exe", 1
            Case "6"
                shell.Run "shutdown /a", 0
                MsgBox "Akcja anulowana!", vbInformation, "Anulowano"
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub ShowToolsMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: NARZEDZIA" & vbCrLf & SEP & vbCrLf & _
               "1. Edytor Rejestru / GodMode / UAC" & vbCrLf & _
               "2. Narzedzia Naprawcze (SFC/DISM)" & vbCrLf & _
               "3. Szybkie Narzedzia (CMD/PS/Uslugi)" & vbCrLf & _
               "4. Opcje Dodatkowe (Motyw/Schowek/Kosz)" & vbCrLf & _
               "5. Generator Bezpiecznych Hasel" & vbCrLf & _
               "6. Menedzer Autostartu" & vbCrLf & _
               "7. Zarzadzanie Dyskami" & vbCrLf & _
               "8. Lista Programow i Funkcji" & vbCrLf & _
               "9. Logi Systemowe" & vbCrLf & _
               "10. Ustawienia Dzwieku" & vbCrLf & _
               "11. Planer Zadan Systemowych" & vbCrLf & _
               "12. Podglad Top 10 Procesow (RAM)" & vbCrLf & _
               "13. Raport Stanu Baterii" & vbCrLf & _
               "14. Sterowanie Diodami Klawiatury" & vbCrLf & _
               "15. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Narzedzia", "1")
        If wyb = "" Or wyb = "15" Then Exit Do
        Select Case wyb
            Case "1" : RunAdvanced
            Case "2" : RunRepair
            Case "3" : RunQuickTools
            Case "4" : RunExtras
            Case "5" : RunPasswordGen
            Case "6" : RunStartupManager
            Case "7" : shell.Run "diskmgmt.msc", 1
            Case "8" : shell.Run "appwiz.cpl", 1
            Case "9" : shell.Run "eventvwr.msc", 1
            Case "10": shell.Run "mmsys.cpl", 1
            Case "11": shell.Run "taskschd.msc", 1
            Case "12": shell.Run "powershell -Command ""Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 Name, WorkingSet64""", 1
            Case "13": shell.Run "cmd /k powercfg /batteryreport /output C:\Windows\Temp\battery.html", 0
            Case "14": RunLEDToggle
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub ShowFunMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: ZABAWA" & vbCrLf & SEP & vbCrLf & _
               "1. Tryb Disco Ekranu" & vbCrLf & _
               "2. Gadajacy Asystent" & vbCrLf & _
               "3. Symulator Deszczu" & vbCrLf & _
               "4. Losowe Ciekawostki" & vbCrLf & _
               "5. Efekt Matrix" & vbCrLf & _
               "6. Symulator Pisania na Maszynie" & vbCrLf & _
               "7. Generator Bledow" & vbCrLf & _
               "8. Wymus BSOD" & vbCrLf & _
               "9. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Zabawa", "1")
        If wyb = "" Or wyb = "9" Then Exit Do
        Select Case wyb
            Case "1" : RunDiscoMode
            Case "2" : RunTalkingAssistant
            Case "3" : RunRainSim
            Case "4" : RunRandomFacts
            Case "5" : RunMatrix
            Case "6" : RunTypewriter
            Case "7" : RunErrorLab
            Case "8" : RunBSODFlow
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

' ===== FUNKCJE SYSTEMOWE =====

Sub RunGamerMode
    If MsgBox("Zatrzymac procesy tla (OneDrive, Teams, Spotify, Edge)?", vbYesNo + vbQuestion, "Tryb Gracza") = vbNo Then Exit Sub
    Dim procs : procs = Array("OneDrive.exe", "Teams.exe", "Skype.exe", "Spotify.exe", "Cortana.exe", "SearchUI.exe", "ShellExperienceHost.exe", "YourPhone.exe", "RuntimeBroker.exe", "GameBar.exe", "msedge.exe", "chrome.exe")
    Dim p
    For Each p In procs : shell.Run "taskkill /f /t /im """ & p & """ 2>nul", 0, True : Next
    MsgBox "Tryb gracza aktywny.", vbInformation, "Sukces"
End Sub

Sub RunCleaner
    If MsgBox("Czyszczenie usunie:" & vbCrLf & "- %TEMP%" & vbCrLf & "- %WINDIR%\Temp" & vbCrLf & "- %WINDIR%\Prefetch" & vbCrLf & "- Cache DNS" & vbCrLf & "Kontynuowac?", vbQuestion + vbYesNo, "Czyszczenie") = vbYes Then
        shell.Run "cmd /c ""del /f /s /q %TEMP%\*.* 2>nul & del /f /s /q %WINDIR%\Temp\*.* 2>nul & del /f /s /q %WINDIR%\Prefetch\*.* 2>nul & ipconfig /flushdns 2>nul""", 0, True
        MsgBox "Czyszczenie zakonczone!", vbInformation, "Sukces"
    End If
End Sub

Sub RunResourceMonitor
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set cpu = wmi.ExecQuery("Select * from Win32_Processor") : For Each item In cpu : cpuLoad = item.LoadPercentage : Next
    Set mem = wmi.ExecQuery("Select * from Win32_OperatingSystem") : For Each item In mem : ramFree = FormatNumber((item.TotalVisibleMemorySize - item.FreePhysicalMemory) / 1048576, 1) & " GB / " & FormatNumber(item.TotalVisibleMemorySize / 1048576, 1) & " GB" : Next
    Set disks = wmi.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3") : diskSpace = ""
    For Each d In disks : diskSpace = diskSpace & " [" & d.DeviceID & "] " & FormatNumber(d.FreeSpace / 1073741824, 1) & " GB wolne / " & FormatNumber(d.Size / 1073741824, 1) & " GB" & vbCrLf : Next
    On Error GoTo 0
    MsgBox "ZASOBY SYSTEMU:" & vbCrLf & "CPU: " & cpuLoad & "%" & vbCrLf & "RAM: " & ramFree & vbCrLf & "Dysk:" & diskSpace, vbInformation, "Monitor"
End Sub

Sub RunHardwareInfo
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set cpu = wmi.ExecQuery("Select * from Win32_Processor") : For Each item In cpu : cpuName = item.Name : Next
    Set gpu = wmi.ExecQuery("Select * from Win32_VideoController") : For Each item In gpu : gpuName = item.Name : Next
    Set os = wmi.ExecQuery("Select * from Win32_OperatingSystem") : For Each item In os : osName = item.Caption : ramGB = FormatNumber(item.TotalVisibleMemorySize / 1048576, 1) : Next
    On Error GoTo 0
    MsgBox "SPRZET I SYSTEM:" & vbCrLf & "CPU: " & cpuName & vbCrLf & "GPU: " & gpuName & vbCrLf & "OS: " & osName & vbCrLf & "RAM: " & ramGB & " GB", vbInformation, "Info"
End Sub

Sub RunNetwork
    netMenu = "NARZEDZIA SIECIOWE:" & vbCrLf & "1. IP Lokalny" & vbCrLf & "2. IP Publiczny" & vbCrLf & "3. Wylacz Adapter" & vbCrLf & "4. Wlacz Adapter" & vbCrLf & "5. Reset Stosu Sieci" & vbCrLf & "6. Test Ping"
    netWybor = InputBox(netMenu, "Siec", "1")
    Select Case netWybor
        Case "1" : Set exec = shell.Exec("cmd /c ipconfig | findstr /C:""IPv4""") : MsgBox "Lokalne IP:" & vbCrLf & exec.StdOut.ReadAll, vbInformation, "IP"
        Case "2" : MsgBox "Sprawdzanie...", vbInformation, "IP" : Set exec = shell.Exec("powershell -Command ""(Invoke-RestMethod http://ifconfig.me/ip).Trim()""") : MsgBox "Publiczne IP:" & vbCrLf & exec.StdOut.ReadAll, vbInformation, "IP"
        Case "3" : adapter = InputBox("Nazwa adaptera (np. Wi-Fi):", "Wylacz", "Wi-Fi") : If adapter <> "" Then shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command Disable-NetAdapter -Name ''' & adapter & ''' -Confirm:$false'""", 0
        Case "4" : adapter = InputBox("Nazwa adaptera (np. Wi-Fi):", "Wlacz", "Wi-Fi") : If adapter <> "" Then shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command Enable-NetAdapter -Name ''' & adapter & ''' -Confirm:$false'""", 0
        Case "5" : MsgBox "Resetowanie stosu sieciowego...", vbExclamation, "Reset" : shell.Run "cmd /c netsh winsock reset && netsh int ip reset && ipconfig /release && ipconfig /renew", 0, True : MsgBox "Reset zakoczony.", vbInformation, "Gotowe"
        Case "6" : host = InputBox("Adres docelowy:", "Ping", "8.8.8.8") : If host <> "" Then Set exec = shell.Exec("ping -n 4 " & host) : MsgBox exec.StdOut.ReadAll, vbInformation, "Wynik Ping"
    End Select
End Sub

Sub RunWiFiPass
    MsgBox "Pobieranie profilow WiFi...", vbInformation, "WiFi"
    Set exec = shell.Exec("cmd /c netsh wlan show profiles")
    MsgBox "PROFILE WIFI:" & vbCrLf & exec.StdOut.ReadAll, vbInformation, "Lista"
    key = InputBox("Podaj nazwe sieci, aby pokazac haslo:", "Haslo WiFi")
    If key <> "" Then Set exec2 = shell.Exec("cmd /c netsh wlan show profile name=""" & key & """ key=clear") : MsgBox exec2.StdOut.ReadAll, vbInformation, "Haslo: " & key
End Sub

Sub RunDefender
    dMenu = "MENADZER DEFENDERA:" & vbCrLf & "1. Wylacz Ochrone w Czasie Rzeczywistym" & vbCrLf & "2. Wlacz Ochrone w Czasie Rzeczywistym"
    dWybor = InputBox(dMenu, "Defender", "1")
    Select Case dWybor
        Case "1" : MsgBox "Wylaczanie Defendera... Wymaga UAC.", vbInformation, "Defender" : shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command Set-MpPreference -DisableRealtimeMonitoring $true'""", 0 : MsgBox "Defender wylaczony.", vbInformation, "Info"
        Case "2" : MsgBox "Wlaczanie Defendera... Wymaga UAC.", vbInformation, "Defender" : shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command Set-MpPreference -DisableRealtimeMonitoring $false'""", 0 : MsgBox "Defender wlaczony.", vbInformation, "Info"
        Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
    End Select
End Sub

Sub RunAdvanced
    advMenu = "NARZEDZIA ZAAWANSOWANE:" & vbCrLf & "1. Edytor Rejestru" & vbCrLf & "2. Utworz GodMode" & vbCrLf & "3. Ustawienia UAC"
    advWybor = InputBox(advMenu, "Zaawansowane", "1")
    Select Case advWybor
        Case "1" : shell.Run "regedit.exe", 1
        Case "2" : d = shell.SpecialFolders("Desktop") : shell.Run "cmd /c mkdir """ & d & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}""", 0
        Case "3" : shell.Run "UserAccountControlSettings.exe", 1
    End Select
End Sub

Sub RunRepair
    repMenu = "ZESTAW NAPRAWCZY:" & vbCrLf & "1. SFC Scan" & vbCrLf & "2. DISM Restore" & vbCrLf & "3. Czyszczenie Dysku"
    repWybor = InputBox(repMenu, "Naprawa", "1")
    Select Case repWybor
        Case "1" : MsgBox "Uruchamianie SFC...", vbInformation, "SFC" : shell.Run "cmd /k sfc /scannow", 1
        Case "2" : MsgBox "Uruchamianie DISM...", vbInformation, "DISM" : shell.Run "cmd /k DISM /Online /Cleanup-Image /RestoreHealth", 1
        Case "3" : shell.Run "cleanmgr", 1
    End Select
End Sub

Sub RunQuickTools
    qMenu = "SZYBKIE NARZEDZIA:" & vbCrLf & "1. CMD" & vbCrLf & "2. PowerShell" & vbCrLf & "3. Menedzer Urzadzen" & vbCrLf & "4. Uslugi" & vbCrLf & "5. Panel Sterowania"
    qWybor = InputBox(qMenu, "Szybkie", "1")
    Select Case qWybor
        Case "1" : shell.Run "cmd.exe", 1
        Case "2" : shell.Run "powershell.exe", 1
        Case "3" : shell.Run "devmgmt.msc", 1
        Case "4" : shell.Run "services.msc", 1
        Case "5" : shell.Run "control.exe", 1
    End Select
End Sub

Sub RunExtras
    eMenu = "OPCJE DODATKOWE:" & vbCrLf & "1. Punkt Przywrocenia" & vbCrLf & "2. Ukryj/Pokaz Pliki" & vbCrLf & "3. Wyczysc Schowek" & vbCrLf & "4. Wylacz/Wlacz Aktualizacje" & vbCrLf & "5. Tryb Ciemny/Jasny" & vbCrLf & "6. Wyczysc Kosz" & vbCrLf & "7. Wylacz Menedzer Zadan" & vbCrLf & "8. Eksport Info do TXT"
    eWybor = InputBox(eMenu, "Dodatki", "1")
    Select Case eWybor
        Case "1" : shell.Run "cmd /c wmic /namespace:\\root\default path SystemRestore Call CreateRestorePoint ""VM Commander"", 100, 7", 0, True : MsgBox "Punkt przywrocenia utworzony.", vbInformation, "Sukces"
        Case "2" : shell.Run "reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"" /v Hidden /t REG_DWORD /d 1 /f", 0 : shell.Run "taskkill /f /im explorer.exe", 0 : shell.Run "explorer.exe", 1 : MsgBox "Widocznosc zmieniona.", vbInformation, "Sukces"
        Case "3" : shell.Run "cmd /c echo off | clip", 0, True : MsgBox "Schowek wyczyszczony.", vbInformation, "Sukces"
        Case "4"
            If MsgBox("Wylaczyc aktualizacje?", vbYesNo, "Aktualizacje") = vbYes Then shell.Run "net stop wuauserv", 0 : shell.Run "sc config wuauserv start= disabled", 0 : MsgBox "Wylaczone.", vbInformation, "Zmiana" Else shell.Run "sc config wuauserv start= auto", 0 : shell.Run "net start wuauserv", 0 : MsgBox "Wlaczone.", vbInformation, "Zmiana"
        Case "5"
            Set reg = CreateObject("WScript.Shell") : On Error Resume Next : val = reg.RegRead("HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme") : On Error GoTo 0
            If val = 0 Then newVal = 1 Else newVal = 0
            reg.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", newVal, "REG_DWORD"
            reg.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme", newVal, "REG_DWORD"
            shell.Run "taskkill /f /im explorer.exe", 0 : shell.Run "explorer.exe", 1 : MsgBox "Motyw zmieniony.", vbInformation, "Sukces"
        Case "6" : shell.Run "powershell -Command ""Clear-RecycleBin -Confirm:$false""", 0, True : MsgBox "Kosz wyczyszczony.", vbInformation, "Sukces"
        Case "7" : val = MsgBox("Wylaczyc Menedzer Zadan?", vbYesNo, "Menedzer") : If val = vbYes Then shell.Run "reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System"" /v DisableTaskMgr /t REG_DWORD /d 1 /f", 0 Else shell.Run "reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System"" /v DisableTaskMgr /t REG_DWORD /d 0 /f", 0 : MsgBox "Zaktualizowano.", vbInformation, "Gotowe"
        Case "8"
            On Error Resume Next : Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
            Set os = wmi.ExecQuery("Select * from Win32_OperatingSystem") : For Each item In os : osName = item.Caption : Next
            Set comp = wmi.ExecQuery("Select * from Win32_ComputerSystem") : For Each item In comp : compName = item.Name : ramGB = FormatNumber(item.TotalVisibleMemorySize / 1048576, 1) : Next
            On Error GoTo 0 : txtPath = fso.GetSpecialFolder(2) & "\sys_info.txt" : Set ts = fso.CreateTextFile(txtPath, True)
            ts.WriteLine "NAZWA: " & compName & vbCrLf & "SYSTEM: " & osName & vbCrLf & "RAM: " & ramGB & " GB" : ts.Close
            shell.Run "notepad.exe """ & txtPath & """", 1 : MsgBox "Wyeksportowano.", vbInformation, "Sukces"
    End Select
End Sub

Sub RunPasswordGen
    lenIn = InputBox("Dlugosc hasla (8-32):", "Generator", "16")
    If lenIn = "" Or Not IsNumeric(lenIn) Then lenIn = "16"
    Dim chars : chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-="
    Dim pass : pass = "" : Dim i
    For i = 1 To CInt(lenIn) : pass = pass & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1) : Next
    MsgBox "Haslo:" & vbCrLf & vbCrLf & pass, vbInformation, "Wygenerowano"
End Sub

Sub RunStartupManager
    menu = "MENADZER AUTOSTARTU:" & vbCrLf & "1. Otworz folder Autostartu" & vbCrLf & "2. Pokaz programy startowe (Registry)" & vbCrLf & "3. Wylacz wybrany program z autostartu"
    sWybor = InputBox(menu, "Autostart", "1")
    Select Case sWybor
        Case "1" : shell.Run "shell:startup", 1
        Case "2"
            Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
            Set items = wmi.ExecQuery("Select * from Win32_StartupCommand")
            list = ""
            For Each item In items : list = list & item.Name & " -> " & item.Command & vbCrLf : Next
            MsgBox "PROGRAMY AUTOSTARTU:" & vbCrLf & list, vbInformation, "Lista"
        Case "3"
            app = InputBox("Wpisz dokladna nazwe programu z listy wyzej:", "Wylacz", "Program.exe")
            If app <> "" Then shell.Run "reg delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v """ & app & """ /f", 0, True : MsgBox "Usunieto z autostartu.", vbInformation, "Gotowe"
    End Select
End Sub

Sub RunLEDToggle
    ledMenu = "STEROVANIE DIODAMI KLAWIATURY:" & vbCrLf & "1. CapsLock" & vbCrLf & "2. NumLock" & vbCrLf & "3. ScrollLock"
    ledWybor = InputBox(ledMenu, "Dioda", "1")
    Select Case ledWybor
        Case "1" : shell.SendKeys "{CAPSLOCK}"
        Case "2" : shell.SendKeys "{NUMLOCK}"
        Case "3" : shell.SendKeys "{SCROLLLOCK}"
        Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
    End Select
    MsgBox "Dioda zmieniona.", vbInformation, "Sukces"
End Sub

' ===== FUNKCJE ROZRYWKOWE =====

Sub RunDiscoMode
    Dim htaPath : htaPath = fso.GetSpecialFolder(2) & "\disco.hta"
    Dim ts : Set ts = fso.CreateTextFile(htaPath, True)
    ts.WriteLine "<html><head><title>Disco Mode</title></head><body style=""margin:0"">"
    ts.WriteLine "<script>"
    ts.WriteLine "var c=['#f00','#0f0','#00f','#ff0','#0ff','#f0f'],i=0;"
    ts.WriteLine "setInterval(function(){document.body.style.backgroundColor=c[i++%c.length];},200);"
    ts.WriteLine "</script><h1 style=""color:white;text-align:center;margin-top:30%"">DISCO MODE</h1></body></html>"
    ts.Close
    shell.Run "mshta.exe """ & htaPath & """", 1
    MsgBox "Zamknij okno HTA, aby wylaczyc disco.", vbInformation, "Disco"
    fso.DeleteFile htaPath, True
End Sub

Sub RunTalkingAssistant
    If MsgBox("Czy chcesz, zeby komputer cos powiedzial?", vbYesNo, "Asystent") = vbNo Then Exit Sub
    txt = InputBox("Co ma powiedziec?", "Glos", "VM Commander dziala poprawnie.")
    If txt <> "" Then
        Set sapi = CreateObject("SAPI.SpVoice")
        sapi.Speak txt
        MsgBox "Powiedziano: " & txt, vbInformation, "Asystent"
    End If
End Sub

Sub RunRainSim
    MsgBox "Symulator deszczu aktywny. Kliknij OK, aby uruchomic.", vbInformation, "Deszcz"
    Dim bat : bat = fso.GetSpecialFolder(2) & "\rain.bat"
    Dim ts : Set ts = fso.CreateTextFile(bat, True)
    ts.WriteLine "@echo off" & vbCrLf & "title Rain Simulator" & vbCrLf & ":start" & vbCrLf & "echo . . . . . . . ." & vbCrLf & "echo  . . . . . . ." & vbCrLf & "echo . . . . . . . ." & vbCrLf & "echo  . . . . . . ." & vbCrLf & "cls" & vbCrLf & "goto start"
    ts.Close
    shell.Run "cmd /k """ & bat & """", 1
    MsgBox "Zamknij okno CMD, aby zakonczyc deszcz.", vbInformation, "Deszcz"
    fso.DeleteFile bat, True
End Sub

Sub RunRandomFacts
    facts = Array("Kurczaki nie spia z zamknietymi oczami.", "Kawior pochodzi od ryb jesiotrowatych.", "Ludzkie oko potrafi rozroznic 10 mln kolorow.", "Mrowki nie spia.", "Woda w studni moze byc starsza od Ziemi.", "Kreciki sa slepe.", "Pierwsza gra komputerowa miala 1kb pamieci.", "Wenecja tonie o 1-2 mm rocznie.")
    MsgBox "NIEPRZYDATNA CIEKAWOSTKA:" & vbCrLf & vbCrLf & facts(Int(Rnd() * UBound(facts) + 1)), vbInformation, "Losowy Fakt"
End Sub

Sub RunTypewriter
    MsgBox "Symulator pisania na maszynie. Kliknij OK, aby rozpoczac.", vbInformation, "Symulator"
    WScript.Sleep 500
    shell.SendKeys "VM COMMANDER v7.0 INITIALIZED"
    WScript.Sleep 200
    shell.SendKeys "{ENTER}"
    shell.SendKeys "TYPING SEQUENCE ENGAGED..."
    WScript.Sleep 200
    shell.SendKeys "{ENTER}"
    shell.SendKeys "{CAPSLOCK}"
    MsgBox "Symulator aktywny. Zamykanie...", vbInformation, "Koniec"
End Sub

Sub RunErrorLab
    errMenu = "GENERATOR BLEDOW:" & vbCrLf & "1. Krytyczny" & vbCrLf & "2. Ostrzezenie" & vbCrLf & "3. Informacja" & vbCrLf & "4. Pytanie"
    e = InputBox(errMenu, "Bledy", "1")
    If e <> "" Then
        t = InputBox("Tresc:", "Tekst", "Wystapil blad!")
        ti = InputBox("Tytul:", "Naglowek", "Blad Systemu")
        Select Case e
            Case "1" : MsgBox t, 16, ti : Case "2" : MsgBox t, 48, ti : Case "3" : MsgBox t, 64, ti : Case "4" : MsgBox t, 32, ti
        End Select
    End If
End Sub

Sub RunMatrix
    Dim bat : bat = fso.GetSpecialFolder(2) & "\vm_matrix.bat"
    Dim ts : Set ts = fso.CreateTextFile(bat, True)
    ts.WriteLine "@echo off" & vbCrLf & "color 0a" & vbCrLf & ":loop" & vbCrLf & "echo %random%%random% %random%%random% %random%%random% %random%%random%" & vbCrLf & "goto loop"
    ts.Close
    shell.Run "cmd /k """ & bat & """", 1
    MsgBox "Zamknij okno CMD, aby zakonczyc.", vbInformation, "Matrix"
    fso.DeleteFile bat, True
End Sub

' ===== BSOD - DOKLADNY PRZEPYW =====

Sub RunBSODFlow
    If MsgBox("CZY NA PEWNO?", vbYesNo + vbCritical, "OSTRZEZENIE") = vbNo Then Exit Sub
    If MsgBox("ZAPISALES ZMIANY?", vbYesNo + vbExclamation, "OSTRZEZENIE") = vbNo Then Exit Sub
    
    MsgBox "Zabijanie aktywnych procesow...", vbInformation, "BSOD"
    shell.Run "taskkill /f /t /fi ""STATUS eq RUNNING"" /fi ""USERNAME ne NT AUTHORITY\SYSTEM""", 0, True
    WScript.Sleep 1500
    
    MsgBox "Powodowanie BSOD...", vbCritical, "BSOD"
    WScript.Sleep 800
    
    Dim psPath : psPath = fso.GetSpecialFolder(2) & "\vm_bsod.ps1"
    Dim ts : Set ts = fso.CreateTextFile(psPath, True)
    ts.WriteLine "$code = @'"
    ts.WriteLine "using System;"
    ts.WriteLine "using System.Runtime.InteropServices;"
    ts.WriteLine "public class BSOD {"
    ts.WriteLine "  [DllImport(" & Chr(34) & "ntdll.dll" & Chr(34) & ")] public static extern uint RtlAdjustPrivilege(int p, bool b, bool t, ref bool v);"
    ts.WriteLine "  [DllImport(" & Chr(34) & "ntdll.dll" & Chr(34) & ")] public static extern uint NtRaiseHardError(uint e, uint n, uint u, IntPtr p, uint v, ref uint r);"
    ts.WriteLine "}"
    ts.WriteLine "'@"
    ts.WriteLine "Add-Type -TypeDefinition $code"
    ts.WriteLine "$priv = $false"
    ts.WriteLine "[BSOD]::RtlAdjustPrivilege(19, $true, $false, [ref]$priv)"
    ts.WriteLine "$resp = 0"
    ts.WriteLine "[BSOD]::NtRaiseHardError(0xC0000022, 0, 0, 0, 6, [ref]$resp)"
    ts.Close
    
    On Error Resume Next
    shell.Run "powershell -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """", 0, False
    If Err.Number <> 0 Then MsgBox "Nie udalo sie wywolac BSOD. Uruchom jako Administrator.", vbCritical, "Blad"
    On Error GoTo 0
    WScript.Sleep 2000
    fso.DeleteFile psPath, True
End Sub
