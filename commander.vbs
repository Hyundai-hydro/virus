Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Dim TITLE : TITLE = "VM COMMANDER v9.0 [STABLE]"
Dim SEP : SEP = String(48, "=")
Randomize

' ===== GLÓWNA PĘTLA =====
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

' ===== ZAKŁADKI =====

Sub ShowSystemMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: SYSTEM" & vbCrLf & SEP & vbCrLf & _
               "1. Menedzer Zadan" & vbCrLf & _
               "2. Tryb Gracza" & vbCrLf & _
               "3. Oczyszczanie Systemu" & vbCrLf & _
               "4. Monitor Zasobow (CPU/RAM/Dysk)" & vbCrLf & _
               "5. Lista Procesow" & vbCrLf & _
               "6. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "System", "1")
        If wyb = "" Or wyb = "6" Then Exit Do
        Select Case wyb
            Case "1" : shell.Run "taskmgr", 1
            Case "2" : RunGamerMode
            Case "3" : RunCleaner
            Case "4" : RunResourceMonitor
            Case "5" : RunProcessList
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub ShowNetworkMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: SIEC" & vbCrLf & SEP & vbCrLf & _
               "1. IP Lokalny" & vbCrLf & _
               "2. IP Publiczny" & vbCrLf & _
               "3. Test Ping" & vbCrLf & _
               "4. Reset Stosu Sieci" & vbCrLf & _
               "5. Wylacz/Wlacz Adapter" & vbCrLf & _
               "6. Otworz Polaczenia Sieciowe" & vbCrLf & _
               "7. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Siec", "1")
        If wyb = "" Or wyb = "7" Then Exit Do
        Select Case wyb
            Case "1" : RunGetLocalIP
            Case "2" : RunGetPublicIP
            Case "3" : RunPingTest
            Case "4" : RunNetworkReset
            Case "5" : RunAdapterToggle
            Case "6" : shell.Run "ncpa.cpl", 1
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
               "5. Uspij / Hibernuj" & vbCrLf & _
               "6. Slide to Shut Down" & vbCrLf & _
               "7. Anuluj Akcje" & vbCrLf & _
               "8. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Zasilanie", "4")
        If wyb = "" Or wyb = "8" Then Exit Do
        Select Case wyb
            Case "1" : shell.Run "powershell -Command ""Start-Process shutdown -ArgumentList '/r /fw /t 0' -Verb RunAs""", 0
            Case "2" : RunLogoff
            Case "3" : RunRestart
            Case "4" : RunShutdown
            Case "5" : RunSleepHibernate
            Case "6" : shell.Run "SlideToShutDown.exe", 1
            Case "7" : shell.Run "shutdown /a", 0 : MsgBox "Akcja anulowana!", vbInformation, "Anulowano"
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub ShowToolsMenu
    Do
        menu = SEP & vbCrLf & "ZAKLADKA: NARZEDZIA" & vbCrLf & SEP & vbCrLf & _
               "1. Edytor Rejestru / GodMode / UAC" & vbCrLf & _
               "2. Narzedzia Naprawcze" & vbCrLf & _
               "3. Szybkie Narzedzia (CMD/PS/Uslugi)" & vbCrLf & _
               "4. Opcje Dodatkowe (Motyw/Schowek/Kosz)" & vbCrLf & _
               "5. Generator Bezpiecznych Hasel" & vbCrLf & _
               "6. Menedzer Autostartu" & vbCrLf & _
               "7. Zarzadzanie Dyskami" & vbCrLf & _
               "8. Lista Programow i Funkcji" & vbCrLf & _
               "9. Logi Systemowe" & vbCrLf & _
               "10. Planer Zadan Systemowych" & vbCrLf & _
               "11. Raport Stanu Baterii" & vbCrLf & _
               "12. Sterowanie Diodami Klawiatury" & vbCrLf & _
               "13. Status Aktywacji Windows" & vbCrLf & _
               "14. Wroc do glownego menu" & vbCrLf & SEP
        wyb = InputBox(menu, "Narzedzia", "1")
        If wyb = "" Or wyb = "14" Then Exit Do
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
            Case "10": shell.Run "taskschd.msc", 1
            Case "11": shell.Run "cmd /c powercfg /batteryreport /output ""%TEMP%\battery.html"" && start ""%TEMP%\battery.html""", 0
            Case "12": RunLEDToggle
            Case "13": shell.Run "slmgr /xpr", 1
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
    For Each p In procs
        On Error Resume Next
        shell.Run "taskkill /f /t /im """ & p & """ 2>nul", 0, True
        On Error GoTo 0
    Next
    MsgBox "Tryb gracza aktywny.", vbInformation, "Sukces"
End Sub

Sub RunCleaner
    If MsgBox("Czyszczenie usunie:" & vbCrLf & "- %TEMP%" & vbCrLf & "- %WINDIR%\Temp" & vbCrLf & "- %WINDIR%\Prefetch" & vbCrLf & "- Cache DNS" & vbCrLf & "Kontynuowac?", vbQuestion + vbYesNo, "Czyszczenie") = vbYes Then
        On Error Resume Next
        shell.Run "cmd /c ""del /f /s /q %TEMP%\*.* 2>nul & del /f /s /q %WINDIR%\Temp\*.* 2>nul & del /f /s /q %WINDIR%\Prefetch\*.* 2>nul & ipconfig /flushdns 2>nul""", 0, True
        On Error GoTo 0
        MsgBox "Czyszczenie zakonczone!", vbInformation, "Sukces"
    End If
End Sub

Sub RunResourceMonitor
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set cpu = wmi.ExecQuery("Select * from Win32_Processor") : For Each item In cpu : cpuLoad = item.LoadPercentage : Next
    Set mem = wmi.ExecQuery("Select * from Win32_OperatingSystem") : For Each item In mem : ramUsed = FormatNumber((item.TotalVisibleMemorySize - item.FreePhysicalMemory) / 1048576, 1) : ramTot = FormatNumber(item.TotalVisibleMemorySize / 1048576, 1) : Next
    Set disks = wmi.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3") : diskInfo = ""
    For Each d In disks : diskInfo = diskInfo & " [" & d.DeviceID & "] " & FormatNumber(d.FreeSpace / 1073741824, 1) & " GB wolne / " & FormatNumber(d.Size / 1073741824, 1) & " GB" & vbCrLf : Next
    On Error GoTo 0
    MsgBox "ZASOBY SYSTEMU:" & vbCrLf & "CPU: " & cpuLoad & "%" & vbCrLf & "RAM: " & ramUsed & " / " & ramTot & " GB" & vbCrLf & "Dysk:" & diskInfo, vbInformation, "Monitor"
End Sub

Sub RunProcessList
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set procs = wmi.ExecQuery("Select Name, WorkingSetSize, ProcessId from Win32_Process ORDER BY WorkingSetSize DESC")
    list = ""
    count = 0
    For Each p In procs
        If p.Name <> "" And p.Name <> "Idle" And p.Name <> "System" Then
            mb = Round(p.WorkingSetSize / 1048576, 1)
            list = list & "[" & p.ProcessId & "] " & p.Name & " | " & mb & " MB" & vbCrLf
            count = count + 1
            If count >= 15 Then Exit For
        End If
    Next
    On Error GoTo 0
    If list = "" Then list = "Brak procesow do wyswietlenia."
    MsgBox "TOP 15 PROCESOW (WEDLUG RAM):" & vbCrLf & list, vbInformation, "Lista Procesow"
End Sub

Sub RunLogoff
    t = InputBox("Opoznienie wylogowania (s):", "Czas", "5")
    If t = "" Or Not IsNumeric(t) Then t = "5"
    shell.Run "cmd /c timeout /t " & CInt(t) & " /nobreak >nul && shutdown /l", 0
End Sub

Sub RunRestart
    t = InputBox("Opoznienie restartu (s):", "Czas", "30")
    If t = "" Or Not IsNumeric(t) Then t = "30"
    m = InputBox("Komunikat na ekranie:", "Tekst", "Restartuje...")
    If m = "" Then m = "Restartuje..."
    shell.Run "shutdown /r /t " & CInt(t) & " /c """ & m & """", 0
End Sub

Sub RunShutdown
    t = InputBox("Opoznienie zamkniecia (s):", "Czas", "30")
    If t = "" Or Not IsNumeric(t) Then t = "30"
    m = InputBox("Komunikat na ekranie:", "Tekst", "Zamykam...")
    If m = "" Then m = "Zamykam..."
    shell.Run "shutdown /s /t " & CInt(t) & " /c """ & m & """", 0
End Sub

Sub RunSleepHibernate
    choice = MsgBox("1 = Uspij (Sleep)" & vbCrLf & "2 = Hibernacja", vbOKCancel, "Zasilanie")
    If choice = vbOK Then shell.Run "rundll32.exe powrprof.dll,SetSuspendState 0,1,0", 0
    If choice = vbCancel Then shell.Run "rundll32.exe powrprof.dll,SetSuspendState Hibernate", 0
End Sub

' ===== FUNKCJE NARZEDZIOWE =====

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
    Do
        menu = SEP & vbCrLf & "NARZEDZIA NAPRAWCZE" & vbCrLf & SEP & vbCrLf & _
               "1. SFC Scan" & vbCrLf & _
               "2. DISM Restore" & vbCrLf & _
               "3. Czyszczenie Dysku" & vbCrLf & _
               "4. Napraw procesu" & vbCrLf & _
               "5. Optymalizacja Dyskow" & vbCrLf & _
               "6. Wroc do narzedzi" & vbCrLf & SEP
        wyb = InputBox(menu, "Naprawa", "1")
        If wyb = "" Or wyb = "6" Then Exit Do
        Select Case wyb
            Case "1" : MsgBox "Uruchamianie SFC...", vbInformation, "SFC" : shell.Run "cmd /k sfc /scannow", 1
            Case "2" : MsgBox "Uruchamianie DISM...", vbInformation, "DISM" : shell.Run "cmd /k DISM /Online /Cleanup-Image /RestoreHealth", 1
            Case "3" : shell.Run "cleanmgr", 1
            Case "4" : RunFixProcesses
            Case "5" : shell.Run "dfrgui", 1
            Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
        End Select
    Loop
End Sub

Sub RunFixProcesses
    If MsgBox("Zabije wszystkie procesy uzytkownika i uruchomi explorer ponownie. Kontynuowac?", vbYesNo + vbExclamation, "Napraw Procesow") = vbNo Then Exit Sub
    On Error Resume Next
    shell.Run "taskkill /f /fi ""STATUS eq RUNNING"" /fi ""USERNAME ne NT AUTHORITY\SYSTEM""", 0, True
    shell.Run "explorer.exe", 1
    On Error GoTo 0
    MsgBox "Procesy zostaly zabite. Explorer uruchomiony ponownie.", vbInformation, "Sukces"
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
    menu = "MENADZER AUTOSTARTU:" & vbCrLf & "1. Otworz folder Autostartu" & vbCrLf & "2. Pokaz programy startowe" & vbCrLf & "3. Wylacz wybrany program"
    sWybor = InputBox(menu, "Autostart", "1")
    Select Case sWybor
        Case "1" : shell.Run "shell:startup", 1
        Case "2"
            On Error Resume Next
            Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
            Set items = wmi.ExecQuery("Select * from Win32_StartupCommand")
            list = ""
            For Each item In items : list = list & item.Name & " -> " & item.Command & vbCrLf : Next
            If list = "" Then list = "Brak wpisow w rejestrze."
            On Error GoTo 0
            MsgBox "PROGRAMY AUTOSTARTU:" & vbCrLf & list, vbInformation, "Lista"
        Case "3"
            app = InputBox("Wpisz nazwe programu:", "Wylacz", "Program.exe")
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

' ===== FUNKCJE SIECIOWE (STABILNE) =====

Sub RunGetLocalIP
    On Error Resume Next
    Set exec = shell.Exec("cmd /c ipconfig | findstr /C:""IPv4""")
    out = exec.StdOut.ReadAll
    On Error GoTo 0
    If Trim(out) = "" Then out = "Nie znaleziono IPv4."
    MsgBox "IP Lokalne:" & vbCrLf & out, vbInformation, "Siec"
End Sub

Sub RunGetPublicIP
    MsgBox "Sprawdzanie IP zewnetrznego...", vbInformation, "Czekaj"
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://api.ipify.org", False
    http.Send
    If http.Status = 200 Then
        MsgBox "IP Publiczne: " & http.responseText, vbInformation, "Siec"
    Else
        Set exec = shell.Exec("powershell -Command ""(Invoke-RestMethod ifconfig.me/ip).Trim()""")
        MsgBox "IP Publiczne: " & exec.StdOut.ReadAll, vbInformation, "Siec"
    End If
    On Error GoTo 0
End Sub

Sub RunPingTest
    host = InputBox("Adres docelowy:", "Ping", "8.8.8.8")
    If host = "" Then Exit Sub
    MsgBox "Testowanie polaczenia... (moze potrwac 5s)", vbInformation, "Ping"
    On Error Resume Next
    Set exec = shell.Exec("ping -n 4 -w 1000 " & host)
    result = exec.StdOut.ReadAll
    On Error GoTo 0
    MsgBox "WYNIK PING:" & vbCrLf & result, vbInformation, "Siec"
End Sub

Sub RunNetworkReset
    MsgBox "Resetowanie stosu sieciowego...", vbExclamation, "Reset"
    On Error Resume Next
    shell.Run "cmd /c netsh winsock reset && netsh int ip reset && ipconfig /flushdns && ipconfig /release && ipconfig /renew", 0, True
    On Error GoTo 0
    MsgBox "Reset zakoczony. Zalecany restart systemu.", vbInformation, "Gotowe"
End Sub

Sub RunAdapterToggle
    act = MsgBox("1 = Wylacz adapter" & vbCrLf & "2 = Wlacz adapter", vbOKCancel, "Adapter")
    If act = vbCancel Then Exit Sub
    adapter = InputBox("Nazwa adaptera (np. Wi-Fi, Ethernet):", "Nazwa", "Wi-Fi")
    If adapter = "" Then Exit Sub
    cmd = "Disable"
    If act = 2 Then cmd = "Enable"
    shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command " & cmd & "-NetAdapter -Name ''' & adapter & ''' -Confirm:$false'""", 0
    MsgBox "Wyslano komende '" & cmd & "' dla adaptera '" & adapter & "'.", vbInformation, "Adapter"
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
    shell.SendKeys "VM COMMANDER v9.0 INITIALIZED"
    WScript.Sleep 200
    shell.SendKeys "{ENTER}"
    shell.SendKeys "TYPING SEQUENCE ENGAGED..."
    WScript.Sleep 200
    shell.SendKeys "{ENTER}"
    shell.SendKeys "{CAPSLOCK}"
    MsgBox "Symulator aktywny. Zamykanie...", vbInformation, "Koniec"
End Sub

Sub RunErrorLab
    errMenu = "GENERATOR BLEDOW:" & vbCrLf & _
              "1. Blad Krytyczny" & vbCrLf & _
              "2. Ostrzezenie Systemowe" & vbCrLf & _
              "3. Informacja" & vbCrLf & _
              "4. Pytanie (Tak/Nie)" & vbCrLf & _
              "5. Blad Dostepu (Access Denied)" & vbCrLf & _
              "6. Brak Pliku (Not Found)" & vbCrLf & _
              "7. Blad Sprzetowy" & vbCrLf & _
              "8. Blad Polaczenia" & vbCrLf & _
              "9. Pomoc (Help)" & vbCrLf & _
              "10. Losowy Blad"
    e = InputBox(errMenu, "Bledy", "1")
    If e = "" Then Exit Sub
    
    txt = InputBox("Tresc:", "Tekst", "Wystapil blad!")
    ti = InputBox("Tytul:", "Naglowek", "System Error")
    
    Select Case CInt(e)
        Case 1 : MsgBox txt, 16, ti
        Case 2 : MsgBox txt, 48, ti
        Case 3 : MsgBox txt, 64, ti
        Case 4 : MsgBox txt, 32, ti
        Case 5 : MsgBox txt, 16, "ODMOWA DOSTEPU"
        Case 6 : MsgBox txt, 48, "PLIK NIE ZNALEZIONY"
        Case 7 : MsgBox txt, 16, "BLAD URZADZENIA"
        Case 8 : MsgBox txt, 64, "BLAD SIECI"
        Case 9 : MsgBox txt, 32, "POMOC SYSTEMU"
        Case 10: MsgBox txt, Int(Rnd() * 5) * 16, "LOSOWY BLAD"
        Case Else : MsgBox "Nieprawidlowy wybor.", vbExclamation, "Blad"
    End Select
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
