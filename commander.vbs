Set shell = CreateObject("WScript.Shell")
' FIX: Const nie akceptuje funkcji. Uzywamy Dim i przypisania w czasie dzialania.
Dim SEP : SEP = String(50, "=")
Dim TITLE : TITLE = "VM SYSTEM COMMANDER v2.1 [PRO]"

Do
    menu = SEP & vbCrLf & _
           "        " & TITLE & vbCrLf & _
           SEP & vbCrLf & _
           "1. Narzedzia Systemowe" & vbCrLf & _
           "2. Laboratorium Bledow" & vbCrLf & _
           "3. Opcje Zaawansowane" & vbCrLf & _
           "4. Menadzer Procesow" & vbCrLf & _
           "5. Czyszczenie Systemu" & vbCrLf & _
           "6. Zarzadzanie Internetem" & vbCrLf & _
           "7. Ustawienia Zasilania" & vbCrLf & _
           "8. Informacje o Systemie" & vbCrLf & _
           "9. Efekt Matrix" & vbCrLf & _
           "10. Narzedzia Naprawcze" & vbCrLf & _
           "11. Wyjdz" & vbCrLf & _
           SEP

    wybor = InputBox(menu, TITLE, "1")
    If wybor = "" Or wybor = "11" Then Exit Do

    Select Case wybor
        Case "1"  : shell.Run "perfmon /res", 1
        Case "2"  : RunErrorLab
        Case "3"  : RunAdvanced
        Case "4"  : RunProcessKill
        Case "5"  : RunCleaner
        Case "6"  : RunNetwork
        Case "7"  : RunPower
        Case "8"  : RunSysInfo
        Case "9"  : RunMatrix
        Case "10" : RunRepair
        Case Else : MsgBox "Nieprawidlowy wybor. Sprobuj ponownie.", vbExclamation, "Blad"
    End Select
Loop
MsgBox "Do zobaczenia!", vbInformation, "Wyjscie"
WScript.Quit

' ================= SUBROUTINES =================

Sub RunErrorLab
    errMenu = "LABORATORIUM BLEDOW:" & vbCrLf & _
              "1. Blad Krytyczny" & vbCrLf & _
              "2. Ostrzezenie" & vbCrLf & _
              "3. Informacja" & vbCrLf & _
              "4. Pytanie"
    errWybor = InputBox(errMenu, "Error Lab", "1")
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
End Sub

Sub RunAdvanced
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
End Sub

Sub RunProcessKill
    shell.Run "taskkill /f /fi ""STATUS eq RUNNING"" /fi ""USERNAME ne NT AUTHORITY\SYSTEM""", 0
    shell.Run "explorer.exe", 1
    MsgBox "Procesy uzytkownika zatrzymane. Eksplorator uruchomiony ponownie.", vbInformation, "Menadzer"
End Sub

Sub RunCleaner
    cleanConfirm = MsgBox("Czyszczenie usunie:" & vbCrLf & _
                          "- %TEMP% (pliki tymczasowe)" & vbCrLf & _
                          "- %WINDIR%\Temp (pliki systemowe)" & vbCrLf & _
                          "- %WINDIR%\Prefetch (cache)" & vbCrLf & _
                          "- Cache DNS" & vbCrLf & vbCrLf & _
                          "Kontynuowac?", vbQuestion + vbYesNo, "Czyszczenie")
    If cleanConfirm = vbYes Then
        shell.Run "cmd /c ""del /f /s /q %TEMP%\*.* 2>nul & del /f /s /q %WINDIR%\Temp\*.* 2>nul & del /f /s /q %WINDIR%\Prefetch\*.* 2>nul & ipconfig /flushdns 2>nul""", 0, True
        MsgBox "Czyszczenie zakonczone pomyslnie!", vbInformation, "Sukces"
    End If
End Sub

Sub RunNetwork
    netMenu = "ZARZADZANIE INTERNETEM:" & vbCrLf & _
              "1. Sprawdz IP Lokalny" & vbCrLf & _
              "2. Sprawdz IP Publiczny" & vbCrLf & _
              "3. Wylacz Siec (Adapter)" & vbCrLf & _
              "4. Wlacz Siec (Adapter)" & vbCrLf & _
              "5. Reset Stosu Sieciowego" & vbCrLf & _
              "6. Ping Test"
    netWybor = InputBox(netMenu, "Network", "1")
    Select Case netWybor
        Case "1"
            Set exec = shell.Exec("cmd /c ipconfig | findstr /C:""IPv4""")
            MsgBox "Lokalne adresy IP:" & vbCrLf & exec.StdOut.ReadAll, vbInformation, "IP Lokalny"
        Case "2"
            MsgBox "Sprawdzanie... (moze potrwac kilka sekund)", vbInformation, "IP Publiczny"
            Set exec = shell.Exec("powershell -Command ""(Invoke-RestMethod http://ifconfig.me/ip).Trim()""")
            MsgBox "Publiczny adres IP:" & vbCrLf & exec.StdOut.ReadAll, vbInformation, "IP Publiczny"
        Case "3"
            adapter = InputBox("Nazwa adaptera (np. Wi-Fi, Ethernet):", "Wylaczanie", "Wi-Fi")
            If adapter <> "" Then shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command Disable-NetAdapter -Name ''' & adapter & ''' -Confirm:$false'""", 0
        Case "4"
            adapter = InputBox("Nazwa adaptera (np. Wi-Fi, Ethernet):", "Wlaczanie", "Wi-Fi")
            If adapter <> "" Then shell.Run "powershell -Command ""Start-Process powershell -Verb RunAs -ArgumentList '-Command Enable-NetAdapter -Name ''' & adapter & ''' -Confirm:$false'""", 0
        Case "5"
            MsgBox "Resetowanie stosu sieciowego...", vbExclamation, "Reset"
            shell.Run "cmd /c netsh winsock reset && netsh int ip reset && ipconfig /release && ipconfig /renew", 0, True
            MsgBox "Reset zakoczony. Uruchom ponownie komputer.", vbInformation, "Zakonczone"
        Case "6"
            host = InputBox("Adres do pingowania:", "Ping Test", "8.8.8.8")
            If host <> "" Then
                Set exec = shell.Exec("ping -n 4 " & host)
                MsgBox exec.StdOut.ReadAll, vbInformation, "Ping: " & host
            End If
    End Select
End Sub

Sub RunPower
    powMenu = "USTAWIENIA ZASILANIA:" & vbCrLf & _
              "1. Restart do BIOS/UEFI" & vbCrLf & _
              "2. Wyloguj uzytkownika" & vbCrLf & _
              "3. Uruchom ponownie" & vbCrLf & _
              "4. Zamknij system" & vbCrLf & _
              "5. Slide to Shut Down" & vbCrLf & _
              "6. Anuluj oczekujaca akcje"
    powWybor = InputBox(powMenu, "Zasilanie", "4")
    Select Case powWybor
        Case "1" : shell.Run "powershell -Command ""Start-Process shutdown -ArgumentList '/r /fw /t 0' -Verb RunAs""", 0
        Case "2"
            timeIn = InputBox("Opoznienie wylogowania (sekundy):", "Czas", "10")
            If timeIn = "" Or Not IsNumeric(timeIn) Then timeIn = "10"
            shell.Run "cmd /c timeout /t " & timeIn & " /nobreak >nul && shutdown /l", 0
        Case "3"
            timeIn = InputBox("Czas do restartu (sekundy):", "Czas", "30")
            If timeIn = "" Or Not IsNumeric(timeIn) Then timeIn = "30"
            msgIn = InputBox("Komunikat na ekranie:", "Tekst", "VM Commander restartuje system...")
            If msgIn = "" Then msgIn = "VM Commander restartuje system..."
            shell.Run "shutdown /r /t " & timeIn & " /c """ & msgIn & """", 0
        Case "4"
            timeIn = InputBox("Czas do zamkniecia (sekundy):", "Czas", "30")
            If timeIn = "" Or Not IsNumeric(timeIn) Then timeIn = "30"
            msgIn = InputBox("Komunikat na ekranie:", "Tekst", "VM Commander zamyka system...")
            If msgIn = "" Then msgIn = "VM Commander zamyka system..."
            shell.Run "shutdown /s /t " & timeIn & " /c """ & msgIn & """", 0
        Case "5" : shell.Run "SlideToShutDown.exe", 1
        Case "6"
            shell.Run "shutdown /a", 0
            MsgBox "Zaplanowana akcja anulowana!", vbInformation, "Anulowano"
    End Select
End Sub

Sub RunSysInfo
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set os = wmi.ExecQuery("Select * from Win32_OperatingSystem")
    For Each item In os
        osName = item.Caption
        ramGB = FormatNumber(item.TotalVisibleMemorySize / 1048576, 2)
        lastBoot = item.LastBootUpTime
        bootTime = Mid(lastBoot, 7, 2) & "." & Mid(lastBoot, 5, 2) & "." & Mid(lastBoot, 1, 4) & " " & Mid(lastBoot, 9, 2) & ":" & Mid(lastBoot, 11, 2)
    Next
    Set comp = wmi.ExecQuery("Select * from Win32_ComputerSystem")
    For Each item In comp
        compName = item.Name
    Next
    On Error GoTo 0
    infoBox = "SYSTEM INFO:" & vbCrLf & _
              "Nazwa: " & compName & vbCrLf & _
              "System: " & osName & vbCrLf & _
              "RAM: " & ramGB & " GB" & vbCrLf & _
              "Uruchomiony od: " & bootTime
    MsgBox infoBox, vbInformation, "Informacje o Systemie"
End Sub

Sub RunMatrix
    Dim fso, batFile, batContent
    Set fso = CreateObject("Scripting.FileSystemObject")
    batFile = fso.GetSpecialFolder(2) & "\vm_matrix.bat"
    batContent = "@echo off" & vbCrLf & _
                 "color 0a" & vbCrLf & _
                 ":start" & vbCrLf & _
                 "echo %random%%random% %random%%random% %random%%random% %random%%random% %random%%random%" & vbCrLf & _
                 "goto start"
    fso.CreateTextFile(batFile, True).Write batContent
    shell.Run "cmd /k """ & batFile & """", 1
    MsgBox "Zamknij okno CMD, aby zakonczyc efekt Matrix.", vbInformation, "Matrix"
    fso.DeleteFile batFile, True
End Sub

Sub RunRepair
    repMenu = "NARZEDZIA NAPRAWCZE:" & vbCrLf & _
              "1. SFC Scan (Naprawa plikow systemowych)" & vbCrLf & _
              "2. DISM Restore (Obraz systemu)" & vbCrLf & _
              "3. Oczyszczanie dysku (cleanmgr)"
    repWybor = InputBox(repMenu, "Naprawa", "1")
    Select Case repWybor
        Case "1"
            MsgBox "Uruchamianie SFC. Okno CMD otworzy sie i wykona skan. Nie zamykaj go recznie.", vbInformation, "SFC"
            shell.Run "cmd /k sfc /scannow", 1
        Case "2"
            MsgBox "Uruchamianie DISM. Moze potrwac 5-10 minut.", vbInformation, "DISM"
            shell.Run "cmd /k DISM /Online /Cleanup-Image /RestoreHealth", 1
        Case "3"
            shell.Run "cleanmgr", 1
    End Select
End Sub
