Option Strict On
Imports System.IO
Module Winapp2ool

    Dim currentVersion As Double = 0.85
    Dim latestVersion As String

    Dim checkedForUpdates As Boolean = False
    Dim updateIsAvail As Boolean = False

    Dim latestWa2Ver As String = ""
    Dim localWa2Ver As String = ""

    Dim dnfOOD As Boolean = False

    'This boolean will prevent us from printing or asking for input under most circumstances, triggered by the -s command line argument 
    Public suppressOutput As Boolean = False

    Dim waUpdateIsAvail As Boolean = False

    Private Sub printUpdNotif(updName As String, oldVer As String, newVer As String)
        printBlankMenuLine()
        printMenuLine("A new version of " & updName & "is available!", "c")
        printMenuLine("Current:   v" & oldVer, "c")
        printMenuLine("Available: v" & newVer, "c")
    End Sub

    Private Sub printMenu()
        printMenuLine(tmenu(menuTopper))
        printMenuLine(menuStr03)
        If dnfOOD Then printMenuLine("Your .NET Framework is out of date. Please update to version 4.6+.", "c")
        If updateIsAvail Then printUpdNotif("Winapp2ool ", currentVersion.ToString, latestVersion)
        If waUpdateIsAvail And Not localWa2Ver = "0" Then printUpdNotif("winapp2.ini ", localWa2Ver, latestWa2Ver)
        printMenuLine(menuStr04)
        printMenuOpt("0. Exit", "Exit the application")
        printMenuOpt("1. WinappDebug", "Check for and correct errors in winapp2.ini")
        printMenuOpt("2. Trim", "Debloat winapp2.ini for your system")
        printMenuOpt("3. Merge", "Merge the contents of an ini file into winapp2.ini")
        printMenuOpt("4. Diff", "Observe the changes between two winapp2.ini files")
        printMenuOpt("5. CCiniDebug", "Sort and trim ccleaner.ini")
        printBlankMenuLine()
        printMenuOpt("6. Downloader", "Download files from the Winapp2 GitHub")
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()

        Console.Title = "Winapp2ool v" & currentVersion & " beta"
        Console.WindowWidth = 120
        initMenu("Winapp2ool - A multitool for winapp2.ini and related files", 35)
        processCommandLineArgs()

        If suppressOutput Then Environment.Exit(1)

        checkUpdates()
        Do Until exitCode = True
            Console.Clear()
            printMenu()
            cwl()
            Console.Write("Enter a number: ")
            Dim input As String = Console.ReadLine

            Select Case input
                Case "0"
                    exitCode = True
                    cwl("Exiting...")
                    Environment.Exit(1)
                Case "1"
                    WinappDebug.main()
                    menuTopper = "Finished running WinappDebug"
                Case "2"
                    Trim.main()
                    menuTopper = "Finished running Trim"
                Case "3"
                    Merge.main()
                    menuTopper = "Finished running Merge"
                Case "4"
                    Diff.main()
                    menuTopper = "Finished running Diff"
                Case "5"
                    CCiniDebug.Main()
                    menuTopper = "Finished running CCiniDebug"
                Case "6"
                    Downloader.main()
                    menuTopper = "Finished running Downloader"
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    Private Sub checkUpdates()
        Try
            'Check for winapp2ool.exe updates
            latestVersion = getRemoteFileDataAtLineNum(toolVerLink, 1)
            updateIsAvail = CDbl(latestVersion) > currentVersion
            If Not Environment.Version.ToString = "4.0.30319.42000" Then dnfOOD = True

            'Check for winapp2.ini updates
            latestWa2Ver = getRemoteFileDataAtLineNum(wa2Link, 1).Split(CChar(" "))(2)
            If Not File.Exists(Environment.CurrentDirectory & "\winapp2.ini") Then
                localWa2Ver = "0"
                menuTopper = "/!\ Winapp2.ini update check failed. /!\"
            Else
                localWa2Ver = getFileDataAtLineNum(Environment.CurrentDirectory & "\winapp2.ini", 1).Split(CChar(" "))(2)
            End If
            waUpdateIsAvail = CDbl(latestWa2Ver) > CDbl(localWa2Ver)

        Catch ex As Exception
            menuTopper = "/!\ Update check failed. /!\"
        End Try
    End Sub

    Public Sub exc(ByRef ex As Exception)
        If ex.Message.ToString.Contains("SSL/TLS") Then
            cwl("Error: download could not be completed.")
            cwl("This issue is caused by an out of date .NET Framework.")
            cwl("Please update .NET Framework to version 4.6 or higher and try again.")
            cwl("If the issue persists after updating .NET Framework, please report this error on GitHub.")
        Else
            cwl("Error: " & ex.ToString)
            cwl("Please report this error on GitHub")
            cwl()
        End If


    End Sub

    Public Sub cwl()
        If Not suppressOutput Then Console.WriteLine()
    End Sub

    Public Sub cwl(msg As String)
        If Not suppressOutput Then Console.WriteLine(msg)
    End Sub

    'Prompt the user to change a file's parameters, flag it as changed, and mark settings as having changed
    Public Sub changeFileParams(ByRef someFile As IFileHandlr, ByRef settingsChangedSetting As Boolean)
        fChooser(someFile.dir, someFile.name, someFile.initName, someFile.secondName)
        settingsChangedSetting = True
        menuTopper = If(someFile.secondName = "", someFile.initName, "save file") & " parameters updated"
    End Sub

    'Toggle a parameter on or off and mark settings as having changed
    Public Sub toggleSettingParam(ByRef setting As Boolean, paramText As String, ByRef settingsChangedSetting As Boolean)
        setting = Not setting
        menuTopper = paramText & If(setting, "enabled", "disabled")
        settingsChangedSetting = True
    End Sub

End Module