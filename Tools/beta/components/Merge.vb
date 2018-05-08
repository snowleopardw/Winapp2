﻿Option Strict On
Imports System.IO

Module Merge

    'File handlers
    Dim winappFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "winapp2.ini")
    Dim mergeFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "")
    Dim outputFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "winapp2.ini", "winapp2-merged.ini")
    Dim firstFile As iniFile
    Dim secondFile As iniFile

    'Menu settings
    Dim settingsChanged As Boolean = False

    'Boolean module parameters
    Dim mergeMode As Boolean = True

    'Return the default parameter states to the command line handler
    Public Sub initMergeParams(ByRef firstFile As IFileHandlr, ByRef secondFile As IFileHandlr, ByRef thirdFile As IFileHandlr, ByRef mm As Boolean)
        initDefaultSettings()
        firstFile = winappFile
        secondFile = mergeFile
        thirdFile = outputFile
        mm = mergeMode
    End Sub

    'Handle commands from command line ans initalize the merger
    Public Sub remoteMerge(firstFile As IFileHandlr, secondFile As IFileHandlr, thirdFile As IFileHandlr, mm As Boolean)
        winappFile = firstFile
        mergeFile = secondFile
        outputFile = thirdFile
        mergeMode = mm
        initMerge()
    End Sub

    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "Merge settings have been reset to their defaults"
    End Sub

    Private Sub initDefaultSettings()
        winappFile.resetParams()
        mergeFile.resetParams()
        outputFile.resetParams()
        settingsChanged = False
    End Sub

    Private Sub printMenu()
        printMenuTop({"Merge the contents of two ini files, while either replacing (default) or removing sections with the same name."})
        printMenuOpt("1. Run (default)", "Merge the two ini files")
        printBlankMenuLine()
        printMenuLine("Preset Merge File Choices:", "l")
        printBlankMenuLine()
        printMenuOpt("2. Removed Entries", "Select Removed Entries.ini")
        printMenuOpt("3. Custom", "Select Custom.ini")
        printBlankMenuLine()
        printMenuOpt("4. File Chooser (winapp2.ini)", "Choose a new name or location for winapp2.ini")
        printMenuOpt("5. File Chooser (merge)", "Choose a name or location not listed above for merging")
        printMenuOpt("6. File Chooser (save)", "Choose a new save location for the merged file")
        printBlankMenuLine()
        printMenuLine("Current winapp2.ini: " & replDir(winappFile.path), "l")
        If Not mergeFile.name = "" Then printMenuLine("Current merge file: " & replDir(mergeFile.path), "l")
        printMenuLine("Current save target: " & replDir(outputFile.path), "l")
        printBlankMenuLine()
        printMenuOpt("7. Toggle Merge Mode", "Switch between merge modes. Current mode: " & If(mergeMode, "Replace & Add", "Replace & Remove"))
        If settingsChanged Then
            printBlankMenuLine()
            printMenuOpt("8. Reset Settings", "Restore the default Merge settings")
        End If
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        initMenu("Merge", 35)
        mergeMode = True
        While Not exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")
            Dim input As String = Console.ReadLine
            Select Case input
                Case "0"
                    exitCode = True
                Case "1", ""
                    If mergeFile.name <> "" Then
                        initMerge()
                    Else
                        menuTopper = "You must select a file to merge"
                    End If
                Case "2"
                    mergeFile.name = "Removed entries.ini"
                    settingsChanged = True
                Case "3"
                    mergeFile.name = "Custom.ini"
                    settingsChanged = True
                Case "4"
                    changeFileParams(winappFile, settingsChanged)
                Case "5"
                    changeFileParams(mergeFile, settingsChanged)
                Case "6"
                    changeFileParams(outputFile, settingsChanged)
                Case "7"
                    toggleSettingParam(mergeMode, "Merge Mode ", settingsChanged)
                Case "8"
                    If settingsChanged Then
                        resetSettings()
                    Else
                        menuTopper = invInpStr
                    End If
                Case Else
                    menuTopper = invInpStr
            End Select
        End While
        revertMenu()
    End Sub

    Public Sub initMerge()
        Console.Clear()
        'Initialize our inifiles
        firstFile = validate(winappFile)
        If exitCode Then Exit Sub
        secondFile = validate(mergeFile)
        If exitCode Then Exit Sub

        'Merge them
        merge()

        'Flip our menu boolean
        revertMenu()
    End Sub

    Private Sub merge()
        Dim out As String = ""

        'Process the merge mode and update the inifiles accordingly
        processMergeMode(firstFile, secondFile)

        'Parse our two files
        Dim tmp As New winapp2file(firstFile)
        Dim tmp2 As New winapp2file(secondFile)

        printMenuLine(bmenu("Merging " & firstFile.name.TrimStart(CChar("\")) & " with " & secondFile.name.TrimStart(CChar("\")), "c"))

        Try

            'Add the entries from the second file to their respective sections in the first file
            tmp.cEntriesW.AddRange(tmp2.cEntriesW)
            tmp.fxEntriesW.AddRange(tmp2.fxEntriesW)
            tmp.tbEntriesW.AddRange(tmp2.tbEntriesW)
            tmp.mEntriesW.AddRange(tmp2.mEntriesW)

            'Rebuild the internal changes
            tmp.rebuildToIniFiles()

            'Sort the merged sections 
            sortIniFile(tmp.cEntries, replaceAndSort(tmp.cEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tmp.fxEntries, replaceAndSort(tmp.fxEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tmp.tbEntries, replaceAndSort(tmp.tbEntries.getSectionNamesAsList, "-", "  "))
            sortIniFile(tmp.mEntries, replaceAndSort(tmp.mEntries.getSectionNamesAsList, "-", "  "))

            Dim file As New StreamWriter(outputFile.path, False)

            'write the merged winapp2string to file
            out += tmp.winapp2string
            file.Write(out)
            file.Close()
        Catch ex As Exception
            exc(ex)
        End Try

        printMenuLine(bmenu("Finished merging files. Press any key to return to the menu", "c"))
        If Not suppressOutput Then Console.ReadKey()
    End Sub

    Private Sub processMergeMode(ByRef first As iniFile, ByRef second As iniFile)
        Dim removeList As New List(Of String)

        'If mergemode is true, replace any matches
        If mergeMode Then
            For Each section In second.sections.Keys
                If first.sections.Keys.Contains(section) Then
                    first.sections.Item(section) = second.sections.Item(section)
                    removeList.Add(section)
                End If
            Next
            For Each section In removeList
                second.sections.Remove(section)
            Next
        Else
            'if mergemode is false, remove any matches
            For Each section In second.sections.Keys
                If first.sections.Keys.Contains(section) Then
                    first.sections.Remove(section)
                    removeList.Add(section)
                End If
            Next
        End If

        'Remove any processed sections from the second file so that only entries to add remain 
        For Each section In removeList
            second.sections.Remove(section)
        Next
    End Sub
End Module