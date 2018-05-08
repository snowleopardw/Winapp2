Option Strict On
Imports System.IO

Module Diff

    'File handlers
    Dim oFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "winapp2.ini")
    Dim nFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "winapp2.ini", "")
    Dim logFile As IFileHandlr = New IFileHandlr(Environment.CurrentDirectory, "diff.txt")
    Dim oldFile As iniFile
    Dim newFile As iniFile
    Dim outputToFile As String

    'Menu settings
    Dim settingsChanged As Boolean = False

    'Boolean module parameters
    Dim downloadFile As Boolean = False
    Dim downloadNCC As Boolean = False
    Dim saveLog As Boolean = False

    'Return the default parameters to the commandline handler
    Public Sub initDiffParams(ByRef firstFile As IFileHandlr, ByRef secondFile As IFileHandlr, ByRef thirdFile As IFileHandlr, ByRef d As Boolean, ByRef dncc As Boolean, ByRef sl As Boolean)
        initDefaultSettings()
        firstFile = oFile
        secondFile = nFile
        thirdFile = logFile
        d = downloadFile
        dncc = downloadNCC
        sl = saveLog
    End Sub

    'Handle calling Diff from the commandline with 
    Public Sub remoteDiff(ByRef firstFile As IFileHandlr, secondFile As IFileHandlr, thirdFile As IFileHandlr, d As Boolean, dncc As Boolean, sl As Boolean)
        oFile = firstFile
        nFile = secondFile
        logFile = thirdFile
        downloadFile = d
        downloadNCC = dncc
        saveLog = sl
        initDiff()
    End Sub

    'Restore all the module settings to their default state
    Private Sub initDefaultSettings()
        oFile.resetParams()
        nFile.resetParams()
        logFile.resetParams()
        downloadFile = False
        downloadNCC = False
        saveLog = False
        settingsChanged = False
    End Sub

    Private Sub resetSettings()
        initDefaultSettings()
        menuTopper = "Diff settings have been reset to their defaults"
    End Sub

    Private Sub printMenu()
        printMenuTop({"Observe the differences between two ini files"})
        printMenuOpt("1. Run (default)", "Run the diff tool")
        printBlankMenuLine()
        printMenuLine("Select Older/Local File:", "l")
        printMenuLine("2. Winapp2.ini", "l")
        printMenuOpt("3. File Chooser (older file)", "Choose a new name or location for your older ini file")
        printBlankMenuLine()
        printMenuLine("Select Newer/Remote File:", "l")
        printMenuOpt("4. winapp2.ini (online)", "Diff against the latest winapp2.ini version on GitHub")
        printMenuOpt("5. winapp2.ini (non-ccleaner)", "Diff against the latest non-ccleaner winapp2.ini version on GitHub")
        printMenuOpt("6. File Chooser (newer file)", "Choose a new name of location for your newer ini file")
        printBlankMenuLine()
        printMenuLine("Log Settings:", "l")
        printMenuOpt("7. Toggle Log Saving", If(saveLog, "Disable", "Enable") & " automatic saving of the Diff output")
        If saveLog Then printMenuOpt("8. File Chooser (log)", "Change where Diff saves its log")

        printBlankMenuLine()
        printMenuLine("Older file: " & replDir(oFile.path), "l")

        printMenuLine("Newer file: " & If(nFile.name = "", "Not yet selected", replDir(nFile.path)), "l")
        If settingsChanged Then
            printBlankMenuLine()
            printMenuOpt(getMenuNumber({saveLog}, 8) & ". Reset Settings", "Restore the default state of the Diff settings")
        End If
        printMenuLine(menuStr02)
    End Sub

    Public Sub main()
        initMenu("Diff", 35)
        Console.WindowHeight = 40
        outputToFile = ""

        Do Until exitCode
            Console.Clear()
            printMenu()
            Console.WriteLine()
            Console.Write("Enter a number, or leave blank to run the default: ")

            Dim input As String = Console.ReadLine()
            Try
                Select Case True
                    Case input = "0"
                        Console.WriteLine("Exiting diff...")
                        exitCode = True
                    Case input = "1" Or input = ""
                        initDiff()
                    Case input = "2"
                        oFile.name = "winapp2.ini"
                    Case input = "3"
                        changeFileParams(oFile, settingsChanged)
                        If exitCode Then revertMenu()
                    Case input = "4"
                        toggleSettingParam(downloadFile, "Download ", settingsChanged)
                        nFile.name = "Online"
                    Case input = "5"

                        Select Case True
                            Case (Not downloadFile And Not downloadNCC) Or (downloadFile And downloadNCC)
                                toggleSettingParam(downloadFile, "Download ", settingsChanged)
                                toggleSettingParam(downloadNCC, "Download Non-CCleaner ", settingsChanged)
                            Case downloadFile And Not downloadNCC
                                toggleSettingParam(downloadNCC, "Download Non-CCleaner ", settingsChanged)
                        End Select

                        nFile.name = If(downloadNCC, "Online (non-ccleaner)", "")
                    Case input = "6"
                        changeFileParams(nFile, settingsChanged)
                        If exitCode Then revertMenu()
                    Case input = "7"
                        toggleSettingParam(saveLog, "Log Saving ", settingsChanged)
                    Case input = "8" And (saveLog Or settingsChanged)
                        If saveLog Then
                            changeFileParams(logFile, settingsChanged)
                        Else
                            resetSettings()
                        End If
                    Case input = "9" And (settingsChanged And saveLog)
                        resetSettings()
                    Case Else
                        menuTopper = invInpStr
                End Select
            Catch ex As Exception
                exc(ex)
            End Try
        Loop
        revertMenu()
        Console.WindowHeight = 30
    End Sub

    Private Sub initDiff()
        oldFile = validate(oFile)
        If downloadFile Then newFile = getRemoteWinapp(downloadNCC)
        differ()
        If saveLog Then saveDiff()
        revertMenu()
        Console.Clear()
    End Sub

    Private Sub differ()
        If exitCode Then Exit Sub
        Console.Clear()
        Try
            'collect & verify version #s and print them out for the menu
            Dim fver As String = oldFile.comments(0).comment.ToString
            fver = If(fver.ToLower.Contains("version"), fver.TrimStart(CChar(";")).Replace("Version:", "version"), " version not given")

            Dim sver As String = newFile.comments(0).comment.ToString
            sver = If(sver.ToLower.Contains("version"), sver.TrimStart(CChar(";")).Replace("Version:", "version"), " version not given")

            outputToFile += tmenu("Changes made between" & fver & " and" & sver) & Environment.NewLine
            outputToFile += menu(menuStr02) & Environment.NewLine
            Console.WriteLine()
            outputToFile += menu(menuStr00) & Environment.NewLine

            'compare the files and then ennumerate their changes
            Dim outList As List(Of String) = compareTo()
            Dim remCt As Integer = 0
            Dim modCt As Integer = 0
            Dim addCt As Integer = 0
            For Each change In outList

                If change.Contains("has been added.") Then
                    addCt += 1
                ElseIf change.Contains("has been removed") Then
                    remCt += 1
                Else
                    modCt += 1
                End If
                outputToFile += change & Environment.NewLine
            Next

            outputToFile += menu("Diff complete.", "c") & Environment.NewLine
            outputToFile += menu(menuStr03) & Environment.NewLine
            outputToFile += menu("Added entries: " & addCt, "l") & Environment.NewLine
            outputToFile += menu("Modified entries: " & modCt, "l") & Environment.NewLine
            outputToFile += menu("Removed entries: " & remCt, "l") & Environment.NewLine
            outputToFile += menu(menuStr02)
            If Not suppressOutput Then Console.Write(outputToFile)
        Catch ex As Exception
            If ex.Message = "The given key was not present in the dictionary." Then
                Console.WriteLine("Error encountered during diff: " & ex.Message)
                Console.WriteLine("This error is typically caused by invalid file names, please double check your input and try again.")
                Console.WriteLine()
            Else
                exc(ex)
            End If
        End Try

        Console.WriteLine()
        printMenuLine(bmenu("Press any key to return to the winapp2ool menu.", "l"))
        Console.ReadKey()
    End Sub

    Private Function compareTo() As List(Of String)

        Dim outList, comparedList As New List(Of String)

        For Each section In oldFile.sections.Values
            Try

                'If we're looking at an entry in the old file and the new file contains it, and we haven't yet processed this entry
                If newFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                    Dim sSection As iniSection = newFile.sections(section.name)

                    'and if that entry in the new file does not compareTo the entry in the old file, we have a modified entry
                    Dim addedKeys, removedKeys As New List(Of iniKey)
                    Dim updatedKeys As New List(Of KeyValuePair(Of iniKey, iniKey))
                    If Not section.compareTo(sSection, removedKeys, addedKeys, updatedKeys) Then
                        Dim tmp As String = getDiff(sSection, "modified.")
                        If addedKeys.Count > 0 Then
                            tmp += Environment.NewLine & "Added:"
                            For Each key In addedKeys
                                tmp += Environment.NewLine & key.toString
                            Next
                        End If
                        If removedKeys.Count > 0 Then
                            tmp += Environment.NewLine & Environment.NewLine & "Removed:"
                            For Each key In removedKeys
                                tmp += Environment.NewLine & key.toString
                            Next
                        End If
                        If updatedKeys.Count > 0 Then
                            tmp += If(removedKeys.Count > 0 Or addedKeys.Count > 0, Environment.NewLine & Environment.NewLine, Environment.NewLine) & "Modified:" & Environment.NewLine
                            For Each pair In updatedKeys
                                tmp += Environment.NewLine & pair.Key.name & Environment.NewLine

                                tmp += "old:   " & pair.Key.toString & Environment.NewLine
                                tmp += "new:   " & pair.Value.toString & Environment.NewLine
                            Next
                        End If
                        tmp += Environment.NewLine & Environment.NewLine & menuStr00
                        outList.Add(tmp)
                    End If

                ElseIf Not newFile.sections.Keys.Contains(section.name) And Not comparedList.Contains(section.name) Then
                    'If we do not have the entry in the new file, it has been removed between versions 
                    outList.Add(getDiff(section, "removed."))
                End If
                comparedList.Add(section.name)
            Catch ex As Exception
                exc(ex)
            End Try
        Next

        For Each section In newFile.sections.Values
            'Any sections from the new file which are not found in the old file have been added
            If Not oldFile.sections.Keys.Contains(section.name) Then outList.Add(getDiff(section, "added."))
        Next

        Return outList
    End Function

    Private Function getDiff(section As iniSection, changeType As String) As String
        'Return a string containing a box containing the change type and entry name, followed by the entry's tostring
        Dim out As String = ""
        out += mkMenuLine(section.name & " has been " & changeType, "c") & Environment.NewLine
        out += mkMenuLine(menuStr02, "") & Environment.NewLine & Environment.NewLine
        out += section.ToString & Environment.NewLine
        If Not changeType = "modified." Then out += menuStr00
        Return out
    End Function

    Private Sub saveDiff()
        'Save diff.txt 
        Try
            Dim file As New StreamWriter(Environment.CurrentDirectory & "\diff.txt", False)
            file.Write(outputToFile)
            file.Close()
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
End Module