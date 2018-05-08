Option Strict On
Imports System.IO
Imports System.Text.RegularExpressions

Public Module iniFileHandler


    Public Class IFileHandlr
        'This class simply holds some information about files we want to load such that we can modify their parts or retain initalized information easily without needing to write the values many times

        Public dir As String
        Public name As String
        Public initDir As String
        Public initName As String
        Public secondName As String

        Public Sub New(directory As String, filename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = ""
        End Sub

        Public Sub New(directory As String, filename As String, rename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = rename
        End Sub

        Public Sub resetParams()
            dir = initDir
            name = initName
        End Sub

        Public Function path() As String
            Return dir & "\" & name
        End Function

        Public Function filename() As String
            Return "\" & name
        End Function

        Public Function initFileName() As String
            Return "\" & initName
        End Function

        Public Function fileRename() As String
            If secondName <> "" Then
                Return "\" & secondName
            Else
                Return ""
            End If
        End Function

    End Class

    'Strips just the key values from a list of inikeys and returns them as a list of strings
    Public Function getValuesFromKeyList(ByVal keyList As List(Of iniKey)) As List(Of String)
        Dim valList As New List(Of String)
        For Each key In keyList
            valList.Add(key.value)
        Next

        Return valList
    End Function

    'Strips just the line numbers from a list of inikeys and returns them as a list of integers
    Public Function getLineNumsFromKeyList(ByRef keyList As List(Of iniKey)) As List(Of Integer)
        Dim lineList As New List(Of Integer)
        For Each key In keyList
            lineList.Add(key.lineNumber)
        Next

        Return lineList
    End Function

    'Strips just the line numbers from the sections of an inifile and returns them as a list of integers
    Public Function getLineNumsFromSections(ByVal file As iniFile) As List(Of Integer)
        Dim outList As New List(Of Integer)
        For i As Integer = 0 To file.sections.Count - 1
            outList.Add(file.sections.Values(i).startingLineNumber)
        Next

        Return outList
    End Function

    'reorders the sections in an inifile to be in the same order as some sorted state provided to the function
    Public Sub sortIniFile(ByRef fileToBeSorted As iniFile, ByVal sortedKeys As List(Of String))
        Dim tempFile As New iniFile
        For Each entryName In sortedKeys
            tempFile.sections.Add(entryName, fileToBeSorted.sections.Item(entryName))
        Next

        fileToBeSorted = tempFile
    End Sub

    'Ensures that any call to an ini file on the system will be to a file that exists in a directory that exists
    Public Function validate(ByRef someFile As IFileHandlr) As iniFile

        'if there's a pending exit, do that.
        If exitCode Then Return Nothing

        'Make sure both the file and the directory actually exist
        While Not File.Exists(someFile.path)
            If Not Directory.Exists(someFile.dir) Then
                dChooser(someFile.dir)
            End If
            If exitCode Then Return Nothing
            If Not File.Exists(someFile.path) Then
                chkFileExist(someFile)
            End If
            If exitCode Then Return Nothing
        End While

        'Make sure that the file isn't empty
        Try
            Dim iniTester As New iniFile(someFile.dir, someFile.name)
            Dim clearAtEnd As Boolean = False
            While iniTester.sections.Count = 0
                clearAtEnd = True
                Console.Clear()
                printMenuLine(bmenu("Empty ini file detected. Press any key to try again.", "c"))
                Console.ReadKey()
                fChooser(someFile.dir, someFile.name, someFile.initName, someFile.secondName)
                If exitCode Then Return Nothing
                iniTester = validate(someFile)
                If exitCode Then Return Nothing
            End While
            If clearAtEnd Then Console.Clear()

            Return iniTester
        Catch ex As Exception
            exc(ex)
            exitCode = True
            Return Nothing
        End Try
    End Function

    Public Sub chkFileExist(someFile As IFileHandlr)
        Dim iExitCode As Boolean = False

        While Not File.Exists(someFile.path)
            If exitCode Then Exit Sub
            menuTopper = "Error"
            While Not iExitCode
                Console.Clear()
                printMenuTop({someFile.name & " does not exist."})
                printMenuOpt("1. File Chooser (default)", "Change the file name")
                printMenuOpt("2. Directory Chooser", "Change the directory")
                printMenuLine(menuStr02)
                Dim input As String = Console.ReadLine
                Select Case input
                    Case "0"
                        iExitCode = True
                        exitCode = True
                        Exit Sub
                    Case "1", ""
                        fChooser(someFile.dir, someFile.name, someFile.initName, someFile.secondName)
                    Case "2"
                        dChooser(someFile.dir)
                    Case Else
                        menuTopper = invInpStr
                End Select
                If Not File.Exists(someFile.path) And Not menuTopper = invInpStr Then menuTopper = "Error"
            End While
        End While
    End Sub

    Public Sub chkDirExist(ByRef dir As String)
        If exitCode Then Exit Sub
        menuTopper = "Error"
        While Not Directory.Exists(dir)
            If exitCode Then Exit Sub
            Dim iExitCode As Boolean = False

            While Not iExitCode
                printMenuTop({dir & "does not exist."})
                printMenuOpt("1. Create Directory", "Create this directory")
                printMenuOpt("2. Directory Chooser (default)", "Specify a new directory")
                printMenuLine(menuStr02)
                Dim input As String = Console.ReadLine()
                Select Case input
                    Case "1"
                        Directory.CreateDirectory(dir)
                    Case "0"
                        dir = Environment.CurrentDirectory
                        exitCode = True
                        iExitCode = True
                        Exit Sub
                    Case "2", ""
                        dChooser(dir)
                    Case Else
                        menuTopper = invInpStr
                End Select
                If Not Directory.Exists(dir) And Not menuTopper = invInpStr Then menuTopper = "Error"
            End While
        End While
    End Sub

    Public Sub fChooser(ByRef dir As String, ByRef name As String, defaultName As String, defaultRename As String)
        If exitCode Then Exit Sub
        Console.Clear()
        menuTopper = "File Chooser"
        printMenuTop({"Choose a file name, or open the directory chooser to choose a directory"})
        printMenuOpt("1. " & defaultName, "Use the default name")
        If defaultRename <> "" Then printMenuOpt(getMenuNumber({defaultName <> ""}, 1) & ". " & defaultRename, "Use the default rename")
        printMenuOpt(getMenuNumber({defaultName <> "", defaultRename <> ""}, 1) & ". Directory Chooser", "Choose a new directory")
        printBlankMenuLine()
        printMenuLine("Current Directory: " & replDir(dir), "l")
        printMenuLine("Current File:      " & name, "l")
        printMenuLine(menuStr02)
        Console.Write("Enter a number, a new file name, or leave blank to continue using '" & name & "': ")
        Dim input As String = Console.ReadLine
        Select Case input
            Case "0"
                exitCode = True
                Console.Clear()
                Exit Sub
            Case ""
                name = name
            Case "1"
                name = defaultName
            Case "2"
                If defaultRename <> "" Then
                    name = defaultRename
                Else
                    dChooser(dir)
                End If
            Case "3"
                dChooser(dir)
            Case Else
                name = input
        End Select
        If exitCode Then Exit Sub
        Dim iExitCode As Boolean = False
        menuTopper = "File Chooser"
        Do Until iExitCode
            Console.Clear()
            printMenuTop({"Confirm your settings or return to the options to change them."})
            printMenuOpt("1. File Chooser", "Change the file name")
            printMenuOpt("2. Directory Chooser", "Change the directory")
            printMenuOpt("3. Confirm (default)", "Save changes")
            printBlankMenuLine()
            printMenuLine("Current Directory: " & replDir(dir), "l")
            printMenuLine("Current File:      " & name, "l")
            printMenuLine(menuStr02)
            Console.Write("Enter a number, or leave blank to run the default: ")
            input = Console.ReadLine()
            Select Case input
                Case "", "3"
                    iExitCode = True
                Case "0"
                    exitCode = True
                Case "1"
                    fChooser(dir, name, defaultName, defaultRename)
                    iExitCode = True
                Case "2"
                    dChooser(dir)
                    iExitCode = True
                Case Else
                    menuTopper = invInpStr
            End Select
        Loop
    End Sub

    Public Sub dChooser(ByRef dir As String)
        If exitCode Then Exit Sub
        Console.Clear()
        menuTopper = "Directory Chooser"
        printMenuTop({"Choose a directory"})
        printMenuOpt("1. Use default (default)", "Use the same folder as winapp2ool.exe")
        printMenuOpt("2. Parent Folder", "Go up a level")
        printMenuOpt("3. Current folder", "Continue using the same folder as below")
        printBlankMenuLine()
        printMenuLine("Current Directory: " & dir, "l")
        printMenuLine(menuStr02)
        Console.Write("Choose a number from above, enter a new directory, or leave blank to run the default: ")
        Dim uPath As String = Console.ReadLine()
        Select Case uPath
            Case ""
                dir = Environment.CurrentDirectory
            Case "parent"
                dir = Directory.GetParent(dir).ToString
            Case "0"
                exitCode = True
                Exit Sub
            Case Else
                dir = uPath
                Console.Clear()
                chkDirExist(dir)
        End Select
        Console.Clear()
        printMenuLine(tmenu("Directory chooser"))
        printMenuLine(moMenu("Current Directory: "))
        printMenuLine(dir, "l")
        printMenuLine(menuStr03)
        printMenuLine("Options", "c")
        printMenuLine("Enter '1' to change directory", "l")
        printMenuLine("Enter anything else to confirm directory change", "l")
        printMenuLine(menuStr02)
        uPath = Console.ReadLine()
        If uPath.Trim = "1" Then
            dChooser(dir)
        End If
    End Sub

    Public Class iniFile
        Dim lineCount As Integer = 1

        'The current state of the directory & name of the file
        Public dir As String
        Public name As String

        'The inital state of the direcotry & name of the file (for restoration purposes) 
        Public initDir As String
        Public initName As String

        'Suggested rename for output files
        Public secondName As String

        'Sections will be initally stored in the order they're read
        Public sections As New Dictionary(Of String, iniSection)

        'Any line comments will be saved in the order they're read 
        Public comments As New Dictionary(Of Integer, iniComment)

        Public Overrides Function toString() As String
            Dim out As String = ""

            For i As Integer = 0 To sections.Count - 2
                out += sections.Values(i).ToString & Environment.NewLine
            Next
            out += sections.Values.Last.ToString

            Return out
        End Function

        Public Sub New()
            dir = ""
            name = ""
            initDir = ""
            initName = ""
            secondName = ""
        End Sub

        Public Sub New(filename As String)
            dir = Environment.CurrentDirectory
            name = filename
            init()
        End Sub

        Public Sub New(directory As String, filename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = ""
            init()
        End Sub

        Public Sub New(directory As String, filename As String, rename As String)
            dir = directory
            name = filename
            initDir = directory
            initName = filename
            secondName = rename
            init()
        End Sub

        Public Function path() As String
            Return dir & "\" & name
        End Function

        'This is used for constructing ini files sourced from the internet
        Public Sub New(lines As String(), name As String)
            Dim sectionToBeBuilt As New List(Of String)
            Dim lineTrackingList As New List(Of Integer)
            Dim lastLineWasEmpty As Boolean = False
            lineCount = 1
            For Each line In lines
                processiniLine(line, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
            Next
            If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
        End Sub

        Private Sub processiniLine(ByRef currentLine As String, ByRef sectionToBeBuilt As List(Of String), ByRef lineTrackingList As List(Of Integer), ByRef lastLineWasEmpty As Boolean)
            Select Case True
                Case currentLine.StartsWith(";")
                    Dim newCom As New iniComment(currentLine, lineCount)
                    comments.Add(comments.Count, newCom)
                Case Not currentLine.StartsWith("[") And Not currentLine.Trim = ""
                    sectionToBeBuilt.Add(currentLine)
                    lineTrackingList.Add(lineCount)
                    lastLineWasEmpty = False
                Case currentLine.Trim <> ""
                    If Not sectionToBeBuilt.Count = 0 Then
                        mkSection(sectionToBeBuilt, lineTrackingList)
                        sectionToBeBuilt.Add(currentLine)
                        lineTrackingList.Add(lineCount)
                        lastLineWasEmpty = False
                    Else
                        sectionToBeBuilt.Add(currentLine)
                        lineTrackingList.Add(lineCount)
                        lastLineWasEmpty = False
                    End If
                Case Else
                    lastLineWasEmpty = True
            End Select
            lineCount += 1
        End Sub

        Public Sub init()
            Try
                Dim r As StreamReader
                Dim sectionToBeBuilt As New List(Of String)
                Dim lineTrackingList As New List(Of Integer)
                Dim lastLineWasEmpty As Boolean = False
                r = New StreamReader(Me.path())
                Do While (r.Peek() > -1)
                    processiniLine(r.ReadLine.ToString, sectionToBeBuilt, lineTrackingList, lastLineWasEmpty)
                Loop
                If sectionToBeBuilt.Count <> 0 Then mkSection(sectionToBeBuilt, lineTrackingList)
                r.Close()
            Catch ex As Exception
                Console.WriteLine(ex.Message & Environment.NewLine & "Failure occurred during iniFile construction at line: " & lineCount & " in " & name)
            End Try
        End Sub

        'find the line number of a particular comment by its string, return -1 if DNE
        Public Function findCommentLine(com As String) As Integer
            For Each comment In comments.Values
                If comment.comment = com Then Return comment.lineNumber
            Next
            Return -1
        End Function

        Public Function getSectionNamesAsList() As List(Of String)
            Dim out As New List(Of String)
            For Each section In sections.Values
                out.Add(section.name)
            Next
            Return out
        End Function

        Private Sub mkSection(sectionToBeBuilt As List(Of String), lineTrackingList As List(Of Integer))
            Dim sectionHolder As New iniSection(sectionToBeBuilt, lineTrackingList)
            Try
                sections.Add(sectionHolder.name, sectionHolder)
            Catch ex As Exception

                'This will catch entries whose names are identical (case sensitive), but will not catch wholly duplicate FileKeys (etc) 
                If ex.Message = "An item with the same key has already been added." Then

                    Dim lineErr As Integer
                    For Each section In sections.Values
                        If section.name = sectionToBeBuilt(0) Then
                            lineErr = section.startingLineNumber
                            Exit For
                        End If
                    Next

                    Console.WriteLine("Error: Duplicate section name detected: " & sectionToBeBuilt(0))
                    Console.WriteLine("Line: " & lineCount)
                    Console.WriteLine("Duplicates the entry on line: " & lineErr)
                    Console.WriteLine("This section will be ignored until it is given a unique name.")
                    Console.WriteLine()
                End If
            Finally
                sectionToBeBuilt.Clear()
                lineTrackingList.Clear()
            End Try
        End Sub
    End Class

    Public Class iniSection
        Public startingLineNumber As Integer
        Public endingLineNumber As Integer
        Public name As String
        Public keys As New Dictionary(Of Integer, iniKey)

        'expects that the number of parameters in keyTypeList is equal to the number of lists in the list of keylists, with the final list being an "error" list 
        'that holds keys of a type not defined in the keytypelist
        Public Sub constructKeyLists(ByVal keyTypeList As List(Of String), ByRef listOfKeyLists As List(Of List(Of iniKey)))

            For Each key In Me.keys.Values
                If keyTypeList.Contains(key.keyType.ToLower) Then
                    listOfKeyLists(keyTypeList.IndexOf(key.keyType.ToLower)).Add(key)
                Else
                    listOfKeyLists.Last.Add(key)
                End If
            Next
        End Sub

        Public Function getFullName() As String
            Return "[" & Me.name & "]"
        End Function

        Public Sub New()
            startingLineNumber = 0
            endingLineNumber = 0
            name = ""
        End Sub

        Public Sub New(ByVal listOfLines As List(Of String), listOfLineCounts As List(Of Integer))

            Dim tmp1 As String() = listOfLines(0).Split(Convert.ToChar("["))
            Dim tmp2 As String() = tmp1(1).Split(Convert.ToChar("]"))

            name = tmp2(0)

            startingLineNumber = listOfLineCounts(0)
            endingLineNumber = listOfLineCounts(listOfLineCounts.Count - 1)

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    Dim curKey As New iniKey(listOfLines(i), listOfLineCounts(i))
                    keys.Add(i - 1, curKey)
                Next
            End If
        End Sub

        Public Sub New(ByVal listOfLines As List(Of String))

            Dim tmp1 As String() = listOfLines(0).Split(Convert.ToChar("["))
            Dim tmp2 As String() = tmp1(1).Split(Convert.ToChar("]"))

            name = tmp2(0)

            startingLineNumber = 1
            endingLineNumber = 1 + listOfLines.Count

            If listOfLines.Count > 1 Then
                For i As Integer = 1 To listOfLines.Count - 1
                    keys.Add(i - 1, New iniKey(listOfLines(i)))
                Next
            End If
        End Sub

        Public Function getKeysAsList() As List(Of String)

            Dim out As New List(Of String)

            For Each key In Me.keys.Values
                out.Add(key.toString)
            Next

            Return out
        End Function

        'returns true if the sections are the same, else returns false
        Public Function compareTo(ss As iniSection, ByRef removedKeys As List(Of iniKey), ByRef addedKeys As List(Of iniKey), ByRef updatedKeys As List(Of KeyValuePair(Of iniKey, iniKey))) As Boolean

            'Create a copy of the section so we can modify it
            Dim secondSection As New iniSection
            secondSection.name = ss.name
            secondSection.startingLineNumber = ss.startingLineNumber
            For i As Integer = 0 To ss.keys.Count - 1
                secondSection.keys.Add(i, ss.keys.Values(i))
            Next

            Dim noMatch As Boolean
            Dim tmpList As New List(Of Integer)

            For Each key In keys.Values
                noMatch = True
                For i As Integer = 0 To secondSection.keys.Values.Count - 1
                    If key.value = secondSection.keys.Values(i).value Then
                        noMatch = False
                        tmpList.Add(i)
                        Exit For
                    End If
                Next
                If noMatch Then
                    'If the key isn't found in the second (newer) section, consider it removed for now
                    removedKeys.Add(key)
                End If
            Next

            'Remove all matched keys
            tmpList.Reverse()
            For Each ind In tmpList
                secondSection.keys.Remove(ind)
            Next

            'Assume any remaining keys have been added
            For Each key In secondSection.keys.Values
                addedKeys.Add(key)
            Next

            'Check for keys whose names match
            Dim rkTemp, akTemp As New List(Of iniKey)
            rkTemp = removedKeys.ToList
            akTemp = addedKeys.ToList
            For Each key In removedKeys
                For Each skey In addedKeys
                    If key.name = skey.name Then

                        Dim oldKey As New winapp2KeyParameters(key)
                        Dim newKey As New winapp2KeyParameters(skey)
                        oldKey.argsList.Sort()
                        newKey.argsList.Sort()

                        If oldKey.argsList.Count = newKey.argsList.Count Then

                            For i As Integer = 0 To oldKey.argsList.Count - 1
                                If Not oldKey.argsList(i) = newKey.argsList(i) Then
                                    updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
                                    rkTemp.Remove(key)
                                    akTemp.Remove(skey)
                                    Exit For
                                End If
                            Next
                            rkTemp.Remove(key)
                            akTemp.Remove(skey)
                        Else
                            updatedKeys.Add(New KeyValuePair(Of iniKey, iniKey)(key, skey))
                            rkTemp.Remove(key)
                            akTemp.Remove(skey)
                        End If
                    End If
                Next
            Next

            'Update the lists
            addedKeys = akTemp
            removedKeys = rkTemp

            Return removedKeys.Count + addedKeys.Count + updatedKeys.Count = 0

        End Function

        Public Overrides Function ToString() As String

            Dim out As String = Me.getFullName

            For Each key In keys.Values
                out += Environment.NewLine & key.toString
            Next
            out += Environment.NewLine
            Return out

        End Function
    End Class

    Public Class iniKey

        Public name As String
        Public value As String
        Public lineNumber As Integer
        Public keyType As String

        'Create an empty key
        Public Sub New()
            name = ""
            value = ""
            lineNumber = 0
            keyType = ""
        End Sub

        'Strip any numbers from the name value in a key (so numbered keys can be identified by "type")
        Private Function stripNums(keyName As String) As String
            Return New Regex("[\d]").Replace(keyName, "")
        End Function

        'Create a key with a line string from a file and line number counter 
        Public Sub New(ByVal line As String, ByVal count As Integer)

            Try
                'valid keys have the format name=value
                Dim splitLine As String() = line.Split(CChar("="))
                name = splitLine(0)
                value = splitLine(1)
                keyType = stripNums(name)
                lineNumber = count
            Catch ex As Exception
                exc(ex)
            End Try
        End Sub

        'for when trackling line numbers doesn't matter
        Public Sub New(ByVal line As String)
            Dim splitLine As String() = line.Split(CChar("="))
            name = splitLine(0)
            value = splitLine(1)
            keyType = stripNums(name)
        End Sub

        Public Overrides Function toString() As String
            'Output the key in name=value format
            Return Me.name & "=" & Me.value
        End Function

        'Output the key in Line: <line number> - name=value format
        Public Function lineString() As String
            Return "Line: " & lineNumber & " - " & name & "=" & value
        End Function

        'compare the name=value format of two different keys, return their equivalency as boolean
        Public Function compareTo(secondKey As iniKey) As Boolean
            If secondKey IsNot Nothing Then Return Me.toString.Equals(secondKey.toString)
            Return False
        End Function
    End Class

    'small wrapper class for capturing ini comment data
    Class iniComment

        Public comment As String
        Public lineNumber As Integer

        Public Sub New(c As String, l As Integer)
            comment = c
            lineNumber = l
        End Sub
    End Class
End Module