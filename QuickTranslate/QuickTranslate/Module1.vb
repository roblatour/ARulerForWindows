Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text

Module Module1

    Const gMaxNumberOfOtherLangages As Integer = 10

    Private gBaseDirectory As String = String.Empty

    Private gInputLanguagesFileName As String = String.Empty

    Private gInputResourceFolder As String = String.Empty

    Private gOutputResourceFolderForForms As String = String.Empty

    Private gOutputResourceFolderForInCodeLiterals As String = String.Empty

    Private gResGen_exe As String = String.Empty

    Private gal_exe As String = String.Empty

    ' gFormTranslationTable:
    '       gTranslationTable(x, 0) english description of language
    '       gTranslationTable(x, 1) language code
    '       gTranslationTable(x, n) rows to be translated

    Private gTranslationTable(gMaxNumberOfOtherLangages, 50000) As String

    Sub Main()

        SetDirectoryStructure()
        ProcessForms()
        ProcessInCode()

    End Sub

    Private Sub SetDirectoryStructure()

        gBaseDirectory = Environment.CurrentDirectory

        ' Find Localized Working Files

        Dim HuntForLocalizationWorkingFilesDirectory As String = gBaseDirectory

        While Not HuntForLocalizationWorkingFilesDirectory.Contains("Localization Working Files")

            If Directory.Exists(HuntForLocalizationWorkingFilesDirectory & "\Localization Working Files") Then
                HuntForLocalizationWorkingFilesDirectory = HuntForLocalizationWorkingFilesDirectory & "\Localization Working Files"
                Exit While
            Else
                Dim parentDir As DirectoryInfo = Directory.GetParent(HuntForLocalizationWorkingFilesDirectory)
                HuntForLocalizationWorkingFilesDirectory = parentDir.ToString
            End If

        End While

        gInputLanguagesFileName = HuntForLocalizationWorkingFilesDirectory & "\Translation.xls"

        gResGen_exe = HuntForLocalizationWorkingFilesDirectory & "\Misc\Resgen.exe"

        gal_exe = HuntForLocalizationWorkingFilesDirectory & "\Misc\al"

        ' Find eruler / eruler direcotry

        Dim HuntErulerSlashErulerDirectory As String = gBaseDirectory

        While Not HuntErulerSlashErulerDirectory.Contains("\eruler\eruler")

            If Directory.Exists(HuntErulerSlashErulerDirectory & "\eruler\eruler") Then
                HuntErulerSlashErulerDirectory = HuntErulerSlashErulerDirectory & "\eruler\eruler"
                Exit While
            Else
                Dim parentDir As DirectoryInfo = Directory.GetParent(HuntErulerSlashErulerDirectory)
                HuntErulerSlashErulerDirectory = parentDir.ToString
            End If

        End While

        gInputResourceFolder = HuntErulerSlashErulerDirectory

        gOutputResourceFolderForForms = HuntErulerSlashErulerDirectory

        gOutputResourceFolderForInCodeLiterals = HuntErulerSlashErulerDirectory & "\Resources\"


    End Sub


    Private Sub ProcessForms()

        LoadLanguages(gInputLanguagesFileName, "InForms$")

        CreateTranslatedResourceFilesForAllForms()

    End Sub

    Private Sub ProcessInCode()

        ReDim gTranslationTable(gMaxNumberOfOtherLangages, 10000)

        LoadLanguages(gInputLanguagesFileName, "InCode$")

        CreateTranslatedResourceFilesForInCodeLiterals()

    End Sub

    Private Sub CreateTranslatedResourceFilesForAllForms()

        Dim InputDefaultResourceFiles As List(Of String) = LoadInputResourcesFiles(gInputResourceFolder)

        'For each default resouce file create and equivalent resource file for each other other language
        For Each InputResourceFileName In InputDefaultResourceFiles
            CreateTranslatedResourceFile(InputResourceFileName)
        Next

    End Sub

    Private Sub CreateTranslatedResourceFile(ByVal InputResourceFileName As String)

        Dim OutputResourceFileName As String
        Dim LanguageCode As String

        Dim MaxSpreadsheetColumns As Integer = gTranslationTable.GetUpperBound(0)
        Dim MaxSpreadsheetRows As Integer = gTranslationTable.GetUpperBound(1)

        Dim InputResourseFileContents As String = My.Computer.FileSystem.ReadAllText(InputResourceFileName)
        Dim OutputResourseFileContents As String

        Try

            For SpreadsheetColumn As Integer = 1 To MaxSpreadsheetColumns

                OutputResourseFileContents = InputResourseFileContents

                LanguageCode = gTranslationTable(SpreadsheetColumn, 1)

                OutputResourceFileName = gOutputResourceFolderForForms & System.IO.Path.GetFileName(InputResourceFileName).Replace(".resx", "." & LanguageCode & ".resx")
                Console.WriteLine(OutputResourceFileName)

                For CurrentRow = 2 To MaxSpreadsheetRows

                    If OutputResourseFileContents.Contains("<value>" & gTranslationTable(0, CurrentRow) & "</value>") Then

                        OutputResourseFileContents = OutputResourseFileContents.Replace("<value>" & gTranslationTable(0, CurrentRow) & "</value>", "<value>" & gTranslationTable(SpreadsheetColumn, CurrentRow) & "</value>")

                    End If

                Next

                My.Computer.FileSystem.WriteAllText(OutputResourceFileName, OutputResourseFileContents, False)

            Next

        Catch ex As Exception

            MsgBox("CreateTranslatedResourceFile" & vbCrLf & ex.Message)
            Environment.Exit(2)

        End Try

    End Sub

    Private Sub CreateTranslatedResourceFilesForInCodeLiterals()

        Const FileName As String = "LanguageFile"
        Const FileExtention As String = ".txt"

        Dim Path_FileName As String = String.Empty
        Dim Path_FileName_FileExtention As String = String.Empty

        Dim LanguageCode As String = String.Empty

        Dim OutputFileContents As New StringBuilder

        For ColumnCount As Integer = 0 To gTranslationTable.GetUpperBound(0)

            LanguageCode = gTranslationTable(ColumnCount, 1)

            '
            OutputFileContents.Clear()  ' for > dot net 3.5
            'OutputFileContents.Length = 0  ' for <= dot net 3.5

            For RowCount As Integer = 0 To gTranslationTable.GetUpperBound(1)

                OutputFileContents.Append(Format(RowCount + 1, "0000") & "=" & gTranslationTable(ColumnCount, RowCount) & vbCrLf)

            Next

            If ColumnCount = 0 Then
                Path_FileName = gOutputResourceFolderForInCodeLiterals & FileName
            Else
                Path_FileName = gOutputResourceFolderForInCodeLiterals & FileName & "." & LanguageCode
            End If

            Path_FileName_FileExtention = Path_FileName & FileExtention

            Console.WriteLine(Path_FileName_FileExtention)
            My.Computer.FileSystem.WriteAllText(Path_FileName_FileExtention, OutputFileContents.ToString, False)

            MakeAResourceFile(Path_FileName)

            MakeADLL(Path_FileName, LanguageCode)

        Next

    End Sub

    Private Sub MakeAResourceFile(ByVal Path_FileName As String)

        Try

            Dim Text_FileName As String = Path_FileName + ".txt"
            Dim Resource_FileName As String = Path_FileName + ".resources"

            If File.Exists(Resource_FileName) Then File.Delete(Resource_FileName)

            Dim myProcess As New Process
            myProcess.StartInfo.Verb = "open"
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = gResGen_exe
            myProcess.StartInfo.Arguments = """" & Text_FileName & """ """ & Resource_FileName & """"
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            myProcess.WaitForExit()

            File.Delete(Text_FileName)

        Catch ex As Exception
        End Try

    End Sub

    Private Sub MakeADLL(ByVal Path_FileName As String, ByVal LanguageCode As String)

        Try

            Dim Resource_FileName As String = Path_FileName + ".resources"
            Dim Dll_FileName As String = Path_FileName + ".dll"

            If File.Exists(Dll_FileName) Then File.Delete(Dll_FileName)

            Dim Arguments As String = String.Empty

            If LanguageCode = "en-US" Then
                Arguments = "/t:lib /embed:LanguageFile.%1.resources  /culture:%1 /out:aruler.LanguageFile.dll".Replace("%1", LanguageCode)
            Else
                Arguments = "/t:lib /embed:LanguageFile.%1.resources  /culture:%1 /out:aruler.LanguageFile.%1.dll".Replace("%1", LanguageCode)
            End If

            Dim myProcess As New Process
            myProcess.StartInfo.Verb = "open"
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = gal_exe
            myProcess.StartInfo.Arguments = Arguments
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.StartInfo.WorkingDirectory = gOutputResourceFolderForInCodeLiterals
            myProcess.Start()
            myProcess.WaitForExit()

        Catch ex As Exception
        End Try

    End Sub

    Private Sub LoadLanguages(ByVal filename As String, ByVal WorkSheetOfInterest As String)

        If Not String.IsNullOrEmpty(filename) Then

            Try

                Dim OExcelHandler As New ExcelHandler()
                Dim ds As DataSet = OExcelHandler.GetDataFromExcel(filename)

                Dim WorksheetName As String

                If ds IsNot Nothing Then

                    For Each Table As DataTable In ds.Tables

                        WorksheetName = Table.TableName.ToString

                        Dim xx As Integer = 0
                        For Each WorksheetColumnName As DataColumn In Table.Columns
                            If xx <= gMaxNumberOfOtherLangages Then
                                gTranslationTable(xx, 0) = WorksheetColumnName.ColumnName
                                xx += 1
                            End If

                        Next

                    Next

                End If

                Dim WorkSheetColumnNameOfInterest As String
                Dim RowCount As Integer = 0

                Dim ClientsWorksheet As DataTable = ds.Tables(WorkSheetOfInterest)

                For ColumnCount As Integer = 0 To gTranslationTable.GetUpperBound(0)

                    WorkSheetColumnNameOfInterest = gTranslationTable(ColumnCount, 0)

                    RowCount = 1

                    For Each WorksheetRow As DataRow In ClientsWorksheet.Rows

                        If TypeOf (WorksheetRow(WorkSheetColumnNameOfInterest)) Is DBNull Then
                        Else
                            gTranslationTable(ColumnCount, RowCount) = WorksheetRow(WorkSheetColumnNameOfInterest)
                        End If
                        RowCount += 1

                    Next

                Next

                ReDim Preserve gTranslationTable(gTranslationTable.GetUpperBound(0), RowCount - 1)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

    End Sub

    Private Function LoadInputResourcesFiles(ByVal InputResourceFolder As String) As List(Of String)

        'Build a collection of all the resource files in the program directory

        Dim AllInputResourceFiles As ReadOnlyCollection(Of String)
        AllInputResourceFiles = My.Computer.FileSystem.GetFiles(InputResourceFolder, FileIO.SearchOption.SearchAllSubDirectories, "*.resx")

        'Create a list of all the default resource files 
        Dim InputDefaultResourceFiles = New List(Of String)

        For Each FileName In AllInputResourceFiles

            If FileName.ToUpper.Contains("\FRM") Then
                If CountCharactersInAString(".", FileName) = 1 Then
                    InputDefaultResourceFiles.Add(FileName)
                End If
            End If

        Next

        'return only the list of default resource files
        Return InputDefaultResourceFiles

    End Function

    Private Function CountCharactersInAString(ByVal Character As String, ByVal InputString As String) As Integer

        Return InputString.Split({Character}, StringSplitOptions.None).Length - 1

    End Function

End Module
