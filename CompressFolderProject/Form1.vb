Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Form1
    Dim directoryTargetLocation As String
    Dim Destinydirectory As String
    Dim exePath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "/7zFile/7z.exe"
    Dim timeList As List(Of TimeSpan) = New List(Of TimeSpan)
    Dim dictOS As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    Dim dictNewOSdetails As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    Dim dictLogVersion As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    Dim dictTopProg As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    Dim countOfOS As Integer, countOflogVersion As Integer = 0, countOfNewOS As Integer = 0, countOfTopProg As Integer = 0
    Dim messageTopProg As String = "", messageLogVersion As String = "", messageOSTransferDetails As String = "",
         messageCountOfNewOS As String = ""
    Function extract7z(zipFileFolder As String, ToFolder As String)
        Dim args As String = "e " + zipFileFolder + " -o" + ToFolder + "" + " -p""cyberspa123""" + " -aoa"
        Dim p As New Process
        Dim pInfo As New ProcessStartInfo
        pInfo.FileName = exePath
        pInfo.Arguments = args
        pInfo.WindowStyle = ProcessWindowStyle.Hidden
        p.StartInfo = pInfo
        p.Start()
        p.WaitForExit()
        '   System.Diagnostics.Process.Start(exePath, args)
        'Threading.Thread.Sleep(1000)
        ' System.IO.File.Delete(zipFileFolder)
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(ToFolder)
            Dim check As String = System.IO.Path.GetExtension(foundFile)
            If (check = ".7z") Then
                Dim zipFolderpath1 As String = System.IO.Path.GetFullPath(ToFolder & "/" & System.IO.Path.GetFileNameWithoutExtension(foundFile))
                extract7z(foundFile, zipFolderpath1)
            End If
        Next
    End Function
    Function showMatch(ByVal fileName As FileInfo, ByVal text As String)
        Dim regex As Regex = New Regex("(\d+)")
        Dim match As Match = regex.Match(text)
        Dim str1 As String = StrReverse(text)
        Dim type As String = str1.Substring(0, 2)
        If match.Success Then
            Dim sizeOfDisk As Decimal = Convert.ToDecimal(match.Value)
            If (type = "BK") Then
                sizeOfDisk = Format((sizeOfDisk / 1024) / 1024, "0.00")
            ElseIf (type = "BM") Then
                sizeOfDisk = Format((sizeOfDisk / 1024), "0.00")
            ElseIf (type = "BT") Then
                sizeOfDisk = (sizeOfDisk * 1024)
            End If
            DataGridView1.Rows.Add("File Name: " & fileName.ToString, sizeOfDisk)
        End If
    End Function

    Function addinList(ByVal StrFiles As String)

        Dim serviceTime As String = StrFiles.Substring(20, 16)
        Dim regex As Regex = New Regex("(\d+)")

        Dim hour As String = serviceTime.Substring(0, 1)
        Dim matchHour As Match = regex.Match(hour)
        Dim min As String = serviceTime.Substring(4, 2)
        Dim matchMin As Match = regex.Match(min)
        Dim sec As String = serviceTime.Substring(10, 2)
        Dim matchSec As Match = regex.Match(min)
        timeList.Add(New TimeSpan(matchHour.Value, matchMin.Value, matchSec.Value))
        Return timeList
    End Function
    Function Caltime(ByVal timeList As List(Of TimeSpan))
        Dim totalTime As TimeSpan
        Dim TimeEntry As TimeSpan
        For i = 0 To timeList.Count - 1
            TimeEntry = timeList.ElementAt(i)
            totalTime += TimeEntry
        Next
        Dim Secs As Decimal = totalTime.TotalSeconds
        Secs = Secs / timeList.Count
        Dim AvgTime As TimeSpan = TimeSpan.FromSeconds(Secs)
        TextBox7.Text = String.Format("{0:00}:{1:00}:{2:00}", AvgTime.Hours, AvgTime.Minutes, AvgTime.Seconds)
        timeList.Sort(New Comparison(Of TimeSpan)(Function(x As TimeSpan, y As TimeSpan) x.CompareTo(y)))
        TextBox8.Text = timeList.ElementAt(0).ToString
        TextBox9.Text = timeList.ElementAt(timeList.Count - 1).ToString
    End Function
    Function addtoMapCountOfOS(ByVal oldDriveToNewDrive As String)
        If (dictOS.TryGetValue(oldDriveToNewDrive, countOfOS)) Then
            dictOS(oldDriveToNewDrive) += 1
        Else
            dictOS.Add(oldDriveToNewDrive, 1)
        End If
        Return dictOS
    End Function
    Function addtoMapCountOfVersion(ByVal versionString As String)
        If (dictLogVersion.TryGetValue(versionString, countOflogVersion)) Then
            dictLogVersion(versionString) += 1
        Else
            dictLogVersion.Add(versionString, 1)
        End If
        Return dictLogVersion
    End Function
    Function addtoMapCountOfNewOS(ByVal countOfewOS As String)
        If (dictNewOSdetails.TryGetValue(countOfewOS, countOfNewOS)) Then
            dictNewOSdetails(countOfewOS) += 1
        Else
            dictNewOSdetails.Add(countOfewOS, 1)
        End If
        Return dictNewOSdetails
    End Function
    Function addtoMapCountTopProg(ByVal topProgSName As String)
        If (dictTopProg.TryGetValue(topProgSName, countOfTopProg)) Then
            dictTopProg(topProgSName) += 1
        Else
            dictTopProg.Add(topProgSName, 1)
        End If
        Return dictTopProg
    End Function
    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Try
    '        ZipFile.CreateFromDirectory(TextBox1.Text, TextBox2.Text)
    '        MessageBox.Show("Compress Succesfully!")
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Threading.Thread.Sleep(1000)
            System.IO.Directory.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "/7zFile", True)
            'System.IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "/ProgList.jar")
            End
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            FolderBrowserDialog1.Description = "Select directory"
            With FolderBrowserDialog1
                If .ShowDialog() = DialogResult.OK Then
                    directoryTargetLocation = .SelectedPath
                    TextBox1.Text = directoryTargetLocation.ToString
                End If
            End With
            FolderBrowserDialog1.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Try
        '    Button4.Enabled = False
        '    If (Not System.IO.Directory.Exists(TextBox2.Text)) Then
        '        System.IO.Directory.CreateDirectory(TextBox2.Text)
        '    End If
        '    For Each foundFile1 As String In My.Computer.FileSystem.GetFiles(TextBox1.Text)
        '        Dim zipFolderpath As String = System.IO.Path.GetFullPath(TextBox2.Text & "/" & System.IO.Path.GetFileNameWithoutExtension(foundFile1))
        '        extract7z(foundFile1, zipFolderpath)
        '    Next


        '    'MessageBox.Show("Unzip Successfully!")

        '    DataGridView1.Rows.Clear()
        '    DataGridView2.Rows.Clear()
        '    Dim filesSize(), filesProg(), fileLogsDateRange(), fileServiceTime() As System.IO.FileInfo
        '    Dim i As Integer = 0
        '    Dim zipinsideFolderDir, zipinsideFolderDir1 As String
        '    Dim avgSize As Decimal
        '    Dim avgProg, counterSizeFile, counterProgFile As Integer
        '    counterSizeFile = 0
        '    counterProgFile = 0
        '    Dim dateList As List(Of DateTime) = New List(Of DateTime)
        '    avgSize = 0
        '    Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)

        '    filesSize = dirInfo.GetFiles("nonFatalErrors.txt", IO.SearchOption.AllDirectories)
        '    TextBox4.Text = "Total Files are: " & filesSize.Count.ToString
        '    For Each file In filesSize
        '        zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
        '        Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
        '        While objreader.Peek <> -1
        '            Dim StrFiles As String = objreader.ReadLine()
        '            If (StrFiles.Contains("Total Data Size: ")) Then
        '                counterSizeFile = counterSizeFile + 1
        '                Dim sizeOfDiskInStr As String = StrFiles.Substring(17)

        '                showMatch(file, sizeOfDiskInStr)
        '            End If
        '            StrFiles = String.Empty
        '        End While
        '        objreader.Close()
        '    Next
        '    filesProg = dirInfo.GetFiles("olddriveinstalledPrograms.txt", IO.SearchOption.AllDirectories)
        '    For Each file In filesProg
        '        zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
        '        Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
        '        While objreader1.Peek <> -1
        '            Dim StrFiles As String = objreader1.ReadLine()
        '            If (StrFiles.Contains("Programs found installed on the old drive")) Then
        '                counterProgFile = counterProgFile + 1
        '                Dim numberOfProg As String = StrFiles.Substring(0, 2)
        '                Dim noOfProgInt As Integer = Convert.ToInt32(numberOfProg)
        '                DataGridView2.Rows.Add()
        '                DataGridView2.Item(0, i).Value = noOfProgInt
        '                i = i + 1
        '            End If
        '            StrFiles = String.Empty
        '        End While
        '        objreader1.Close()
        '    Next

        '    fileLogsDateRange = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
        '    For Each file In fileLogsDateRange
        '        zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
        '        Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
        '        While objreader1.Peek <> -1
        '            Dim StrFiles As String = objreader1.ReadLine()
        '            If (StrFiles.Contains("Freshstart app launched at:")) Then
        '                Dim logDate As String = StrFiles.Substring(28, 10)
        '                dateList.Add(DateTime.ParseExact(logDate, "dd_MM_yyyy", CultureInfo.InvariantCulture))
        '            ElseIf (StrFiles.Contains("Next button clicked at")) Then
        '                Dim logDate As String = StrFiles.Substring(23, 10)
        '                dateList.Add(DateTime.ParseExact(logDate, "dd_MM_yyyy", CultureInfo.InvariantCulture))
        '            End If
        '            If (StrFiles.Contains("")) Then

        '            End If

        '            StrFiles = String.Empty
        '        End While
        '        objreader1.Close()
        '    Next
        '    dateList.Sort(New Comparison(Of Date)(Function(x As Date, y As Date) y.CompareTo(x)))
        '    Dim d1 As Date = dateList.ElementAt(dateList.Count - 1)
        '    Dim s1 As String = d1.Month & "_" & d1.Day & "_" & d1.Year
        '    Dim d2 As Date = dateList.ElementAt(0)
        '    Dim s2 As String = d2.Month & "_" & d2.Day & "_" & d2.Year
        '    TextBox6.Text = s1 & " to " & s2

        '    fileServiceTime = dirInfo.GetFiles("nonFatalErrors.txt", IO.SearchOption.AllDirectories)
        '    For Each file In fileServiceTime
        '        zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
        '        Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
        '        While objreader1.Peek <> -1
        '            Dim StrFiles As String = objreader1.ReadLine()
        '            If (StrFiles.Contains("Total service time:")) Then
        '                timeList = addinList(StrFiles)
        '            End If

        '            StrFiles = String.Empty
        '        End While
        '        objreader1.Close()
        '    Next

        '    Caltime(timeList)

        '    For Each row As DataGridViewRow In DataGridView1.Rows
        '        If Not row.IsNewRow Then
        '            Dim AvgsizeOfDisk As Decimal = row.Cells(1).Value.ToString()
        '            avgSize = avgSize + AvgsizeOfDisk
        '        End If
        '    Next
        '    TextBox10.Text = Format((avgSize), "0.00") & "GB"
        '    TextBox3.Text = Format((avgSize / counterSizeFile), "0.00") & "GB"
        '    For Each row1 As DataGridViewRow In DataGridView2.Rows
        '        If Not row1.IsNewRow Then
        '            Dim avgProg1 As Integer = row1.Cells(0).Value.ToString()
        '            avgProg = avgProg + avgProg1
        '        End If
        '    Next
        '    TextBox5.Text = Format((avgProg / counterProgFile), "0.00")
        '    Button4.Enabled = True
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message.ToString)
        'End Try
        ProgressBar1.Visible = True
        Timer1.Start()



    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        'DataGridView1.Rows.Clear()
        'DataGridView2.Rows.Clear()
        'Dim filesSize(), filesProg(), fileLogsDateRange(), fileServiceTime() As System.IO.FileInfo
        'Dim i As Integer = 0
        'Dim zipinsideFolderDir, zipinsideFolderDir1 As String
        'Dim avgSize As Decimal
        'Dim avgProg, counterSizeFile, counterProgFile As Integer
        'counterSizeFile = 0
        'counterProgFile = 0
        'Dim dateList As List(Of DateTime) = New List(Of DateTime)


        'avgSize = 0
        'Try
        '    FolderBrowserDialog1.Description = "Select directory"
        '    With FolderBrowserDialog1
        '        If .ShowDialog() = DialogResult.OK Then
        '            directoryTargetLocation = .SelectedPath
        '            Dim dirInfo As New System.IO.DirectoryInfo(directoryTargetLocation)

        '            filesSize = dirInfo.GetFiles("nonFatalErrors.txt", IO.SearchOption.AllDirectories)
        '            TextBox4.Text = "Total Files are: " & filesSize.Count.ToString
        '            For Each file In filesSize
        '                zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
        '                Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
        '                While objreader.Peek <> -1
        '                    Dim StrFiles As String = objreader.ReadLine()
        '                    If (StrFiles.Contains("Total Data Size: ")) Then
        '                        counterSizeFile = counterSizeFile + 1
        '                        Dim sizeOfDiskInStr As String = StrFiles.Substring(17)

        '                        showMatch(file, sizeOfDiskInStr)
        '                    End If
        '                    StrFiles = String.Empty
        '                End While
        '                objreader.Close()
        '            Next


        '            filesProg = dirInfo.GetFiles("olddriveinstalledPrograms.txt", IO.SearchOption.AllDirectories)
        '            For Each file In filesProg
        '                zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
        '                Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
        '                While objreader1.Peek <> -1
        '                    Dim StrFiles As String = objreader1.ReadLine()
        '                    If (StrFiles.Contains("Programs found installed on the old drive")) Then
        '                        counterProgFile = counterProgFile + 1
        '                        Dim numberOfProg As String = StrFiles.Substring(0, 2)
        '                        Dim noOfProgInt As Integer = Convert.ToInt32(numberOfProg)
        '                        DataGridView2.Rows.Add()
        '                        DataGridView2.Item(0, i).Value = noOfProgInt
        '                        i = i + 1
        '                    End If
        '                    StrFiles = String.Empty
        '                End While
        '                objreader1.Close()
        '            Next

        '            fileLogsDateRange = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
        '            For Each file In fileLogsDateRange
        '                zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
        '                Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
        '                While objreader1.Peek <> -1
        '                    Dim StrFiles As String = objreader1.ReadLine()
        '                    If (StrFiles.Contains("Freshstart app launched at:")) Then
        '                        Dim logDate As String = StrFiles.Substring(28, 10)
        '                        dateList.Add(DateTime.ParseExact(logDate, "dd_MM_yyyy", CultureInfo.InvariantCulture))
        '                    ElseIf (StrFiles.Contains("Next button clicked at")) Then
        '                        Dim logDate As String = StrFiles.Substring(23, 10)
        '                        dateList.Add(DateTime.ParseExact(logDate, "dd_MM_yyyy", CultureInfo.InvariantCulture))
        '                    End If
        '                    If (StrFiles.Contains("")) Then

        '                    End If

        '                    StrFiles = String.Empty
        '                End While
        '                objreader1.Close()
        '            Next
        '            dateList.Sort(New Comparison(Of Date)(Function(x As Date, y As Date) y.CompareTo(x)))
        '            Dim d1 As Date = dateList.ElementAt(dateList.Count - 1)
        '            Dim s1 As String = d1.Month & "_" & d1.Day & "_" & d1.Year
        '            Dim d2 As Date = dateList.ElementAt(0)
        '            Dim s2 As String = d2.Month & "_" & d2.Day & "_" & d2.Year
        '            TextBox6.Text = s1 & " to " & s2

        '            fileServiceTime = dirInfo.GetFiles("nonFatalErrors.txt", IO.SearchOption.AllDirectories)
        '            For Each file In fileServiceTime
        '                zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
        '                Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
        '                While objreader1.Peek <> -1
        '                    Dim StrFiles As String = objreader1.ReadLine()
        '                    If (StrFiles.Contains("Total service time:")) Then
        '                        timeList = addinList(StrFiles)
        '                    End If

        '                    StrFiles = String.Empty
        '                End While
        '                objreader1.Close()
        '            Next
        '        End If
        '    End With
        '    Caltime(timeList)

        '    For Each row As DataGridViewRow In DataGridView1.Rows
        '        If Not row.IsNewRow Then
        '            Dim AvgsizeOfDisk As Decimal = row.Cells(1).Value.ToString()
        '            avgSize = avgSize + AvgsizeOfDisk
        '        End If
        '    Next
        '    TextBox10.Text = Format((avgSize), "0.00") & "GB"
        '    TextBox3.Text = Format((avgSize / counterSizeFile), "0.00") & "GB"
        '    For Each row1 As DataGridViewRow In DataGridView2.Rows
        '        If Not row1.IsNewRow Then
        '            Dim avgProg1 As Integer = row1.Cells(0).Value.ToString()
        '            avgProg = avgProg + avgProg1
        '        End If
        '    Next
        '    TextBox5.Text = Format((avgProg / counterProgFile), "0.00")
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message.ToString)
        'End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ProgressBar1.Visible = False
        If Not Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\7zFile") Then
            File.WriteAllBytes(My.Computer.FileSystem.SpecialDirectories.Temp & "/_7zFile.zip", My.Resources._7zFile)
            ZipFile.ExtractToDirectory(My.Computer.FileSystem.SpecialDirectories.Temp & "/_7zFile.zip", My.Computer.FileSystem.SpecialDirectories.Temp)

            System.IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "/_7zFile.zip")
        End If
        TextBox2.Text = "C:\Users\Owner\Desktop\LogHistory"
        '  File.WriteAllBytes(My.Computer.FileSystem.SpecialDirectories.Temp & "/ProgList.jar", My.Resources.ProgList)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar1.Increment(10)
        If ProgressBar1.Value = 100 Then
            Timer1.Stop()
            Try
                Button4.Enabled = False
                If (Not System.IO.Directory.Exists(TextBox2.Text)) Then
                    System.IO.Directory.CreateDirectory(TextBox2.Text)
                End If
                For Each foundFile1 As String In My.Computer.FileSystem.GetFiles(TextBox1.Text)
                    Dim zipFolderpath As String = System.IO.Path.GetFullPath(TextBox2.Text & "/" & System.IO.Path.GetFileNameWithoutExtension(foundFile1))
                    extract7z(foundFile1, zipFolderpath)
                Next
                'MessageBox.Show("Unzip Successfully!")
                DataGridView1.Rows.Clear()
                DataGridView2.Rows.Clear()
                Dim filesSize(), filesProg(), fileLogsDateRange(), fileServiceTime() As System.IO.FileInfo
                Dim i As Integer = 0
                Dim zipinsideFolderDir, zipinsideFolderDir1 As String
                Dim avgSize As Decimal
                Dim avgProg, counterSizeFile, counterProgFile As Integer
                counterSizeFile = 0
                counterProgFile = 0
                Dim dateList As List(Of DateTime) = New List(Of DateTime)
                avgSize = 0
                Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)

                filesSize = dirInfo.GetFiles("nonFatalErrors.txt", IO.SearchOption.AllDirectories)
                TextBox4.Text = "Total Files are: " & filesSize.Count.ToString
                For Each file In filesSize
                    zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
                    Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
                    While objreader.Peek <> -1
                        Dim StrFiles As String = objreader.ReadLine()
                        If (StrFiles.Contains("Total Data Size: ")) Then
                            counterSizeFile = counterSizeFile + 1
                            Dim sizeOfDiskInStr As String = StrFiles.Substring(17)

                            showMatch(file, sizeOfDiskInStr)
                        End If
                        StrFiles = String.Empty
                    End While
                    objreader.Close()
                Next
                filesProg = dirInfo.GetFiles("olddriveinstalledPrograms.txt", IO.SearchOption.AllDirectories)
                For Each file In filesProg
                    zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
                    Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
                    While objreader1.Peek <> -1
                        Dim StrFiles As String = objreader1.ReadLine()
                        If (StrFiles.Contains("Programs found installed on the old drive")) Then
                            counterProgFile = counterProgFile + 1
                            Dim numberOfProg As String = StrFiles.Substring(0, 2)
                            Dim noOfProgInt As Integer = Convert.ToInt32(numberOfProg)
                            DataGridView2.Rows.Add()
                            DataGridView2.Item(0, i).Value = noOfProgInt
                            i = i + 1
                        End If
                        StrFiles = String.Empty
                    End While
                    objreader1.Close()
                Next

                fileLogsDateRange = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
                For Each file In fileLogsDateRange
                    zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
                    Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
                    While objreader1.Peek <> -1
                        Dim StrFiles As String = objreader1.ReadLine()
                        If (StrFiles.Contains("Freshstart app launched at:")) Then
                            Dim logDate As String = StrFiles.Substring(28, 10)
                            dateList.Add(DateTime.ParseExact(logDate, "dd_MM_yyyy", CultureInfo.InvariantCulture))
                        ElseIf (StrFiles.Contains("Next button clicked at")) Then
                            Dim logDate As String = StrFiles.Substring(23, 10)
                            dateList.Add(DateTime.ParseExact(logDate, "dd_MM_yyyy", CultureInfo.InvariantCulture))
                        End If
                        If (StrFiles.Contains("")) Then

                        End If

                        StrFiles = String.Empty
                    End While
                    objreader1.Close()
                Next
                dateList.Sort(New Comparison(Of Date)(Function(x As Date, y As Date) y.CompareTo(x)))
                Dim d1 As Date = dateList.ElementAt(dateList.Count - 1)
                Dim s1 As String = d1.Month & "_" & d1.Day & "_" & d1.Year
                Dim d2 As Date = dateList.ElementAt(0)
                Dim s2 As String = d2.Month & "_" & d2.Day & "_" & d2.Year
                TextBox6.Text = s1 & " to " & s2

                fileServiceTime = dirInfo.GetFiles("nonFatalErrors.txt", IO.SearchOption.AllDirectories)
                For Each file In fileServiceTime
                    zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
                    Dim objreader1 As New System.IO.StreamReader(zipinsideFolderDir1)
                    While objreader1.Peek <> -1
                        Dim StrFiles As String = objreader1.ReadLine()
                        If (StrFiles.Contains("Total service time:")) Then
                            timeList = addinList(StrFiles)
                        End If

                        StrFiles = String.Empty
                    End While
                    objreader1.Close()
                Next

                Caltime(timeList)

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        Dim AvgsizeOfDisk As Decimal = row.Cells(1).Value.ToString()
                        avgSize = avgSize + AvgsizeOfDisk
                    End If
                Next
                TextBox10.Text = Format((avgSize), "0.00") & "GB"
                TextBox3.Text = Format((avgSize / counterSizeFile), "0.00") & "GB"
                For Each row1 As DataGridViewRow In DataGridView2.Rows
                    If Not row1.IsNewRow Then
                        Dim avgProg1 As Integer = row1.Cells(0).Value.ToString()
                        avgProg = avgProg + avgProg1
                    End If
                Next
                TextBox5.Text = Format((avgProg / counterProgFile), "0.00")
                Button4.Enabled = True
            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString)
            End Try

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            '  System.Diagnostics.Process.Start(My.Computer.FileSystem.SpecialDirectories.Temp & "\ProgList.jar")
            dictTopProg.Clear()
            messageTopProg = String.Empty
            Dim filesOSdetails() As System.IO.FileInfo
            Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)
            Dim zipinsideFolderDir As String = Nothing
            filesOSdetails = dirInfo.GetFiles("olddriveinstalledPrograms.txt", IO.SearchOption.AllDirectories)
            For Each file In filesOSdetails
                zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
                Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
                While objreader.Peek <> -1
                    Dim StrFiles As String = objreader.ReadLine()
                    ' StrFiles = String.Empty

                    If Not (String.IsNullOrEmpty(StrFiles) Or StrFiles = " ") Then

                        dictTopProg = addtoMapCountTopProg(StrFiles)
                        StrFiles = Nothing
                    End If

                End While
                objreader.Close()
            Next
            Dim sortedDict = (From entry In dictTopProg Order By entry.Value Descending Select entry).Take(10)

            Dim pair As KeyValuePair(Of String, Integer)

            For Each pair In sortedDict
                'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                '  If (pair.Value <= 3) Then
                messageTopProg += pair.Key & ":  " & pair.Value & Environment.NewLine
                ' End If
            Next
            MessageBox.Show(messageTopProg)

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            If (String.IsNullOrEmpty(messageTopProg) Or messageTopProg = " " Or
                String.IsNullOrEmpty(messageLogVersion) Or messageLogVersion = " " Or
                String.IsNullOrEmpty(messageOSTransferDetails) Or messageOSTransferDetails = " ") Then
                'MessageBox.Show("Please click all buttons to Generate Log details")


                dictTopProg.Clear()
                messageTopProg = String.Empty
                Dim filesOSdetails() As System.IO.FileInfo
                Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)
                Dim zipinsideFolderDir As String = Nothing
                filesOSdetails = dirInfo.GetFiles("olddriveinstalledPrograms.txt", IO.SearchOption.AllDirectories)
                For Each file In filesOSdetails
                    zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
                    Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
                    While objreader.Peek <> -1
                        Dim StrFiles As String = objreader.ReadLine()
                        ' StrFiles = String.Empty

                        If Not (String.IsNullOrEmpty(StrFiles) Or StrFiles = " ") Then

                            dictTopProg = addtoMapCountTopProg(StrFiles)
                            StrFiles = Nothing
                        End If

                    End While
                    objreader.Close()
                Next
                Dim sortedDict = (From entry In dictTopProg Order By entry.Value Descending Select entry).Take(10)

                Dim pair As KeyValuePair(Of String, Integer)

                For Each pair In sortedDict
                    'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                    '  If (pair.Value <= 3) Then
                    messageTopProg += pair.Key & ":  " & pair.Value & Environment.NewLine
                    ' End If
                Next
                dictOS.Clear()
                messageOSTransferDetails = String.Empty
                Dim filesOSdetails1() As System.IO.FileInfo
                Dim dirInfo1 As New System.IO.DirectoryInfo(TextBox2.Text)
                Dim zipinsideFolderDir1 As String = Nothing, oldDriveString As String = Nothing, newDriveString As String = Nothing
                filesOSdetails1 = dirInfo1.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
                For Each file In filesOSdetails1
                    zipinsideFolderDir1 = file.DirectoryName + "\" + file.ToString
                    Dim objreader As New System.IO.StreamReader(zipinsideFolderDir1)
                    While objreader.Peek <> -1
                        Dim StrFiles As String = objreader.ReadLine()
                        If (StrFiles.Contains("Old drive is:")) Then
                            Dim words As String() = StrFiles.Split(New Char() {","c})
                            oldDriveString = words(1) & words(2)
                            StrFiles = String.Empty
                        End If
                        If (StrFiles.Contains("New drive is:")) Then
                            Dim words As String() = StrFiles.Split(New Char() {":"c})
                            newDriveString = words(1)
                            StrFiles = String.Empty
                        End If
                        If Not (String.IsNullOrEmpty(oldDriveString) Or String.IsNullOrEmpty(newDriveString)) Then
                            Dim mergeString As String = oldDriveString & " to " & newDriveString
                            dictOS = addtoMapCountOfOS(mergeString)
                            oldDriveString = Nothing
                            newDriveString = Nothing
                        End If

                    End While
                    objreader.Close()
                Next
                Dim sortedDict1 = (From entry In dictOS Order By entry.Value Descending Select entry)
                Dim pair1 As KeyValuePair(Of String, Integer)
                For Each pair1 In sortedDict1
                    'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                    messageOSTransferDetails += pair1.Key & ":  " & pair1.Value & Environment.NewLine
                Next

                dictLogVersion.Clear()
                messageLogVersion = String.Empty
                Dim fileslogVersiondetails() As System.IO.FileInfo
                Dim dirInfo2 As New System.IO.DirectoryInfo(TextBox2.Text)
                Dim zipinsideFolderDir2 As String = Nothing, VersionString As String = Nothing
                Dim CountVersion As Integer = 0
                fileslogVersiondetails = dirInfo2.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
                For Each file In fileslogVersiondetails
                    zipinsideFolderDir2 = file.DirectoryName + "\" + file.ToString
                    Dim objreader As New System.IO.StreamReader(zipinsideFolderDir2)
                    While objreader.Peek <> -1
                        Dim StrFiles As String = objreader.ReadLine()
                        If (StrFiles.Contains("FreshStart app version:")) Then
                            Dim words As String() = StrFiles.Split(New Char() {":"c})
                            VersionString = words(1)
                            dictLogVersion = addtoMapCountOfVersion(VersionString)
                            StrFiles = String.Empty
                        End If
                    End While
                    objreader.Close()
                Next
                Dim sortedDict2 = (From entry In dictLogVersion Order By entry.Value Descending Select entry)

                Dim pair2 As KeyValuePair(Of String, Integer)
                For Each pair2 In sortedDict2
                    'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                    messageLogVersion += pair2.Key & ":  " & pair2.Value & Environment.NewLine
                Next

                dictNewOSdetails.Clear()
                messageCountOfNewOS = String.Empty
                Dim filesOSdetails3() As System.IO.FileInfo
                Dim dirInfo3 As New System.IO.DirectoryInfo(TextBox2.Text)
                Dim zipinsideFolderDir3 As String = Nothing, newDriveString3 As String = Nothing
                filesOSdetails3 = dirInfo3.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
                For Each file In filesOSdetails3
                    zipinsideFolderDir3 = file.DirectoryName + "\" + file.ToString
                    Dim objreader As New System.IO.StreamReader(zipinsideFolderDir3)
                    While objreader.Peek <> -1
                        Dim StrFiles As String = objreader.ReadLine()
                        If (StrFiles.Contains("New drive is:")) Then
                            Dim words As String() = StrFiles.Split(New Char() {":"c})
                            newDriveString3 = words(1)
                            StrFiles = String.Empty
                        End If
                        If Not (String.IsNullOrEmpty(newDriveString3)) Then

                            dictNewOSdetails = addtoMapCountOfNewOS(newDriveString3)
                            newDriveString3 = Nothing
                        End If
                    End While
                    objreader.Close()
                Next
                Dim sortedDict3 = (From entry In dictNewOSdetails Order By entry.Value Descending Select entry)

                Dim pair3 As KeyValuePair(Of String, Integer)
                For Each pair3 In sortedDict3
                    'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                    messageCountOfNewOS += pair3.Key & ":  " & pair3.Value & Environment.NewLine
                Next
            End If
            Dim xlApp As Excel.Application
                xlApp = New Microsoft.Office.Interop.Excel.Application()
                If xlApp Is Nothing Then
                    MessageBox.Show("Excel is not properly installed!!")
                    Return
                End If
                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            xlWorkBook = xlApp.Workbooks.Add(misValue)

                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Cells(1, 1) = "Total Data Size"
                xlWorkSheet.Cells(1, 2) = TextBox10.Text
                xlWorkSheet.Cells(2, 1) = "Average Data Size"
                xlWorkSheet.Cells(2, 2) = TextBox3.Text
                xlWorkSheet.Cells(3, 1) = "Average Prog Installed"
                xlWorkSheet.Cells(3, 2) = TextBox5.Text
                xlWorkSheet.Cells(4, 1) = "Name and count of Top 10 Prog Installed "
                xlWorkSheet.Cells(4, 2) = messageTopProg
                xlWorkSheet.Cells(1, 6) = "Average FS time "
                xlWorkSheet.Cells(1, 7) = TextBox7.Text
                xlWorkSheet.Cells(2, 6) = "Minimum FS time "
                xlWorkSheet.Cells(2, 7) = TextBox8.Text
                xlWorkSheet.Cells(3, 6) = "Maximum FS time "
                xlWorkSheet.Cells(3, 7) = TextBox9.Text
                xlWorkSheet.Cells(4, 6) = "Old to New OS Count Details"
                xlWorkSheet.Cells(4, 7) = messageOSTransferDetails
                xlWorkSheet.Cells(4, 8) = "New OS count Details"
                xlWorkSheet.Cells(4, 9) = messageCountOfNewOS
                xlWorkSheet.Cells(4, 10) = "Count of FS version"
                xlWorkSheet.Columns.Range("G4").ColumnWidth = 50
                xlWorkSheet.Columns.Range("B4").ColumnWidth = 50
                xlWorkSheet.Columns.Range("I4").ColumnWidth = 30
                xlWorkSheet.Columns.Range("K4").ColumnWidth = 15
                xlWorkSheet.Cells(4, 11) = messageLogVersion
                xlWorkSheet.Columns("A").AutoFit()

                xlWorkSheet.Columns("F").AutoFit()
                xlWorkSheet.Columns("H").AutoFit()
                xlWorkSheet.Columns("J").AutoFit()
                'For i = 0 To DataGridView1.RowCount - 2
                '    For j = 0 To DataGridView1.ColumnCount - 1
                '        xlWorkSheet.Cells(i + 1, j + 1) =
                '            DataGridView1(j, i).Value.ToString()
                '    Next
                'Next
                Dim xlsxName As String = TextBox6.Text & ".xlsx"
                xlWorkSheet.SaveAs("C:\" & xlsxName)
                'xlWorkSheet.SaveAs("C:\" & TextBox6.Text & ".xlsx")
                xlWorkBook.Close()
                xlApp.Quit()
                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)

                MsgBox("You can find the file C:\" & xlsxName)


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            dictOS.Clear()
            messageOSTransferDetails = String.Empty
            Dim filesOSdetails() As System.IO.FileInfo
            Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)
            Dim zipinsideFolderDir As String = Nothing, oldDriveString As String = Nothing, newDriveString As String = Nothing
            filesOSdetails = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
            For Each file In filesOSdetails
                zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
                Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
                While objreader.Peek <> -1
                    Dim StrFiles As String = objreader.ReadLine()
                    If (StrFiles.Contains("Old drive is:")) Then
                        Dim words As String() = StrFiles.Split(New Char() {","c})
                        oldDriveString = words(1) & words(2)
                        StrFiles = String.Empty
                    End If
                    If (StrFiles.Contains("New drive is:")) Then
                        Dim words As String() = StrFiles.Split(New Char() {":"c})
                        newDriveString = words(1)
                        StrFiles = String.Empty
                    End If
                    If Not (String.IsNullOrEmpty(oldDriveString) Or String.IsNullOrEmpty(newDriveString)) Then
                        Dim mergeString As String = oldDriveString & " to " & newDriveString
                        dictOS = addtoMapCountOfOS(mergeString)
                        oldDriveString = Nothing
                        newDriveString = Nothing
                    End If

                End While
                objreader.Close()
            Next
            Dim sortedDict = (From entry In dictOS Order By entry.Value Descending Select entry)
            Dim pair As KeyValuePair(Of String, Integer)
            For Each pair In sortedDict
                'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                messageOSTransferDetails += pair.Key & ":  " & pair.Value & Environment.NewLine
            Next
            MessageBox.Show(messageOSTransferDetails)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            dictNewOSdetails.Clear()
            messageCountOfNewOS = String.Empty
            Dim filesOSdetails() As System.IO.FileInfo
            Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)
            Dim zipinsideFolderDir As String = Nothing, newDriveString As String = Nothing
            filesOSdetails = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
            For Each file In filesOSdetails
                zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
                Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
                While objreader.Peek <> -1
                    Dim StrFiles As String = objreader.ReadLine()
                    If (StrFiles.Contains("New drive is:")) Then
                        Dim words As String() = StrFiles.Split(New Char() {":"c})
                        newDriveString = words(1)
                        StrFiles = String.Empty
                    End If
                    If Not (String.IsNullOrEmpty(newDriveString)) Then

                        dictNewOSdetails = addtoMapCountOfNewOS(newDriveString)
                        newDriveString = Nothing
                    End If
                End While
                objreader.Close()
            Next
            Dim sortedDict = (From entry In dictNewOSdetails Order By entry.Value Descending Select entry)

            Dim pair As KeyValuePair(Of String, Integer)
            For Each pair In sortedDict
                'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                messageCountOfNewOS += pair.Key & ":  " & pair.Value & Environment.NewLine
            Next
            MessageBox.Show(messageCountOfNewOS)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            dictLogVersion.Clear()
            messageLogVersion = String.Empty
            Dim fileslogVersiondetails() As System.IO.FileInfo
            Dim dirInfo As New System.IO.DirectoryInfo(TextBox2.Text)
            Dim zipinsideFolderDir As String = Nothing, VersionString As String = Nothing
            Dim CountVersion As Integer = 0
            fileslogVersiondetails = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
            For Each file In fileslogVersiondetails
                zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
                Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
                While objreader.Peek <> -1
                    Dim StrFiles As String = objreader.ReadLine()
                    If (StrFiles.Contains("FreshStart app version:")) Then
                        Dim words As String() = StrFiles.Split(New Char() {":"c})
                        VersionString = words(1)
                        dictLogVersion = addtoMapCountOfVersion(VersionString)
                        StrFiles = String.Empty
                    End If
                End While
                objreader.Close()
            Next
            Dim sortedDict = (From entry In dictLogVersion Order By entry.Value Descending Select entry)

            Dim pair As KeyValuePair(Of String, Integer)
            For Each pair In sortedDict
                'Console.WriteLine("{0}, {1}", pair.Key, pair.Value)
                messageLogVersion += pair.Key & ":  " & pair.Value & Environment.NewLine
            Next
            MessageBox.Show(messageLogVersion)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try

    End Sub
End Class
