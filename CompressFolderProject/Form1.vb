Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions


Public Class Form1
    Dim directoryTargetLocation As String
    Dim Destinydirectory As String
    Dim exePath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "/7zFile/7z.exe"
    Dim timeList As List(Of TimeSpan) = New List(Of TimeSpan)
    Function extract7z(zipFileFolder As String, ToFolder As String)
        Dim args As String = "e " + zipFileFolder + " -o" + ToFolder + "" + " -p""cyberspa123"""
        System.Diagnostics.Process.Start(exePath, args)
        Threading.Thread.Sleep(1000)
        System.IO.File.Delete(zipFileFolder)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ZipFile.CreateFromDirectory(TextBox1.Text, TextBox2.Text)
            MessageBox.Show("Compress Succesfully!")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Threading.Thread.Sleep(500)
            System.IO.Directory.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "/7zFile", True)
            System.IO.Directory.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "\ProgList.jar", True)
            End
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FolderBrowserDialog1.Description = "Select directory"
        With FolderBrowserDialog1
            If .ShowDialog() = DialogResult.OK Then
                directoryTargetLocation = .SelectedPath
                TextBox1.Text = directoryTargetLocation.ToString
            End If
        End With
        FolderBrowserDialog1.Dispose()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            For Each foundFile1 As String In My.Computer.FileSystem.GetFiles(TextBox1.Text)
                Dim zipFolderpath As String = System.IO.Path.GetFullPath(TextBox2.Text & "/" & System.IO.Path.GetFileNameWithoutExtension(foundFile1))
                extract7z(foundFile1, zipFolderpath)
            Next
            MessageBox.Show("Unzip Successfully!")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
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
        Try
            FolderBrowserDialog1.Description = "Select directory"
            With FolderBrowserDialog1
                If .ShowDialog() = DialogResult.OK Then
                    directoryTargetLocation = .SelectedPath
                    Dim dirInfo As New System.IO.DirectoryInfo(directoryTargetLocation)

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
                    Dim d2 As Date = dateList.ElementAt(0)
                    TextBox6.Text = d1 & " to " & d2

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
                End If
            End With
            Caltime(timeList)

            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim AvgsizeOfDisk As Decimal = row.Cells(1).Value.ToString()
                    avgSize = avgSize + AvgsizeOfDisk
                End If
            Next
            TextBox3.Text = "Average Size of disk is: " & Format((avgSize / counterSizeFile), "0.00")
            For Each row1 As DataGridViewRow In DataGridView2.Rows
                If Not row1.IsNewRow Then
                    Dim avgProg1 As Integer = row1.Cells(0).Value.ToString()
                    avgProg = avgProg + avgProg1
                End If
            Next
            TextBox5.Text = (avgProg / counterProgFile)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\7zFile") Then
            File.WriteAllBytes(My.Computer.FileSystem.SpecialDirectories.Temp & "/_7zFile.zip", My.Resources._7zFile)
            ZipFile.ExtractToDirectory(My.Computer.FileSystem.SpecialDirectories.Temp & "/_7zFile.zip", My.Computer.FileSystem.SpecialDirectories.Temp)

            System.IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "/_7zFile.zip")
        End If
        File.WriteAllBytes(My.Computer.FileSystem.SpecialDirectories.Temp & "/ProgList.jar", My.Resources.ProgList)
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            System.Diagnostics.Process.Start(My.Computer.FileSystem.SpecialDirectories.Temp & "\ProgList.jar")

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim filesOSdetails() As System.IO.FileInfo
        Dim dirInfo As New System.IO.DirectoryInfo("C:\Users\Owner\Desktop\ErrorLog")
        Dim zipinsideFolderDir, oldDriveString, newDriveString As String
        Dim win7HP64Win7HP64, win7HP64Win7Pro64, win7HP64Win7Ult64, win7HP64Win8Std64, win7HP64Win8Pro64,
            win7HP64Win10H64, win7HP64Win10Pro64 As Integer
        filesOSdetails = dirInfo.GetFiles("frm2Errorlog.txt", IO.SearchOption.AllDirectories)
        For Each file In filesOSdetails
            zipinsideFolderDir = file.DirectoryName + "\" + file.ToString
            Dim objreader As New System.IO.StreamReader(zipinsideFolderDir)
            While objreader.Peek <> -1
                Dim StrFiles As String = objreader.ReadLine()
                If (StrFiles.Contains("Old drive is:")) Then
                    oldDriveString = StrFiles.Substring(18)
                    StrFiles = String.Empty
                End If
                If (StrFiles.Contains("New drive is:")) Then
                    newDriveString = StrFiles.Substring(14)
                    StrFiles = String.Empty
                End If
                If (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 7 Home Premium ,64bit")) Then
                    win7HP64Win7HP64 = win7HP64Win7HP64 + 1
                ElseIf (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 7 Professional  ,64bit")) Then
                    win7HP64Win7Pro64 = win7HP64Win7Pro64 + 1
                ElseIf (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 7 Ultimate  ,64bit")) Then
                    win7HP64Win7Ult64 = win7HP64Win7Ult64 + 1
                ElseIf (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 8 Standard  ,64bit")) Then
                    win7HP64Win8Std64 = win7HP64Win8Std64 + 1
                ElseIf (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 8 Professional  ,64bit")) Then
                    win7HP64Win8Pro64 = win7HP64Win8Pro64 + 1
                ElseIf (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 10 Home ,64bit")) Then
                    win7HP64Win10H64 = win7HP64Win10H64 + 1
                ElseIf (oldDriveString.Contains("Windows 7 Home Premium ,64bit") AndAlso
                    newDriveString.Contains("Windows 10 Professional ,64bit")) Then
                    win7HP64Win10Pro64 = win7HP64Win10Pro64 + 1
                End If

            End While
            objreader.Close()
        Next
    End Sub
End Class
