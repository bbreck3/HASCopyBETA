Option Explicit On
Imports System.IO
Imports System.Text
Imports System.ComponentModel

Public Class HASCopyBETA

    Dim source As String
    Dim target As String
    Dim sourcePath As String '
    Dim targetPath As String
    Dim file1 As String

    ' Dim numFiles As  _
    '    System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    Dim source_list As New List(Of String)
    Dim target_list As New List(Of String)
    Dim files() As String

    Dim fileName As String
    Dim fileNames As String
    ' Dim Form2 As New Form2
    Dim filesTransferred As Integer = 0
    Dim fso = CreateObject("Scripting.FileSystemObject")
    Dim srclist As List(Of String)
    Dim Trglist As List(Of String)
    Dim src As Integer
    Dim tot_prog As Double = 0.0

    Dim trg_dir
    Dim src_dir
    Dim source_size
    Dim target_size
    Private Property WScript As Object
    Dim target_tot = 0
    Dim temp = 0
    Dim machine_name = My.Computer.Name
    Dim filepath As String
    Dim fs As FileStream
    Dim writer As StreamWriter
    Dim filereader As StreamReader
    Dim dir_exc As String
    Dim list As String
    Dim current_file As String
    Dim file_count As Double = 1.0
    Dim ex_file_count As Integer = 0
    Dim dir_count As Double = 1.0
    Dim verify_by_1 As Integer = 0


    Private Sub Source_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sourceDialog As FolderBrowserDialog = New FolderBrowserDialog
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog

        If sourceDialog.ShowDialog() = DialogResult.OK Then
            Label1.Text = sourceDialog.SelectedPath
            source = sourceDialog.SelectedPath
            source_size = CountFiles(source, 0.0)
            Label11.Text = "Source Size: " & source_size
            Label11.Refresh()
        End If
        If DialogResult.Cancel Then
            Return
        End If
    End Sub

    Sub GetDirectories(ByVal StartPath As String, ByRef DirectoryList As ArrayList)
        Try
            Dim DirList As New ArrayList
            Dim Dirs() As String = Directory.GetDirectories(StartPath)
            DirList.AddRange(Dirs)
            For Each Dir As String In Dirs
                GetDirectories(Dir, DirectoryList)
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Destionation_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog

        Dim targetDialog As FolderBrowserDialog = New FolderBrowserDialog
        If targetDialog.ShowDialog() = DialogResult.OK Then
            Label2.Text = targetDialog.SelectedPath
            Label9.Text = targetDialog.SelectedPath.Length
            target = targetDialog.SelectedPath
            filepath = target + "/" + machine_name + "_ERROR_LOG.txt"
            writer = New StreamWriter(filepath, True, System.Text.Encoding.ASCII)
            writer.WriteLine("**************************************************************")
            writer.WriteLine("*                    HASCopyBETA Data Migration Report        *")
            writer.WriteLine("**************************************************************")

            'fs = File.Create(filepath)
            'Dim info As Byte() = New UTF8Encoding(True).GetBytes("This is a test")
            'fs.Write(info, 0, info.Length)
            'fs.Close()
        End If
        If DialogResult.Cancel Then
            Return
        End If





    End Sub

    Private Sub Start_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Process.Start("CMD", "/c cd  " + source + " & start cmd.exe /c Robocopy " + """" + source + """" + " " + """" + target + """" + " /S") '" /LOG+:log.txt" + " /S")

        BackgroundWorker1.WorkerSupportsCancellation = True
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker1.RunWorkerAsync()

    End Sub

   

    Private Sub Reset_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Label1.Text = "None Specified"
        Label2.Text = "None Specified"
        ProgressBar1.Value = 0

        target_size = 0

        Label5.Text = target_size.ToString()

        Label5.Refresh()
        Label4.Text = "..."
        Label4.Refresh()
    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub

    Sub DirectoryCopy( _
         ByVal sourceDirName As String, _
         ByVal destDirName As String, _
         ByVal copySubDirs As Boolean)

        ' Get the subdirectories for the specified directory.
        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Try

            If Not dir.Exists Then
                Throw New DirectoryNotFoundException( _
                    "Source directory does not exist or could not be found: " _
                    + sourceDirName)
            End If

            Dim dirs As DirectoryInfo() = dir.GetDirectories()
            ' If the destination directory doesn't exist, create it.
            If dirs.Length > 260 Then
                Dim failedFiles As FileInfo() = dir.GetFiles()
                For Each file In failedFiles
                    file_count = file_count + 1
                    Label6.Text = file_count.ToString
                    Label6.Refresh()
                    Label4.Text = file.ToString
                    Label4.Refresh()

                Next file

            End If
            If Not Directory.Exists(destDirName) Then
                Directory.CreateDirectory(destDirName)
            End If
            ' Get the files in the directory and copy them to the new location.
            Dim files As FileInfo() = dir.GetFiles()
            For Each file In files
                If file.Length > 240 Then
                    Label6.Text = file_count.ToString
                    Label6.Refresh()
                    Label4.Text = file.ToString
                    Label4.Refresh()
                End If
                Dim temppath As String = Path.Combine(destDirName, file.Name)
                file.CopyTo(temppath, False)
                file_count = file_count + 1
                Label6.Text = file_count.ToString
                Label6.Refresh()
                Label4.Text = file.ToString
                Label4.Refresh()
            Next file

            ' If copying subdirectories, copy them and their contents to new location.
            If copySubDirs Then
                For Each subdir In dirs
                   
                    Dim temppath As String = Path.Combine(destDirName, subdir.Name)
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs)
                Next subdir
            End If

        Catch pathToLongException As PathTooLongException

            ' file_count = file_count + 1
            ' Label6.Text = file_count.ToString
            'Label6.Refresh()
        Catch ex As Exception

            'file_count = file_count + 1
            'Label6.Text = file_count.ToString
            'Label6.Refresh()
        End Try
    End Sub

    Private Function CountFiles(InFolder As String, ByRef Result As Double)
        Result += IO.Directory.GetFiles(InFolder).Count


        For Each f As String In IO.Directory.GetDirectories(InFolder)
            CountFiles(f, Result)
        Next

        Return Result

    End Function





    Private Function DirectorySize(ByVal sPath As String, ByVal bRecursive As Boolean) As Double


        Dim lngNumberOfDirectories As Double
        Dim Size As Long = 0
        Dim diDir As New DirectoryInfo(sPath)
        Try

            Dim fil As FileInfo

            For Each fil In diDir.GetFiles()

                Size += fil.Length

            Next fil

            If bRecursive = True Then

                Dim diSubDir As DirectoryInfo

                For Each diSubDir In diDir.GetDirectories()

                    Size += DirectorySize(diSubDir.FullName, True)

                    lngNumberOfDirectories += 1

                Next diSubDir

            End If

            Return Size

        Catch ex As System.IO.FileNotFoundException
        Catch exx As Exception

            Return 0

        End Try
        Return Size
    End Function

    Protected Overrides Sub OnLoad(ByVal e1 As EventArgs)
        MyBase.OnLoad(e1)
        Label8.Text = machine_name

        'TO run parrelel or cross threads, needs something called delegates, to get by this the below is used: it basically tells the program not to check for it at all
        Control.CheckForIllegalCrossThreadCalls = False
        ' Label6.Visible = False
        Label5.Text = tot_prog & " % Complete"



    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e2 As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'Process.Start("CMD", "/c cd  " + source + " & start cmd.exe /c Robocopy " + """" + source + """" + " " + """" + target + """" + " /LOG+:log.txt" + " /S /TEE")
        DirectoryCopy(source, target, True)
    End Sub
    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

    End Sub

    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        ' Cancel the asynchronous operation. 
        Me.BackgroundWorker1.CancelAsync()

        ' Disable the Cancel button.
        'cancelAsyncButton.Enabled = False

    End Sub
End Class
