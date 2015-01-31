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
    Dim tot_prog As Double

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
    Private Sub Source_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sourceDialog As FolderBrowserDialog = New FolderBrowserDialog
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog

        If sourceDialog.ShowDialog() = DialogResult.OK Then
            Label1.Text = sourceDialog.SelectedPath
            source = sourceDialog.SelectedPath
            source_size = CountFiles(source, 0.0)
        End If
        If DialogResult.Cancel Then
            Return
        End If



        '   source_size = CountFiles(source, 0.0)
        ' DirectorySize(source, True)


    End Sub

    Private Sub Destionation_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog

        Dim targetDialog As FolderBrowserDialog = New FolderBrowserDialog
        If targetDialog.ShowDialog() = DialogResult.OK Then
            Label2.Text = targetDialog.SelectedPath
            target = targetDialog.SelectedPath
            filepath = target + "/" + machine_name + "_ERROR_LOG.txt"
            writer = New StreamWriter(filepath, True, System.Text.Encoding.ASCII)
            'fs = File.Create(filepath)
            'Dim info As Byte() = New UTF8Encoding(True).GetBytes("This is a test")
            'fs.Write(info, 0, info.Length)
            'fs.Close()
        End If
        If DialogResult.Cancel Then
            Return
        End If



        ' Dim file_writer As New System.IO.StreamWriter(filepath)
        'file_writer.WriteLine("This is a test")
        'target_size = DirectorySize(target, True)


    End Sub

    Private Sub Start_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Process.Start("CMD", "/c cd  " + source + " & start cmd.exe /c Robocopy " + """" + source + """" + " " + """" + target + """" + " /S") '" /LOG+:log.txt" + " /S")
        'target_size = DirectorySize(target, True)
        'While target_size IsNot source_size
        ' temp = target_size
        'target_tot = target_tot + temp
        ' tot_prog = Math.Round(target_size)

        'temp = 0
        'target_size = DirectorySize(target, True)
        'tot_prog = Math.Round(target_size / source_size)
        'Label5.Text = tot_prog.ToString()
        'Label5.Refresh()
        'End While


        'ProgressBar1.Style = ProgressBarStyle.Continuous


        'bw.RunWorkerAsync()
        BackgroundWorker1.RunWorkerAsync()

        ' CopyDirectory(source, target)
        ' My.Computer.FileSystem.CopyFile(source)








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


        'trg_dir = New IO.DirectoryInfo(target)
        'src_dir = New IO.DirectoryInfo(source)
        Try


            'tot_prog = (trg_dir.GetFiles.Length / src_dir.GetFiles.Length)




        Catch ex As Exception
            ' MsgBox(src_dir.FullName())


        End Try
        ' ProgressBar1.Increment(tot_prog)


        'Label5.Text = ProgressBar1.Value


        ' Label5.Refresh()


        If ProgressBar1.Value = 100 Then
            Timer1.Stop()
            'Form2.Show()

        End If
    End Sub











    Private Sub CopyDirectory(ByVal sourcePath As String, ByVal destPath As String)

        'Does not work correctly ...
        '  If Label1.Text = "None Selected" Or Label2.Text = "None Selected" Then
        'MsgBox("Source or Destination not specified. \n Please specify a Source and a Destination.")

        ' End If

        If Not Directory.Exists(destPath) Then
            Directory.CreateDirectory(destPath)
        End If
        Try
            For Each file1 As String In Directory.GetFiles(sourcePath)
                ' Label4.Text = file1.ToString
                'Label4.Refresh()



                Dim dest As String = Path.Combine(destPath, Path.GetFileName(file1))
                File.Copy(file1, dest, True)  ' Added True here to force the an overwrite
                target_size = CountFiles(target, 0.0)
                Label4.Text = file1.ToString
                Label4.Refresh()

                ' My.Computer.FileSystem.CopyFile(file1, dest, FileIO.UIOption.AllDialogs)
                ' target_size = DirectorySize(destPath, True)
                'temp = target_size
                '  target_tot = target_tot + temp


                tot_prog = Math.Round((target_size / source_size), 2) * 100
                ProgressBar1.Value = tot_prog
                ' Dim test = ProgressBar1.Value

                Label5.Text = tot_prog.ToString()
                Label5.Refresh()

            Next

            ' Use directly the sourcePath passed in, not the parent of that path
            For Each dir1 As String In Directory.GetDirectories(sourcePath)
                dir_exc = dir1
                Dim destdir As String = Path.Combine(destPath, Path.GetFileName(dir1))
                CopyDirectory(dir1, destdir)
            Next
            'Attempts to catch the error from any issue in copy a file ands it to a Text File
            'currently does not work
        Catch e As Exception

            list = list + dir_exc + Environment.NewLine
            writer.WriteLine(list)



            'My.Application.Log.WriteException(e, TraceEventType.Error, filepath & ".")
            ' IO.File.AppendAllText(filepath, String.Format("{0}{1}", Environment.NewLine, e.ToString()))




            'MsgBox(e)
            'Dim error_file As Byte() = New UTF8Encoding(True).GetBytes(file1.ToString())
            'fs.Write(error_file, 0, error_file.Length)
            ' fs.Close()

        End Try

    End Sub
    Private Function CountFiles(InFolder As String, ByRef Result As Double)
        Result += IO.Directory.GetFiles(InFolder).Count
        Dim temp = ProgressBar1.Value

        For Each f As String In IO.Directory.GetDirectories(InFolder)
            CountFiles(f, Result)
        Next
        If temp = 15 Then

            writer.Close()
        End If
        'This is a problem...once the prog bar hits 100 percent it stays there and there there is an infinite msgbox loop
        'MsgBox("Copy Complete!")
        'Return Result
        'End If


        Return Result

    End Function





    Private Function DirectorySize(ByVal sPath As String, ByVal bRecursive As Boolean) As Double

        'Dim lngNumberOfDirectories As Long
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

            ' File not found. Take no action

        Catch exx As Exception

            ' Another error occurred

            Return 0

        End Try
        Return Size
    End Function

    Protected Overrides Sub OnLoad(ByVal e As EventArgs)
        MyBase.OnLoad(e)

        Label8.Text = machine_name


        '  Dim string_to_write As New StringBuilder()
        ' For Each txtName As String In Directory
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        CopyDirectory(source, target)
        Timer1.Start()
    End Sub
End Class
