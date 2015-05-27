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
    Dim verify_by_1 As Integer = 0
    Dim timer As Integer = 60000
    Dim failed_to_copy As String()
    Dim failed_file_copy As IEnumerable(Of Integer)
    Dim file_list As New ArrayList()


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



        '   source_size = CountFiles(source, 0.0)
        ' DirectorySize(source, True)


    End Sub

    Private Sub Destionation_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog

        Dim targetDialog As FolderBrowserDialog = New FolderBrowserDialog
        If targetDialog.ShowDialog() = DialogResult.OK Then
            ' While targetDialog.SelectedPath.Length > 170
            'MsgBox("The selected path is too long. Please Choose a shorter path")
            'targetDialog = New FolderBrowserDialog
            'Label2.Text = "None Selected"
            'Label2.Refresh()


            'targetDialog.ShowDialog()



            'End While

            Label2.Text = targetDialog.SelectedPath
            target = targetDialog.SelectedPath
            filepath = target + "/" + machine_name + "_ERROR_LOG.txt"


            writer = New StreamWriter(filepath, True, System.Text.Encoding.ASCII)
            writer.WriteLine("**************************************************************")
            writer.WriteLine("*                    HASCopyBETA Failed Copy Log              *")
            writer.WriteLine("**************************************************************")
            ' fs = File.Create(filepath)
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

        ' CopyDirectory(source, target)
        ' My.Computer.FileSystem.CopyFile(source)








    End Sub



    Private Sub Reset_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Label1.Text = "None Specified"
        Label2.Text = "None Specified"
        ProgressBar1.Value = 0


        target_size = 0
        source_size = 0
        Label9.Text = "Exception file count"




        Label5.Text = target_size.ToString()

        Label5.Refresh()
        Label4.Text = "..."
        Label4.Refresh()






    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        
    End Sub











    Private Sub CopyDirectory(ByVal sourcePath As String, ByVal destPath As String)



        Dim srcInfo As New DirectoryInfo(source)
        Dim src As Long = DirectorySize(srcInfo, True)
        'MsgBox(src)

        'Dim trgInfo As New DirectoryInfo(target)

        'MsgBox(trg)


        If Not Directory.Exists(destPath) Then
            Directory.CreateDirectory(destPath)
        End If
        Try



            ' Use directly the sourcePath passed in, not the parent of that path
            For Each dir1 As String In Directory.GetDirectories(sourcePath)
                Dim destdir As String = Path.Combine(destPath, Path.GetFileName(dir1))
                CopyDirectory(dir1, destdir)
            Next



          




            For Each file1 As String In Directory.GetFiles(sourcePath)


                current_file = file1
                file_count = file_count + 1
                Dim dest As String = Path.Combine(destPath, Path.GetFileName(file1))
                If dest.Length > 255 Then
                    ex_file_count = ex_file_count + 1
                    list = list + file1 + Environment.NewLine
                    writer.WriteLine(list)


                End If
                File.Copy(file1, dest, True)  ' Added True here to force the an overwrite 
                Dim trg As Long = CountFiles(target, 0.0)
                Label10.Text = "Target Size: " & trg.ToString

                Label4.Text = file1.ToString
                Label4.Refresh()

                Label6.Text = "File Count:" & file_count.ToString
                Label6.Refresh()

                tot_prog = Math.Round((file_count / source_size), 2) * 100

                ProgressBar1.Value = tot_prog

                Label5.Text = tot_prog.ToString() & "Percent Complete"
                Label5.Refresh()

                Label11.Text = "source size: " & source_size
                Label11.Refresh()
            Next

            'Attempts to catch the error from any issue in copy a file ands it to a Text File
            'currently does not work
        Catch ex As Exception When TypeOf ex Is PathTooLongException OrElse TypeOf ex Is IOException OrElse TypeOf ex Is NullReferenceException
            'MsgBox(ex.ToString())
            ex_file_count = ex_file_count + 1
            Label9.Text = ex_file_count.ToString()
            'list = list + current_file + Environment.NewLine
            ' writer.WriteLine(list)





            'Dim error_file As Byte() = New UTF8Encoding(True).GetBytes(file1.ToString())
            'fs.Write(error_file, 0, error_file.Length)


        End Try



    End Sub

    Private Function DirectorySize(ByVal dInfo As DirectoryInfo, _
   ByVal includeSubDir As Boolean) As Long
        ' Enumerate all the files
        Dim totalSize As Long = dInfo.EnumerateFiles() _
          .Sum(Function(file) file.Length)

        ' If Subdirectories are to be included
        If includeSubDir Then
            ' Enumerate all sub-directories
            totalSize += dInfo.EnumerateDirectories() _
             .Sum(Function(dir) DirectorySize(dir, True))
        End If
        Return totalSize
    End Function



    Private Function CountFiles(InFolder As String, ByRef Result As Double)
        Result += IO.Directory.GetFiles(InFolder).Count


        For Each f As String In IO.Directory.GetDirectories(InFolder)
            CountFiles(f, Result)
        Next



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

    Protected Overrides Sub OnLoad(ByVal e1 As EventArgs)
        MyBase.OnLoad(e1)
        Label8.Text = machine_name

        'TO run parrelel or cross threads, needs something called delegates, to get by this the below is used: it basically tells the program not to check for it at all
        Control.CheckForIllegalCrossThreadCalls = False
        ' Label6.Visible = False
        Label5.Text = tot_prog & " % Complete"


    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e2 As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ' ProgressBar1.Maximum = source_size
        CopyDirectory(source, target)

    End Sub

   

    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        ' Cancel the asynchronous operation. 
        Me.BackgroundWorker1.CancelAsync()

        ' Disable the Cancel button.
        'cancelAsyncButton.Enabled = False

    End Sub 'cancelAsyncButton_Click



    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub
End Class
