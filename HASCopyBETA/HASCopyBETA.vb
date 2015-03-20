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
            Label2.Text = targetDialog.SelectedPath
            target = targetDialog.SelectedPath
            filepath = target + "/" + machine_name + "_ERROR_LOG.txt"
            writer = New StreamWriter(filepath, True, System.Text.Encoding.ASCII)
            writer.WriteLine("**************************************************************")
            writer.WriteLine("*                    HASCopyBETA Failed Copy Log              *")
            writer.WriteLine("**************************************************************")

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




        If Not Directory.Exists(destPath) Then
            Directory.CreateDirectory(destPath)
        End If
        Try


            ' ProgressBar1.Maximum = source_size




            For Each file1 As String In Directory.GetFiles(sourcePath)
                current_file = file1

                file_count = file_count + 1
                ' Label4.Text = file1.ToString
                'Label4.Refresh()



                Dim dest As String = Path.Combine(destPath, Path.GetFileName(file1))
                File.Copy(file1, dest, True)  ' Added True here to force the an overwrite

                'target_size = target_size + file_count

                Label4.Text = file1.ToString
                Label4.Refresh()
                'file_count = file_count + 1
                ' My.Computer.FileSystem.CopyFile(file1, dest, FileIO.UIOption.AllDialogs)
                ' target_size = DirectorySize(destPath, True)
                temp = target_size
                'target_tot = target_tot + temp


                tot_prog = Math.Round((file_count / source_size), 2) * 100

                Label11.Text = "Source Size: " & source_size
                Label11.Refresh()
                'Math.Round((file_count / source_size), 2) * 100

                ProgressBar1.Value = tot_prog
                If (ProgressBar1.Value = 100 And file_count = source_size) Or (ProgressBar1.Value = 100 And (file_count + ex_file_count) = source_size) Then

                    MsgBox("Copy Complete")
                    writer.Close()
                End If
                'This is a problem:
                'the 
                ' BackgroundWorker1.ReportProgress(tot_prog)

                Label6.Text = "Total file count: " & file_count.ToString
                Label6.Refresh()

                Label5.Text = tot_prog & " % Complete"
                Label5.Refresh()
                System.Threading.Thread.Sleep(200)


                ' Dim test = ProgressBar1.Value



            Next




            ' Use directly the sourcePath passed in, not the parent of that path
            For Each dir1 As String In Directory.GetDirectories(sourcePath)

                Dim destdir As String = Path.Combine(destPath, Path.GetFileName(dir1))
                CopyDirectory(dir1, destdir)
            Next


            ' Next
            'Attempts to catch the error from any issue in copy a file ands it to a Text File
            'currently does not work
        Catch ex As Exception When TypeOf ex Is NullReferenceException OrElse TypeOf ex Is PathTooLongException
            verify_by_1 = verify_by_1 + 1
            Label10.Text = "Verify count by one only: " & verify_by_1
            Label10.Refresh()

            ' MsgBox("file is is to long to copy: " + current_file)


            ex_file_count = ex_file_count + 1
            ' MsgBox(ex.ToString)
            'MsgBox(current_file.ToString)
            ' MsgBox(file_count)
            'Label9.Text = ex_file_count.ToString
            ' Label9.Refresh()

            list = list + current_file + Environment.NewLine
            writer.WriteLine(list)



            'ex_file_count = ex_file_count + 1

            ''tot_prog = (((file_count + ex_file_count) / (source_size))) * 100
            ' file_count = file_count + 1
            Label9.Text = "Exception file count totatl: " & ex_file_count.ToString
            Label9.Refresh()
            ' Label6.Text = "Total file count: " & file_count.ToString
            ' Label6.Refresh()

            Label11.Text = "Source Size: " & source_size
            Label11.Refresh()





            '  tot_prog = (((file_count + ex_file_count) / (source_size))) * 100
            'Math.Round((file_count / source_size), 2) * 100
            ' ProgressBar1.Value = tot_prog
            ' BackgroundWorker1.ReportProgress(tot_prog)


            verify_by_1 = 0

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
        ' Dim temp = ProgressBar1.Value

        For Each f As String In IO.Directory.GetDirectories(InFolder)
            CountFiles(f, Result)
        Next

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

    Protected Overrides Sub OnLoad(ByVal e1 As EventArgs)
        MyBase.OnLoad(e1)
        Label8.Text = machine_name

        'TO run parrelel or cross threads, needs something called delegates, to get by this the below is used: it basically tells the program not to check for it at all
        Control.CheckForIllegalCrossThreadCalls = False
        ' Label6.Visible = False
        Label5.Text = tot_prog & " % Complete"


        '  Dim string_to_write As New StringBuilder()
        ' For Each txtName As String In Directory
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e2 As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ' ProgressBar1.Maximum = source_size


        CopyDirectory(source, target)














        '' any cleanup code go here
        '' ensure that you close all open resources before exitting out of this Method.
        '' try to skip off whatever is not desperately necessary if CancellationPending is True

        '' set the e.Cancel to True to indicate to the RunWorkerCompleted that you cancelled out
        ' If BackgroundWorker1.CancellationPending Then
        'e.Cancel = True
        ' BackgroundWorker1.ReportProgress(100, "Cancelled.")
        '  End If
    End Sub

    'Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e2 As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

    ' ProgressBar1.Value = tot_prog












    ' End Sub
    'Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e2 As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
    '  MsgBox("Copy Complete")
    ''open pdf here
    '  writer.Close()


    ' End Sub

    'Empliments later....cause early termination of the background worker

    ' Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e2 As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
    '' This event is fired when your BackgroundWorker exits.
    '' It may have exitted Normally after completing its task, 
    '' or because of Cancellation, or due to any Error.

    '  If e2.Error IsNot Nothing Then
    '' if BackgroundWorker terminated due to error
    '   MessageBox.Show(e2.Error.Message)
    '  Label1.Text = "Error occurred!"

    ' ElseIf e2.Cancelled Then
    '' otherwise if it was cancelled
    '     MessageBox.Show("Task cancelled!")
    '     Label1.Text = "Task Cancelled!"

    '   Else
    '' otherwise it completed normally
    '       MessageBox.Show("Task completed!")
    '      Label1.Text = "Error completed!"
    '   End If
    '
    ' Button1.Enabled = True
    ' Button2.Enabled = False
    ' End Sub

    ' Timer1.Start()

    ' Private Sub Cancel_Click(sender As Object, e1 As EventArgs) Handles Button5.Click
    '    Me.BackgroundWorker1.CancelAsync()

    '   MsgBox("canceled")
    'End Sub

    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        ' Cancel the asynchronous operation. 
        Me.BackgroundWorker1.CancelAsync()

        ' Disable the Cancel button.
        'cancelAsyncButton.Enabled = False

    End Sub 'cancelAsyncButton_Click



    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub
End Class
