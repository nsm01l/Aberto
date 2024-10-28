Public Class Form1

    ''' <summary>
    ''' D&amp;D された
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Label1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Label1.DragDrop
        ' 代入
        Dim filenames As String() = e.Data.GetData(DataFormats.FileDrop)

        Dim s As Integer = filenames.Length
        For i As Integer = 0 To s - 1
            Dim n As Integer = filenames(i).Length
            Dim str3 As String = filenames(i).Substring(n - 4, 4).ToLower
            Dim str4 As String = filenames(i).Substring(n - 5, 5).ToLower

            If (str3 = ".txt" Or str3 = ".htm" Or str3 = ".css" Or str4 = ".html") Then
                Dim proc As Process = New Process
                proc.StartInfo.FileName = "Notepad.exe"
                proc.StartInfo.Arguments = filenames(i)
                Try
                    proc.Start()
                Catch
                    Label1.BackColor = Color.FromArgb(255, 192, 192)
                    MsgBox("Error: " & filenames(i), MsgBoxStyle.OkOnly)
                End Try
            End If
        Next i
    End Sub

    ''' <summary>
    ''' D&amp;D が来た
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Label1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Label1.DragEnter
        ' 型判定
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Label1.BackColor = Color.FromArgb(192, 192, 255)
    End Sub

    Private Sub NotifyIcon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        Me.Show()
        Me.Focus()
    End Sub
End Class
