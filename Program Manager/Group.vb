Imports System.Security
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class Group

    Private Sub Group_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Then
            e.Cancel = True
            Me.WindowState = FormWindowState.Minimized
        ElseIf e.CloseReason = CloseReason.UserClosing And Me.WindowState = FormWindowState.Minimized Then
            If MsgBox("Are you sure you want do delete the group '" & Me.Text & "'?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ' System.IO.Directory.Delete(ActiveMdiChild.Tag, True)
                Dim KX As Microsoft.Win32.RegistryKey = Nothing
                Dim PG As Microsoft.Win32.RegistryKey = Nothing
                KX = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE").OpenSubKey("Program Manager", True)
                Try
                    PG = KX.CreateSubKey("Groups", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                Catch ex As Exception

                End Try
                Try
                    PG = KX.OpenSubKey("Groups", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree).CreateSubKey(Me.Text, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                Catch ex As Exception

                End Try
                Try
                    PG = KX.OpenSubKey("Groups", True).OpenSubKey(Me.Text, True)

                Catch ex As Exception

                End Try
                PG.SetValue("Ignore", 1)
                e.Cancel = False
            End If
        ElseIf e.CloseReason = CloseReason.MdiFormClosing Then
            '
        Else
            e.Cancel = True
            Me.WindowState = FormWindowState.Minimized
        End If
    End Sub

    Private Sub Group_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If Not Form1.IsOnLoad Then
            Dim KX As Microsoft.Win32.RegistryKey = Nothing
            Dim PG As Microsoft.Win32.RegistryKey = Nothing
            KX = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE").OpenSubKey("Program Manager", True)
            If (KX Is Nothing) Then
                KX = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE", True).CreateSubKey("Program Manager", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
            End If
            Try
                PG = KX.CreateSubKey("Groups", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
            Catch ex As Exception

            End Try
            Try
                PG = KX.OpenSubKey("Groups", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree).CreateSubKey(Me.Text, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
            Catch ex As Exception

            End Try
            Try
                PG = KX.OpenSubKey("Groups", True).OpenSubKey(Me.Text, True)

            Catch ex As Exception

            End Try
            If WindowState = FormWindowState.Normal Then
                PG.SetValue("Width", Me.Width)
                PG.SetValue("Height", Me.Height)
                PG.SetValue("Left", Me.Left)
                PG.SetValue("Top", Me.Top)
            End If
            Try
                PG.SetValue("WindowState", Me.WindowState.ToString())
                PG.Close()
                KX.Close()

            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Group_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Group_Move(sender, e)
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If ListView1.SelectedItems.Count > 0 Then
            Try
                System.Diagnostics.Process.Start(ListView1.SelectedItems(0).Tag)
                Form1.LastFile = ListView1.SelectedItems(0).Tag
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Program Manager", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub RunAsAdministratorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunAsAdministratorToolStripMenuItem.Click
        If ListView1.SelectedItems.Count > 0 Then
            Dim processInfo As ProcessStartInfo = New ProcessStartInfo()
            processInfo.Verb = "runas"
            processInfo.UseShellExecute = True
            processInfo.FileName = ListView1.SelectedItems(0).Tag

            Dim pricipal As WindowsPrincipal = New WindowsPrincipal(WindowsIdentity.GetCurrent())
            Dim hasAdministrativeRight As Boolean = pricipal.IsInRole(WindowsBuiltInRole.Administrator)
            Try
                System.Diagnostics.Process.Start(processInfo)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub ListView1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.Click
        If ListView1.SelectedItems.Count > 0 Then
            Form1.LastFile = ListView1.SelectedItems(0).Tag
        End If
        RunAsAdministratorToolStripMenuItem.Enabled = ListView1.SelectedItems.Count > 0
    End Sub

    Private Sub Group_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Form1.LastFile = "GRP:" & Me.Text
    End Sub

    Private Sub Group_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ListView1.AllowDrop = True
        RunAsAdministratorToolStripMenuItem.Image = VistaSecurity.GetUacShieldImage()
    End Sub

    Private Sub ListView1_DragLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.DragLeave
        '    Stop
    End Sub

    'Private Sub ListView1_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles ListView1.ItemDrag
    '    ListView1.DoDragDrop(ListView1.SelectedItems, DragDropEffects.Move)
    'End Sub
End Class