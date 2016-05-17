Imports System
Imports System.Runtime
Imports System.Runtime.InteropServices

Public Class Form1

    Public LastFile As String = ""
    Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
    Public IsOnLoad As Boolean = True
    Declare Ansi Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As IntPtr, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As IntPtr) As Integer
    Public Const CSIDL_COMMON_STARTMENU As Integer = &H16
    Private Declare Auto Function SHGetSpecialFolderLocation Lib "shell32" _
            (ByVal hwndOwner As IntPtr, ByVal nFolder As Integer, ByRef ppidl As _
            Integer) As Integer

    Private Declare Auto Function SHGetPathFromIDList Lib "shell32" _
        (ByVal pidl As Integer, ByVal pszPath As System.Text.StringBuilder) As _
        Boolean

    Const ProgramsFolder As String = "Programs"

    Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Integer)
    Public CommonStartMenu As String = GetSpecialFolderLocation(Me.Handle, CSIDL_COMMON_STARTMENU) & "\" & ProgramsFolder

    Public Function GetSpecialFolderLocation(ByVal hWnd As IntPtr, ByVal csidl _
      As Integer) As String
        Dim path As New System.Text.StringBuilder(260)
        Dim pidl As Integer
        Dim rpath As String = String.Empty

        If SHGetSpecialFolderLocation(hWnd, csidl, pidl) = 0 Then
            If SHGetPathFromIDList(pidl, path) Then
                rpath = path.ToString()
            End If
            CoTaskMemFree(pidl)
        End If
        Return rpath
    End Function

    Dim hinst As IntPtr
    Dim myTray As IntPtr


    Private Structure CPLINFO
        Dim idIcon As Int32
        Dim idName As Int32
        Dim idInfo As Int32
        Dim lData As Int32
    End Structure

    Private Const CPL_INIT = 1
    Private Const CPL_GETCOUNT = 2
    Private Const CPL_INQUIRE = 3
    Private Const CPL_SELECT = 4
    Private Const CPL_DBLCLK = 5
    Private Const CPL_STOP = 6
    Private Const CPL_EXIT = 7
    Private Const CPL_NEWINQUIRE = 8

    Private ci As CPLINFO

    Private Declare Function NewLinkHere Lib "appwiz.cpl" Alias "NewLinkHereW" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long


    Dim KX As Microsoft.Win32.RegistryKey


    Public Class Impersonation
        <DllImport("advapi32")> _
        Public Shared Function LogonUser(ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, _
        ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, ByRef phToken As Integer) As Boolean
        End Function

        <DllImport("kernel32")> _
        Public Shared Function GetLastError() As Integer
        End Function
    End Class

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If InStr(Command, "HELPER") > 0 Then
            If InStr(Command, "CREATELINK:") > 0 Then
                Dim FN As String = Mid(Command, InStr(Command, "CREATELINK:") + 11)
                System.IO.File.Create(Mid(Command, InStr(Command, "CREATELINK:") + 11)).Close()
                With New System.Diagnostics.Process
                    .StartInfo.UseShellExecute = True
                    .StartInfo.Arguments = "appwiz.cpl,NewLinkHere " & FN
                    .StartInfo.FileName = "rundll32.exe"
                    .Start()
                    .WaitForExit()
                    Dim R As Int32 = .ExitCode
                End With


            End If

            Application.Exit()
            Exit Sub
        End If
        Me.Text = "Program Manager - " & Environment.UserDomainName.ToUpper & "\" & Environment.UserName

        'InitSystemTray()
        Try
            Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE", True).OpenSubKey("Program Manager", True)
        Catch ex As Exception

        End Try
        'Show()
        Try
            KX = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE", True).OpenSubKey("Program Manager", True)
            Me.Size = New Size(KX.GetValue("Width", 1024), KX.GetValue("Height", 768))
            Me.Location = New Point(KX.GetValue("Left", 100), KX.GetValue("Top", 100))
        Catch ex As Exception
            KX = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE", True).CreateSubKey("Program Manager", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
            Me.Size = New Size(1024, 768)
            Me.Location = New Point(100, 100)
        End Try

        Try
            '' COMMON START MENU!
            Dim M() As String = System.IO.Directory.GetDirectories(CommonStartMenu)
            For i As Integer = 0 To M.Length - 1
                Dim PG As Microsoft.Win32.RegistryKey
                Dim IG As Boolean = False
                Try
                    PG = KX.OpenSubKey("Groups").OpenSubKey(System.IO.Path.GetFileName(M(i)) & " (All Users)", True)
                    If PG.GetValue("Ignore", 0) = 1 Then
                        IG = True
                    Else
                        IG = False
                    End If
                Catch ex As Exception
                    Try
                        KX.CreateSubKey("Groups", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                    Catch ex2 As Exception

                    End Try
                    PG = KX.OpenSubKey("Groups", True).CreateSubKey(System.IO.Path.GetFileName(M(i)) & " (All Users)", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                    PG.SetValue("WindowState", "Normal")
                    PG.SetValue("Width", 500)
                    PG.SetValue("Height", 500)
                End Try
                If Not IG Then

                    With New Group
                        .MdiParent = Me
                        .Tag = M(i)
                        .Text = System.IO.Path.GetFileName(M(i)) & " (All Users)"
                        Dim F() As String = System.IO.Directory.GetFiles(M(i), "*.lnk")
                        For j As Integer = 0 To F.Length - 1
                            Dim II As Icon = Icon.ExtractAssociatedIcon(F(j))
                            .ImageList1.Images.Add(F(j), II.ToBitmap())
                            With .ListView1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(F(j)), F(j))
                                .Tag = F(j)
                            End With
                            Try
                                .Size = New Size(PG.GetValue("Width", 500), PG.GetValue("Height", 400))
                                .Show()
                                .Left = PG.GetValue("Left", 100)
                                .Top = PG.GetValue("Top", 100)
                                .WindowState = [Enum].Parse(GetType(FormWindowState), PG.GetValue("WindowState", "Normal"), True)
                            Catch ex As Exception

                            End Try
                        Next
                        .Show()
                    End With
                End If
            Next
        Catch ex As Exception

        End Try

        Try


            '' USER START MENU!!
            Dim M() As String = System.IO.Directory.GetDirectories(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) & "\" & ProgramsFolder)
            For i As Integer = 0 To M.Length - 1
                Dim PG As Microsoft.Win32.RegistryKey
                Dim IG As Boolean = False
                Try
                    PG = KX.OpenSubKey("Groups", True).OpenSubKey(System.IO.Path.GetFileName(M(i)), True)
                    If PG.GetValue("Ignore", 0) = 1 Then
                        IG = True
                    Else
                        IG = False
                    End If
                Catch ex As Exception
                    Try
                        PG = KX.OpenSubKey("Groups", True).CreateSubKey(System.IO.Path.GetFileName(M(i)), Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                        IG = False
                    Catch ex2 As Exception

                    End Try
                End Try
                If Not IG Then

                    With New Group
                        .MdiParent = Me
                        .Tag = M(i)
                        .Text = System.IO.Path.GetFileName(M(i))
                        Dim F() As String = System.IO.Directory.GetFiles(M(i), "*.lnk")
                        For j As Integer = 0 To F.Length - 1
                            Dim II As Icon = Icon.ExtractAssociatedIcon(F(j))
                            .ImageList1.Images.Add(F(j), II.ToBitmap())
                            With .ListView1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(F(j)), F(j))
                                .Tag = F(j)
                            End With
                            Try
                                .Size = New Size(PG.GetValue("Width", 500), PG.GetValue("Height", 400))
                                .Show()
                                .Left = PG.GetValue("Left", 100)
                                .Top = PG.GetValue("Top", 100)
                                .WindowState = [Enum].Parse(GetType(FormWindowState), PG.GetValue("WindowState", "Normal"), True)
                            Catch ex As Exception

                            End Try
                        Next
                        .Show()
                    End With
                End If
            Next


            With New Group
                .MdiParent = Me
                .Text = "Main"
                Dim F() As String = System.IO.Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) & "\" & ProgramsFolder, "*.lnk")
                For j As Integer = 0 To F.Length - 1
                    Dim II As Icon = Icon.ExtractAssociatedIcon(F(j))
                    .ImageList1.Images.Add(F(j), II.ToBitmap())
                    With .ListView1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(F(j)), F(j))

                    End With

                Next
                F = System.IO.Directory.GetFiles(CommonStartMenu, "*.lnk")
                For j As Integer = 0 To F.Length - 1
                    Dim II As Icon = Icon.ExtractAssociatedIcon(F(j))
                    .ImageList1.Images.Add(F(j), II.ToBitmap())
                    With .ListView1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(F(j)), F(j))

                    End With

                Next
                .WindowState = FormWindowState.Minimized
                .Show()
            End With
        Catch ex As Exception

        End Try

        Me.LayoutMdi(MdiLayout.ArrangeIcons)
        IsOnLoad = False
    End Sub

    Private Sub Form1_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        Try
            If Not KX Is Nothing Then KX.Close()
            KX = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE").OpenSubKey("Program Manager", True)
            If Not KX Is Nothing Then
                KX.SetValue("Width", Me.Width)
                KX.SetValue("Height", Me.Height)
                KX.SetValue("Left", Me.Left)
                KX.SetValue("Top", Me.Top)
                KX.SetValue("WindowState", Me.WindowState.ToString())
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        e.Graphics.FillRectangle(Brushes.White, e.ClipRectangle)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaintBackground(e)
        e.Graphics.FillRectangle(Brushes.White, e.ClipRectangle)
    End Sub

    Private Sub WindowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WindowToolStripMenuItem.Click
    End Sub

    Private Sub Form1_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Form1_Move(sender, e)
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        With New dlgNew
            If InStr(ActiveMdiChild.Tag, "ProgramData") > 0 Then
                VistaSecurity.AddShieldToButton(.OK_Button)
            End If
            .ShowDialog(Me)
            'NewLinkHere(Me.Handle.ToInt32, 0, 0, 0)
                If .DialogResult = Windows.Forms.DialogResult.OK Then
                If .RadioButton2.Checked = True Then
                    VistaSecurity.StartElevated("CREATELINK:" & ActiveMdiChild.Tag & "\file.lnk")
                Else

                End If
            End If
        End With
    End Sub

    Private Sub ExitWindowsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitWindowsToolStripMenuItem.Click

    End Sub

    Private Sub RunToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunToolStripMenuItem.Click
        Dim uFlag As Long
        uFlag = (&H10 Or &H20 Or &H40 Or &H80)
        uFlag = uFlag / 16
        SHRunDialog(Me.Handle.ToInt32(), 0, 0, "Run", "&Command Line:", 0)
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)

    End Sub

    Private Sub TileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub MinimizeOnUseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MinimizeOnUseToolStripMenuItem.Click

    End Sub

    Private Sub PropertiesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PropertiesToolStripMenuItem.Click
        Dim fileInfo As New Win32.SHELLEXECUTEINFO
        fileInfo.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(fileInfo)
        fileInfo.lpVerb = "properties"
        fileInfo.fMask = Win32.Shell.SEE_MASK_INVOKEIDLIST
        fileInfo.nShow = 1
        fileInfo.lpFile = LastFile      'File name to display properties.
        Win32.Shell.ShellExecuteEx(fileInfo)

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If LastFile.StartsWith("GRP:") Then
            If MsgBox("Are you sure you want do delete the group '" & LastFile.Substring(4) & "'?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ' System.IO.Directory.Delete(ActiveMdiChild.Tag, True)
                KX.OpenSubKey("Groups", True).OpenSubKey(ActiveMdiChild.Text, True).SetValue("Ignore", 1)
                ActiveMdiChild.Close()
            End If
        End If
    End Sub

    Private Sub AboutProgramManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutProgramManagerToolStripMenuItem.Click
        ShellAbout(Me.Handle, "Program Manager", "", Me.Icon.Handle)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label1.Text = Now.ToString("HH:mm:ss")
    End Sub
End Class
