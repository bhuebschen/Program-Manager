Imports System.Security.Principal

Module VistaSecurity

    'Declare API
    Private Declare Ansi Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Const BCM_FIRST As Int32 = &H1600
    Private Const BCM_SETSHIELD As Int32 = (BCM_FIRST + &HC)

    Public Function GetUacShieldImage() As Bitmap
        Static shield_bm As Bitmap = Nothing
        If shield_bm Is Nothing Then
            Const WID As Integer = 50
            Const HGT As Integer = 50
            Const MARGIN As Integer = 4

            ' Make the button. For some reason, it must
            ' have text or the UAC shield won't appear.
            Dim btn As New Button
            btn.Text = " "
            btn.Size = New System.Drawing.Size(WID, HGT)
            AddShieldToButton(btn)

            ' Draw the button onto a bitmap.
            Dim bm As New Bitmap(WID, HGT)
            btn.Refresh()
            btn.DrawToBitmap(bm, New Rectangle(0, 0, WID, HGT))

            ' Find the part containing the shield.
            Dim min_x As Integer = WID
            Dim max_x As Integer = 0
            Dim min_y As Integer = WID
            Dim max_y As Integer = 0

            ' Fill on the left.
            For y As Integer = MARGIN To HGT - MARGIN - 1
                ' Get the leftmost pixel's color.
                Dim target_color As Color = bm.GetPixel(MARGIN, _
                    y)

                ' Fill in with this color as long as we see the
                ' target.
                For x As Integer = MARGIN To WID - MARGIN - 1
                    ' See if this pixel is part of the shield.
                    If bm.GetPixel(x, y).Equals(target_color) _
                        Then
                        ' It's not part of the shield.
                        ' Clear the pixel.
                        bm.SetPixel(x, y, Color.Transparent)
                    Else
                        ' It's part of the shield.
                        If min_y > y Then min_y = y
                        If min_x > x Then min_x = x
                        If max_y < y Then max_y = y
                        If max_x < x Then max_x = x
                    End If
                Next x
            Next y

            ' Clip out the shield part.
            Dim shield_wid As Integer = max_x - min_x + 1
            Dim shield_hgt As Integer = max_y - min_y + 1
            shield_bm = New Bitmap(16, 16)
            Dim shield_gr As Graphics = _
                Graphics.FromImage(shield_bm)
            shield_gr.DrawImage(bm, 0 + ((16 - shield_wid) \ 2), 0 + ((16 - shield_hgt) \ 2), _
                New Rectangle(min_x, min_y, shield_wid, _
                    shield_hgt), _
                GraphicsUnit.Pixel)
        End If

        Return shield_bm
    End Function

    Public Function IsVistaOrHigher() As Boolean
        Return Environment.OSVersion.Version.Major < 6
    End Function

    ' Checks if the process is elevated
    Public Function IsAdmin() As Boolean
        Dim id As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim p As WindowsPrincipal = New WindowsPrincipal(id)
        Return p.IsInRole(WindowsBuiltInRole.Administrator)
    End Function

    ' Add a shield icon to a button
    Public Sub AddShieldToButton(ByRef b As Button)
        b.FlatStyle = FlatStyle.System
        SendMessage(b.Handle, BCM_SETSHIELD, 0, &HFFFFFFFF)
    End Sub

    ' Restart the current process with administrator credentials
    Public Sub RestartElevated()
        Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        startInfo.UseShellExecute = True
        startInfo.WorkingDirectory = Environment.CurrentDirectory
        startInfo.FileName = Application.ExecutablePath
        startInfo.Verb = "runas"
        Try
            Dim p As Process = Process.Start(startInfo)
        Catch ex As Exception
            Return 'If cancelled, do nothing
        End Try
        'Application.Exit()
    End Sub

    Public Sub StartElevated(ByVal CMD As String)
        Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        startInfo.UseShellExecute = True
        startInfo.WorkingDirectory = Environment.CurrentDirectory
        startInfo.FileName = Application.ExecutablePath
        startInfo.Arguments = "HELPER " & CMD
        startInfo.Verb = "runas"
        Try
            Dim p As Process = Process.Start(startInfo)
        Catch ex As Exception
            Return 'If cancelled, do nothing
        End Try
        'Application.Exit()
    End Sub

End Module