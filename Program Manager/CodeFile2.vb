<System.Serializable()> Public Class NotifyIconBar
    Inherits Control
    Implements IShellControl

    Dim tltTip As New ToolTip

    Public Sub New()
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.AllPaintingInWmPaint, True)
        tltTip.ShowAlways = True

        SyncLock SystemTray.Icons
            For Each i As TrayNotifyIcon In SystemTray.Icons
                _icons.Add(New NotifyIconData(i))
            Next
        End SyncLock
        AddHandler SystemTray.NotifyIconAdded, AddressOf IconAdded
        AddHandler SystemTray.NotifyIconBeforeRemoved, AddressOf IconRemoved
        AddHandler SystemTray.NotifyIconUpdated, AddressOf IconModified
    End Sub

    Public Overrides Function ToString() As String Implements IShellControl.ToString
        Return "Notify Icon Bar"
    End Function

    Private Class NotifyIconData
        Public TrayItem As TrayNotifyIcon
        Public Rect As Rectangle

        Public Sub New(ByVal i As TrayNotifyIcon)
            Me.TrayItem = i
        End Sub

    End Class

#Region "Properties"

    Dim _horizSpacing As Integer = 3

    <System.ComponentModel.DefaultValue(3)> Public Property HorizontalSpacing() As Integer
        Get
            Return _horizSpacing
        End Get
        Set(ByVal value As Integer)
            _horizSpacing = value
            Refresh()
        End Set
    End Property

    Dim _vertSpacing As Integer = 3

    <System.ComponentModel.DefaultValue(3)> Public Property VerticalSpacing() As Integer
        Get
            Return _vertSpacing
        End Get
        Set(ByVal value As Integer)
            _vertSpacing = value
            Refresh()
        End Set
    End Property

    Dim _bck As ShellRenderer = New VisualStyleRenderer(VisualStyles.VisualStyleElement.Taskbar.BackgroundBottom.Normal)

    Public Property Background() As ShellRenderer Implements IShellControl.Background
        Get
            Return _bck
        End Get
        Set(ByVal value As ShellRenderer)
            _bck = value
        End Set
    End Property

#End Region

    Dim _icons As New ObjectModel.Collection(Of NotifyIconData)

    Private Sub ArrangeItems()
        Dim currentx As Integer = HorizontalSpacing, currenty As Integer = VerticalSpacing
        If Me.Width > Me.Height Then
            For Each i As NotifyIconData In _icons
                i.Rect = New Rectangle(currentx, currenty, 16, 16)
                currentx += 16 + HorizontalSpacing
                If currentx + 16 + HorizontalSpacing > Me.Width Then
                    currentx = HorizontalSpacing
                    currenty += 16 + VerticalSpacing
                End If
            Next
        Else
            For Each i As NotifyIconData In _icons
                i.Rect = New Rectangle(currentx, currenty, 16, 16)
                currenty += 16 + VerticalSpacing
                If currenty + 16 + VerticalSpacing > Me.Height Then
                    currenty = VerticalSpacing
                    currentx += 16 + HorizontalSpacing
                End If
            Next
        End If
    End Sub

#Region "Overriden Events"

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)

        ArrangeItems()

        Dim b As New Bitmap(Me.Width, Me.Height, Imaging.PixelFormat.Format32bppArgb)
        Dim gr As Graphics = Graphics.FromImage(b)

        Me.BackColor = Color.FromKnownColor(KnownColor.ButtonFace)
        'Me.Background.Render(gr, New Rectangle(0, 0, Me.Width, Me.Height))

        SyncLock _icons
            For Each n As NotifyIconData In _icons
                If n.TrayItem.Icon IsNot Nothing Then
                    Try
                        gr.DrawIcon(n.TrayItem.Icon, n.Rect)
                    Catch ex As Exception

                    End Try
                End If
            Next
        End SyncLock

        gr.Dispose()
        mask = b
        e.Graphics.DrawImage(b, New Point(0, 0))
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseMove(e)

        For Each i As NotifyIconData In _icons
            If i.Rect.Contains(e.Location) Then
                If Me.tltTip.GetToolTip(Me) <> i.TrayItem.Tooltip Then Me.tltTip.SetToolTip(Me, i.TrayItem.Tooltip)
                i.TrayItem.SendMouseMove()
                Exit Sub
            End If
        Next

        'The mouse was not on an icon
        Me.tltTip.SetToolTip(Me, "")
    End Sub

    Protected Overrides Sub OnMouseClick(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseClick(e)
        For Each n As NotifyIconData In _icons
            If n.Rect.Contains(e.Location) Then
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    n.TrayItem.SendContextMenu()
                ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
                    n.TrayItem.SendLeftMouseClick()
                End If
                Exit Sub
            End If
        Next
    End Sub

    Protected Overrides Sub OnMouseDoubleClick(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDoubleClick(e)
        For Each n As NotifyIconData In _icons
            If n.Rect.Contains(e.Location) Then
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    n.TrayItem.SendRightMouseDoubleClick()
                ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
                    n.TrayItem.SendLeftMouseDoubleClick()
                End If
                Exit Sub
            End If
        Next
    End Sub

#End Region

    Private Sub IconAdded(ByVal item As TrayNotifyIcon)
        SyncLock _icons
            _icons.Add(New NotifyIconData(item))
        End SyncLock
        Refresh()
    End Sub

    Private Delegate Sub RefreshDelegate()

    Public Overrides Sub Refresh() Implements IShellControl.Refresh
        If Me.InvokeRequired Then
            Me.BeginInvoke(New RefreshDelegate(AddressOf Refresh))
        Else
            If Not _updating Then MyBase.Refresh()
        End If
    End Sub

    Private Sub IconModified(ByVal item As TrayNotifyIcon)
        Refresh()
    End Sub

    Private Sub IconRemoved(ByVal item As TrayNotifyIcon)
        SyncLock _icons
            Dim ind As Integer
            For Each i As NotifyIconData In _icons
                If i.TrayItem = item Then
                    ind = _icons.IndexOf(i)
                    Exit For
                End If
            Next
            _icons.RemoveAt(ind)
        End SyncLock
        Refresh()
    End Sub

    Public Property Config() As ShellControlConfig Implements IShellControl.Config
        Get
            Return ShellControlHelper.GetConfig(Of NotifyIconBarConfig)(Me)
        End Get
        Set(ByVal value As ShellControlConfig)
            ShellControlHelper.SetConfig(Me, value)
        End Set
    End Property

    Dim mask As Bitmap

    Public Function GetAlphaMask() As System.Drawing.Bitmap Implements IAlphaPaint.GetAlphaMask
        If mask Is Nothing Then
            Refresh()
            Return GetAlphaMask()
        Else
            Return mask
        End If
    End Function

    Dim _updating As Boolean

    Public Sub BeginUpdate() Implements IShellControl.BeginUpdate
        _updating = True
    End Sub

    Public Sub EndUpdate() Implements IShellControl.EndUpdate
        _updating = False
    End Sub

End Class

Public Class NotifyIconBarConfig
    Inherits ShellControlConfig

    Dim _horizSpacing As Integer

    Public Property HorizontalSpacing() As Integer
        Get
            Return _horizSpacing
        End Get
        Set(ByVal value As Integer)
            _horizSpacing = value
        End Set
    End Property

    Dim _vertSpacing As Integer

    Public Property VerticalSpacing() As Integer
        Get
            Return _vertSpacing
        End Get
        Set(ByVal value As Integer)
            _vertSpacing = value
        End Set
    End Property

End Class