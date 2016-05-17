Imports System.Data
Imports System.Runtime.InteropServices

''' <summary>
''' A form that is docked to a side of the screen.
''' </summary>
Public Class AppbarForm
    Inherits System.Windows.Forms.Form

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.Container = Nothing

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If (disposing) Then
            If (components IsNot Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "APPBAR"

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure APPBARDATA
        Public cbSize As Integer
        Public hWnd As IntPtr
        Public uCallbackMessage As Integer
        Public uEdge As Integer
        Public rc As RECT
        Public lParam As IntPtr
    End Structure

    Private Enum ABMsg As Integer
        ABM_NEW = 0
        ABM_REMOVE = 1
        ABM_QUERYPOS = 2
        ABM_SETPOS = 3
        ABM_GETSTATE = 4
        ABM_GETTASKBARPOS = 5
        ABM_ACTIVATE = 6
        ABM_GETAUTOHIDEBAR = 7
        ABM_SETAUTOHIDEBAR = 8
        ABM_WINDOWPOSCHANGED = 9
        ABM_SETSTATE = 10
    End Enum

    Private Enum ABNotify As Integer
        ABN_STATECHANGE = 0
        ABN_POSCHANGED
        ABN_FULLSCREENAPP
        ABN_WINDOWARRANGE
    End Enum

    ''' <summary>
    ''' Represents the edge on which the form is docked to.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ABEdge As Integer
        Left = 0
        Top
        Right
        Bottom
    End Enum

    Dim _currentABEdge As ABEdge = ABEdge.Top

    ''' <summary>
    ''' Gets or sets the side of the screen on which the form is docked.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property Edge() As ABEdge
        Get
            Return _currentABEdge
        End Get
        Set(ByVal value As ABEdge)
            _currentABEdge = value
            If Not Me.DesignMode Then ABSetPos()
        End Set
    End Property

    Private fBarRegistered As Boolean = False

    Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Integer, ByRef pData As APPBARDATA) As Integer
    Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal Index As Integer) As Integer
    Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal repaint As Boolean) As Boolean
    Private Declare Auto Function RegisterWindowMessage Lib "user32.dll" (ByVal msg As String) As Integer
    Private uCallBack As Integer

    Private Sub RegisterBar()
        Dim abd As New APPBARDATA
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        If (Not fBarRegistered) Then
            uCallBack = RegisterWindowMessage("AppBarMessage")
            abd.uCallbackMessage = uCallBack

            Dim ret As Integer = SHAppBarMessage(CInt(ABMsg.ABM_NEW), abd)
            fBarRegistered = True

            ABSetPos()
        Else
            SHAppBarMessage(CInt(ABMsg.ABM_REMOVE), abd)
            fBarRegistered = False
        End If
    End Sub

    Private Sub ABSetPos()
        Dim abd As New APPBARDATA
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        abd.uEdge = CInt(_currentABEdge)

        If (abd.uEdge = CInt(ABEdge.Left) OrElse abd.uEdge = CInt(ABEdge.Right)) Then
            abd.rc.top = 0
            abd.rc.bottom = SystemInformation.PrimaryMonitorSize.Height
            If (abd.uEdge = CInt(ABEdge.Left)) Then
                abd.rc.left = 0
                abd.rc.right = Size.Width
            Else
                abd.rc.right = SystemInformation.PrimaryMonitorSize.Width
                abd.rc.left = abd.rc.right - Size.Width
            End If
        Else
            abd.rc.left = 0
            abd.rc.right = SystemInformation.PrimaryMonitorSize.Width
            If (abd.uEdge = CInt(ABEdge.Top)) Then
                abd.rc.top = 0
                abd.rc.bottom = Size.Height
            Else
                abd.rc.bottom = SystemInformation.PrimaryMonitorSize.Height
                abd.rc.top = abd.rc.bottom - Size.Height
            End If
        End If

        ' Query the system for an approved size and position.
        SHAppBarMessage(CInt(ABMsg.ABM_QUERYPOS), abd)

        ' Adjust the rectangle, depending on the edge to which the
        ' appbar is anchored.
        Select Case (abd.uEdge)
            Case ABEdge.Left
                abd.rc.right = abd.rc.left + Size.Width
            Case ABEdge.Right
                abd.rc.left = abd.rc.right - Size.Width
            Case ABEdge.Top
                abd.rc.bottom = abd.rc.top + Size.Height
            Case ABEdge.Bottom
                abd.rc.top = abd.rc.bottom - Size.Height
        End Select

        ' Pass the final bounding rectangle to the system.
        SHAppBarMessage(ABMsg.ABM_SETPOS, abd)

        ' Move and size the appbar so that it conforms to the
        ' bounding rectangle passed to the system.
        MoveWindow(abd.hWnd, abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top, True)
    End Sub

    Private Sub ABActivate()
        Dim abd As New APPBARDATA
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        SHAppBarMessage(ABMsg.ABM_ACTIVATE, abd)
    End Sub

    Private Sub ABWindowPosChange()
        Dim abd As New APPBARDATA
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        SHAppBarMessage(ABMsg.ABM_WINDOWPOSCHANGED, abd)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If (m.Msg = uCallBack) Then
            Select Case (m.WParam.ToInt64())
                Case ABNotify.ABN_POSCHANGED
                    If Not Me.DesignMode Then ABSetPos()
                Case WindowInfo.WindowMessages.WM_ACTIVATE
                    ABActivate()
                Case WindowInfo.WindowMessages.WM_WINDOWPOSCHANGED
                    ABWindowPosChange()
                Case ABNotify.ABN_FULLSCREENAPP
                    If m.LParam = True Then
                        Me.Hide()
                    Else
                        Me.Show()
                    End If
                Case ABNotify.ABN_WINDOWARRANGE
                    If m.LParam = True Then
                        Me.Hide()
                    Else
                        Me.Show()
                    End If
            End Select
        End If

        Try
            MyBase.WndProc(m)
        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            If Not Me.DesignMode Then
                cp.Style = cp.Style And (Not &HC00000) ' WS_CAPTION
                cp.Style = cp.Style And (Not &H800000) ' WS_BORDER
                cp.ExStyle = &H80 Or &H8 ' WS_EX_TOOLWINDOW | WS_EX_TOPMOST
            End If
            Return cp
        End Get
    End Property

#End Region

    Private Sub Me_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not Me.DesignMode Then RegisterBar()
    End Sub

    Private Sub Me_Closing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        If Not Me.DesignMode Then RegisterBar()
    End Sub

    Protected Overrides Sub OnResizeEnd(ByVal e As System.EventArgs)
        If Not Me.DesignMode Then ABSetPos()
        MyBase.OnResize(e)
    End Sub

End Class