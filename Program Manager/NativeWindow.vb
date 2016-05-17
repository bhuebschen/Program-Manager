Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class NativeWindowEx
    Inherits NativeWindow

    Public Delegate Function WndProcDelegate(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    Private Declare Auto Function RegisterClassEx Lib "user32" (ByRef pcWndClassEx As WNDCLASSEX) As Integer
    Private Declare Auto Function UnregisterClassAPI Lib "user32.dll" Alias "UnregisterClass" (ByVal lpClassname As String, ByVal hInstance As IntPtr) As Boolean
    Private Declare Auto Function GetModuleHandle Lib "kernel32.dll" (ByVal filename As String) As IntPtr

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Private Structure WNDCLASSEX
        Public cbSize As Integer ' Size in bytes of the WNDCLASSEX structure
        Public style As Integer ' Class style
        Public lpfnWndProc As WndProcDelegate ' Pointer to the classes Window Procedure
        Public cbClsExtra As Integer ' Number of extra bytes to allocate for class
        Public cbWndExtra As Integer ' Number of extra bytes to allocate for window
        Public hInstance As IntPtr ' Applications instance handle Class
        Public hIcon As IntPtr ' Handle to the classes icon
        Public hCursor As IntPtr ' Handle to the classes cursor
        Public hbrBackground As IntPtr ' Handle to the classes background brush
        Public lpszMenuName As String ' Resource name of class menu
        Public lpszClassName As String ' Name of the Window Class
        Public hIconSm As IntPtr ' Handle to the classes small icon
    End Structure

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        MyBase.WndProc(m)
        RaiseEvent MessageRecieved(m)
    End Sub

    Public Event MessageRecieved(ByRef m As Message)

    Public Overridable Sub RegisterClass(ByVal ccp As ClassCreateParams)
        Dim wc As New WNDCLASSEX
        wc.cbSize = Marshal.SizeOf(GetType(WNDCLASSEX))
        wc.lpszClassName = ccp.Name
        wc.lpfnWndProc = ccp.WndProc
        wc.style = ccp.Style
        wc.cbClsExtra = ccp.ExtraClassBytes
        wc.cbWndExtra = ccp.ExtraWindowBytes
        wc.hbrBackground = ccp.BackgroundBrush
        If ccp.Cursor IsNot Nothing Then wc.hCursor = ccp.Cursor.Handle
        If ccp.Icon IsNot Nothing Then wc.hIcon = ccp.Icon.Handle
        If ccp.SmallIcon IsNot Nothing Then wc.hIconSm = ccp.SmallIcon.Handle
        wc.hInstance = Marshal.GetHINSTANCE(GetType(NativeWindowEx).Module)
        If RegisterClassEx(wc) = False Then
            Throw New System.ComponentModel.Win32Exception
        End If
    End Sub

    Public Overridable Sub UnRegisterClass(ByVal name As String)
        If UnregisterClassAPI(name, GetModuleHandle(Nothing)) = False Then
            Throw New Win32Exception
        End If
    End Sub

    Private Declare Auto Function CreateWindowEx Lib "user32.dll" (ByVal dwExStyle As Integer, _
ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, _
ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As IntPtr, _
ByVal hMenu As IntPtr, ByVal hInstance As IntPtr, ByVal lpParam As IntPtr) As IntPtr

    Public Overrides Sub CreateHandle(ByVal cp As System.Windows.Forms.CreateParams)
        MyBase.CreateHandle(cp)
    End Sub

    Public Overloads Sub CreateHandle(ByVal cp As CreateParamsEx)
        Me.AssignHandle(CreateWindowEx(cp.ExStyle, cp.ClassName, cp.Caption, cp.Style, cp.X, cp.Y, cp.Width, cp.Height, cp.Parent, 0, cp.hInstance, cp.Param))
    End Sub

End Class

Public Class CreateParamsEx
    Inherits CreateParams

    Dim _hInstance As IntPtr = Marshal.GetHINSTANCE(GetType(NativeWindowEx).Module)

    Public Property hInstance() As IntPtr
        Get
            Return _hInstance
        End Get
        Set(ByVal value As IntPtr)
            _hInstance = value
        End Set
    End Property

End Class

Public Class ClassCreateParams

    Dim _style As Integer

    Public Property Style() As ClassStyles
        Get
            Return _style
        End Get
        Set(ByVal value As ClassStyles)
            _style = value
        End Set
    End Property

    Dim _name As String

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Dim _wndproc As NativeWindowEx.WndProcDelegate

    Public Property WndProc() As NativeWindowEx.WndProcDelegate
        Get
            Return _wndproc
        End Get
        Set(ByVal value As NativeWindowEx.WndProcDelegate)
            _wndproc = value
        End Set
    End Property

    Dim _bckBrush As IntPtr

    Public Property BackgroundBrush() As IntPtr
        Get
            Return _bckbrush
        End Get
        Set(ByVal value As IntPtr)
            _bckBrush = value
        End Set
    End Property

    Dim _ico As Icon

    Public Property Icon() As Icon
        Get
            Return _ico
        End Get
        Set(ByVal value As Icon)
            _ico = value
        End Set
    End Property

    Dim _smico As Icon

    Public Property SmallIcon() As Icon
        Get
            Return _smico
        End Get
        Set(ByVal value As Icon)
            _smico = value
        End Set
    End Property

    Dim _curs As Cursor

    Public Property Cursor() As Cursor
        Get
            Return _curs
        End Get
        Set(ByVal value As Cursor)
            _curs = value
        End Set
    End Property

    Dim _clsBytes As Integer

    Public Property ExtraClassBytes() As Integer
        Get
            Return _clsBytes
        End Get
        Set(ByVal value As Integer)
            _clsBytes = value
        End Set
    End Property

    Dim _wndBytes As Integer

    Public Property ExtraWindowBytes() As Integer
        Get
            Return _wndBytes
        End Get
        Set(ByVal value As Integer)
            _wndBytes = value
        End Set
    End Property

    <Flags()> _
    Public Enum ClassStyles As Integer
        VRedraw = &H1
        HRedraw = &H2
        DoubleClicks = &H8
        OwnDC = &H20
        ClassDC = &H40
        ParentDC = &H80
        NoClose = &H200
        SaveBits = &H800
        ByteAlignClient = &H1000
        ByteAlignWindow = &H2000
        GlobalClass = &H4000
        IME = &H10000
        'Windows XP and above
        DropShadow = &H20000
    End Enum

End Class

Public Enum ExtendedWindowStyles As Integer
    WS_EX_DLGMODALFRAME = &H1
    WS_EX_NOPARENTNOTIFY = &H4
    WS_EX_TOPMOST = &H8L
    WS_EX_ACCEPTFILES = &H10
    WS_EX_TRANSPARENT = &H20
    WS_EX_MDICHILD = &H40
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    WS_EX_LAYERED = &H80000
    WS_EX_NOINHERITLAYOUT = &H100000
    WS_EX_LAYOUTRTL = &H400000
    WS_EX_COMPOSITED = &H2000000
    WS_EX_NOACTIVATE = &H8000000
End Enum

Public Enum WindowStyles As Integer
    WS_OVERLAPPED = &H0
    WS_POPUP = &H80000000
    WS_CHILD = &H40000000
    WS_MINIMIZE = &H20000000
    WS_VISIBLE = &H10000000
    WS_DISABLED = &H8000000
    WS_CLIPSIBLINGS = &H4000000
    WS_CLIPCHILDREN = &H2000000
    WS_MAXIMIZE = &H1000000
    WS_CAPTION = (WS_BORDER Or WS_DLGFRAME)
    WS_BORDER = &H800000
    WS_DLGFRAME = &H400000
    WS_VSCROLL = &H200000
    WS_HSCROLL = &H100000
    WS_SYSMENU = &H80000
    WS_THICKFRAME = &H40000
    WS_GROUP = &H20000
    WS_TABSTOP = &H10000
    WS_MINIMIZEBOX = &H20000
    WS_MAXIMIZEBOX = &H10000
    WS_TILED = WS_OVERLAPPED
    WS_ICONIC = WS_MINIMIZE
    WS_SIZEBOX = WS_THICKFRAME
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    WS_CHILDWINDOW = (WS_CHILD)
End Enum
