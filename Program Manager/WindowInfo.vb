''' <summary>
''' A full featured window wrapper class.
''' </summary>
''' <remarks></remarks>
Public Class WindowInfo
    Implements IWin32Window

    Private Declare Auto Function GetWindowText Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal text As System.Text.StringBuilder, ByVal maxLength As Integer) As Integer
    Private Declare Auto Function GetWindowTextLength Lib "user32.dll" (ByVal hwnd As IntPtr) As Integer
    Private Declare Auto Function SetWindowText Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal newText As String) As Boolean
    Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef rgn As IntPtr) As Integer
    Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal rgn As IntPtr, ByVal redraw As Boolean) As Integer
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef procId As Integer) As Integer
    Private Declare Auto Function SendMessageAPI Lib "user32.dll" Alias "SendMessage" (ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer) As Integer
    Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef rect As RECT) As Boolean
    Private Declare Auto Function FindWindowAPI Lib "user32.dll" Alias "FindWindow" (ByVal className As String, ByVal windowTitle As String) As IntPtr
    Private Declare Function EnumWindows Lib "user32.dll" (ByVal callback As ListWindowDelegate, ByVal lparam As Integer) As Boolean
    Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal wCmd As Integer) As IntPtr
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Integer
    Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal fEnable As Integer) As Integer
    Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal refresh As Boolean) As Boolean
    Private Declare Function GetDesktopWindowAPI Lib "user32.dll" Alias "GetDesktopWindow" () As IntPtr
    Private Declare Function AnimateWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal dwTime As Integer, ByVal dwFlags As Integer) As Boolean
    Private Declare Function ChildWindowFromPointEx Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal p As PT, ByVal flags As Integer) As IntPtr
    Private Declare Function EndTaskAPI Lib "user32.dll" Alias "EndTask" (ByVal hwnd As IntPtr, ByVal shutdown As Boolean, ByVal force As Boolean) As Boolean
    Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal callback As ListWindowDelegate, ByVal lparam As Integer) As Boolean
    Private Declare Function FindWindowEx Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal afterChild As IntPtr, ByVal className As String, ByVal text As String) As IntPtr
    Private Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal flags As Integer) As IntPtr
    Private Declare Function GetParentAPI Lib "user32.dll" Alias "GetParent" (ByVal hwnd As IntPtr) As IntPtr
    Private Declare Function GetForegroundWindowAPI Lib "user32.dll" Alias "GetForegroundWindow" () As IntPtr
    Private Declare Function IsChildAPI Lib "user32.dll" Alias "IsChild" (ByVal hwndParent As IntPtr, ByVal hwnd As IntPtr) As Boolean
    Private Declare Function IsHungAppWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function GetWindowInfo Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef info As WININFO) As Boolean
    Private Declare Function SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Integer) As Boolean
    Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal hwndAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Integer) As Boolean
    Private Declare Auto Function GetClassName Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal className As System.Text.StringBuilder, ByVal buffersize As Integer) As Integer
    Private Declare Auto Function GetWindowModuleFileName Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef lpszFileName As System.Text.StringBuilder, ByVal cchFileNameMax As Integer) As Integer
    Private Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As IntPtr, ByRef lpRect As RECT, ByVal bErase As Boolean) As Boolean
    Private Declare Auto Function GetWindowLong Lib "user32.dll" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Auto Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As _SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As Integer
    Private Declare Auto Function PostMessageAPI Lib "user32.dll" Alias "PostMessage" (ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer) As Integer
    Private Declare Auto Function GetClassLong Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Private Declare Auto Function SendMessageTimeoutAPI Lib "user32.dll" Alias "SendMessageTimeout" (ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer, ByVal uflags As Integer, ByVal uTimeout As Integer, ByRef returnVal As Integer) As Integer
    Private Declare Auto Sub SwitchToThisWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal fAltTab As Boolean)
    Private Declare Auto Function DestroyWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean

#Region "Shared Methods"

    Public Overloads Shared Function FindWindowByClass(ByVal className As String) As WindowInfo
        Return FindWindow(className, Nothing)
    End Function

    Public Overloads Shared Function FindWindowByText(ByVal windowText As String) As WindowInfo
        Return FindWindow(Nothing, windowText)
    End Function

    Public Overloads Shared Function FindWindow(ByVal classname As String, ByVal windowText As String) As WindowInfo
        Return New WindowInfo(FindWindowAPI(classname, windowText))
    End Function

    Private Delegate Function ListWindowDelegate(ByVal hwnd As IntPtr, ByVal lparam As Integer) As Boolean

    Private Shared infos As ObjectModel.Collection(Of WindowInfo)

    Public Shared Function GetWindows() As ObjectModel.Collection(Of WindowInfo)
        infos = New ObjectModel.Collection(Of WindowInfo)
        If EnumWindows(AddressOf ListWindow, 0) = False Then
            Throw New Exception("An error occurred while enumerating windows.")
        End If
        Return infos
    End Function

    Public Delegate Function WindowChooseDelegate(ByVal w As WindowInfo) As Boolean
    Shared chooser As WindowChooseDelegate

    Public Shared Function GetWindows(ByVal windowChoose As WindowChooseDelegate) As ObjectModel.Collection(Of WindowInfo)
        chooser = windowChoose
        Dim r As Object = GetWindows()
        chooser = Nothing
        Return r
    End Function

    Private Shared Function ListWindow(ByVal hwnd As IntPtr, ByVal lparam As Integer) As Boolean
        If IsValid(hwnd) Then
            If chooser Is Nothing OrElse chooser(New WindowInfo(hwnd)) Then
                infos.Add(New WindowInfo(hwnd))
            End If
        End If
        Return True
    End Function

    Public Shared Function GetDesktopWindow() As WindowInfo
        Return New WindowInfo(GetDesktopWindowAPI)
    End Function

    Public Shared Function GetForegroundWindow() As WindowInfo
        Return New WindowInfo(GetForegroundWindowAPI())
    End Function

    Public Shared Function IsValid(ByVal w As WindowInfo) As Boolean
        Return IsWindow(w.Handle)
    End Function

    Public Shared Function IsValid(ByVal h As IntPtr) As Boolean
        Return IsWindow(h)
    End Function

    Public Shared Operator =(ByVal value1 As WindowInfo, ByVal value2 As WindowInfo) As Boolean
        If value1 Is Nothing And value2 Is Nothing Then
            Return True
        ElseIf value1 Is Nothing Or value2 Is Nothing Then
            Return False
        Else
            Return (value1.Handle = value2.Handle)
        End If
    End Operator

    Public Shared Operator <>(ByVal value1 As WindowInfo, ByVal value2 As WindowInfo) As Boolean
        Return Not (value1 = value2)
    End Operator

#End Region

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> _
    Private Structure PT
        Public x As Integer
        Public y As Integer
    End Structure

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> _
    Private Structure _SHFILEINFO
        Public hIcon As Integer
        Public iIcon As Integer
        Public dwAttributes As Integer
        <VBFixedString(260)> Public szDisplayName As String
        <VBFixedString(80)> Public szTypeName As String
    End Structure

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> _
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    Dim hWnd As IntPtr

    ''' <summary>
    ''' Creates a new instance of the WindowInfo class using the supplied window handle.
    ''' </summary>
    ''' <param name="handle"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal handle As IntPtr)
        If IsWindow(handle) Then
            hWnd = handle
        Else
            Throw New Exception("The specified window handle does not represent any existing windows.")
        End If
    End Sub

    ''' <summary>
    ''' Closes and destroys the window.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Close()
        If DestroyWindow(Handle) = False Then Throw New System.ComponentModel.Win32Exception
    End Sub

    ''' <summary>
    ''' Returns the class name of the window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ClassName() As String
        Get
            Dim s As New System.Text.StringBuilder(256)
            If GetClassName(hWnd, s, s.Capacity) = 0 Then
                Throw New Exception("Error retrieving window class name.")
            End If
            Return s.ToString
        End Get
    End Property

    ''' <summary>
    ''' Returns the full filename and path of the module associated with the window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Fullname() As String
        Get
            Dim modulePathname As New System.Text.StringBuilder(1024)
            Dim length As Integer = GetWindowModuleFileName(hWnd, modulePathname, modulePathname.Capacity)
            Return modulePathname.ToString(0, length)
        End Get
    End Property

    Private Const WM_GETICON As Integer = &H7F
    Private Const WM_SETICON As Integer = &H80

    ''' <summary>
    ''' Represents an icon size for a window.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum WindowIconSize As Integer
        ''' <summary>
        ''' A small icon used for the window caption.
        ''' </summary>
        ''' <remarks></remarks>
        Small = 0
        ''' <summary>
        ''' A big icon used for Alt-Tab.
        ''' </summary>
        ''' <remarks></remarks>
        Big = 1
        ''' <summary>
        ''' Returns either the small icon, if available, or the system generated one.
        ''' </summary>
        ''' <remarks></remarks>
        Small2 = 2
    End Enum

    Private Const GWL_EXSTYLE As Integer = (-20)
    Private Const WS_EX_APPWINDOW As Integer = &H40000
    Private Const WS_EX_TOOLWINDOW As Integer = &H80

    Private Const GW_OWNER As Integer = 4

    Public ReadOnly Property ShowInTaskbar() As Boolean
        Get
            If (Me.Visible) And (GetParentAPI(hWnd) = 0) Then
                ' extended options call up
                Dim exStyles As Integer = GetWindowLong(hWnd, GWL_EXSTYLE)
                ' parents windows examine again:
                Dim ownerWin As IntPtr = GetWindow(hWnd, GW_OWNER)
                ' examining whether:
                ' - no ToolWindow and no Childfenster or
                ' - application windows and Childfenster
                If (((exStyles And WS_EX_TOOLWINDOW) = 0) And (ownerWin = 0)) Or ((exStyles And WS_EX_APPWINDOW) And (ownerWin <> 0)) Then
                    Return True
                End If
            End If
            Return False
        End Get
    End Property

    Private Const SHGFI_ICON As Integer = &H100     ' get icon
    Private Const SHGFI_DISPLAYNAME As Integer = &H200     ' get display name
    Private Const SHGFI_TYPENAME As Integer = &H400     ' get type name
    Private Const SHGFI_ATTRIBUTES As Integer = &H800     ' get attributes
    Private Const SHGFI_ICONLOCATION As Integer = &H1000     ' get icon location
    Private Const SHGFI_EXETYPE As Integer = &H2000     ' return exe type
    Private Const SHGFI_SYSICONINDEX As Integer = &H4000     ' get system icon index
    Private Const SHGFI_LINKOVERLAY As Integer = &H8000     ' put a link overlay on icon
    Private Const SHGFI_SELECTED As Integer = &H10000     ' show icon in selected state

    'Windows 2000+
    Private Const SHGFI_ATTR_SPECIFIED As Integer = &H20000     ' get only specified attributes

    Private Const SHGFI_LARGEICON As Integer = &H0     ' get large icon
    Private Const SHGFI_SMALLICON As Integer = &H1     ' get small icon
    Private Const SHGFI_OPENICON As Integer = &H2     ' get open icon
    Private Const SHGFI_SHELLICONSIZE As Integer = &H4     ' get shell size icon
    Private Const SHGFI_PIDL As Integer = &H8     ' pszPath is a pidl
    Private Const SHGFI_USEFILEATTRIBUTES As Integer = &H10     ' use passed dwFileAttribute
    'IE 5+
    Private Const SHGFI_ADDOVERLAYS As Integer = &H20     ' apply the appropriate overlays
    Private Const SHGFI_OVERLAYINDEX As Integer = &H40     ' Get the index of the overlay
    ' in the upper 8 bits of the iIcon

    Private Const GCL_CBCLSEXTRA = -20
    Private Const GCL_CBWNDEXTRA = -18
    Private Const GCL_HBRBACKGROUND = -10
    Private Const GCL_HCURSOR = -12
    Private Const GCL_HICON = -14
    Private Const GCL_HMODULE = -16
    Private Const GCL_MENUNAME = -8
    Private Const GCL_STYLE = -26
    Private Const GCL_WNDPROC = -24
    Private Const GCL_HICONSM As Long = -34

    ''' <summary>
    ''' Gets or sets the icon with the specified size for the window.
    ''' </summary>
    ''' <param name="iconsize"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If the icon has not been set for the window, then it retrieves the file's icon.</remarks>
    Public Property Icon(ByVal iconsize As WindowIconSize) As Icon
        Get
            Try
                Dim hIco As Integer = SendMessageTimeout(WM_GETICON, CInt(iconsize), 0, SendMessageTimeoutFlags.AbortIfHung, 200)
                If hIco = 0 Then
                    hIco = GetClassLong(hWnd, IIf(iconsize = WindowIconSize.Big, GCL_HICON, GCL_HICONSM))
                    If hIco = 0 Then
                        Dim shfi As New _SHFILEINFO
                        Dim fileName As String = Me.Process.MainModule.FileName
                        If Not fileName.Length > 260 Then
                            SHGetFileInfo(fileName, 0, shfi, Runtime.InteropServices.Marshal.SizeOf(shfi), SHGFI_ICON Or IIf(iconsize = WindowIconSize.Big, SHGFI_LARGEICON, SHGFI_SMALLICON))
                            hIco = shfi.hIcon
                        End If
                    End If
                End If
                Return Drawing.Icon.FromHandle(hIco)
            Catch ex As Exception
                Debug.Print(ex.Message)
                Return Nothing
            End Try
        End Get
        Set(ByVal value As Icon)
            PostMessage(WM_SETICON, CInt(IIf(iconsize = WindowIconSize.Small2, WindowIconSize.Small, iconsize)), value.Handle)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the window caption.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Text() As String
        Get
            Dim length As Integer = GetWindowTextLength(hWnd)
            Dim t As New System.Text.StringBuilder(length)
            GetWindowText(hWnd, t, length + 1)
            Return t.ToString
        End Get
        Set(ByVal value As String)
            If SetWindowText(hWnd, value) = False Then
                Throw New Exception("An error occured setting the window text.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the clip region of the window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Region() As Region
        Get
            Dim r As IntPtr
            GetWindowRgn(hWnd, r)
            Return Drawing.Region.FromHrgn(r)
        End Get
        Set(ByVal value As Region)
            SetWindowRgn(hWnd, value.GetHrgn(Nothing), True)
        End Set
    End Property

    ''' <summary>
    ''' Returns the process object that owns this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Process() As Process
        Get
            Return Diagnostics.Process.GetProcessById(ProcessId)
        End Get
    End Property

    ''' <summary>
    ''' Returns the ID of the process that owns this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProcessId() As Integer
        Get
            Dim procId As Integer
            Dim threadId As Integer = GetWindowThreadProcessId(hWnd, procId)
            Return procId
        End Get
    End Property

    ''' <summary>
    ''' Returns the thread object that owns this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Thread() As ProcessThread
        Get
            Dim procId As Integer
            Dim threadId As Integer = GetWindowThreadProcessId(hWnd, procId)
            Return Diagnostics.Process.GetProcessById(procId).Threads(threadId)
        End Get
    End Property

    ''' <summary>
    ''' Returns the ID of the thread that owns this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ThreadId() As Integer
        Get
            Dim procId As Integer
            Return GetWindowThreadProcessId(hWnd, procId)
        End Get
    End Property

    ''' <summary>
    ''' Returns the handle used to represent this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Handle() As IntPtr Implements IWin32Window.Handle
        Get
            Return hWnd
        End Get
    End Property

    ''' <summary>
    ''' Sends a message to the window for processing.
    ''' </summary>
    ''' <param name="m"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendMessage(ByRef m As Message) As Integer
        m.Result = SendMessage(m.Msg, m.WParam, m.LParam)
        Return m.Result
    End Function

    ''' <summary>
    ''' Sends a message to the window for processing.
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="wparam"></param>
    ''' <param name="lparam"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendMessage(ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer) As Integer
        Return SendMessageAPI(hWnd, msg, wparam, lparam)
    End Function

    ''' <summary>
    ''' Instead of sending a message and waiting for it to be processed, posts a message on the window's message queue and immediately returns.
    ''' </summary>
    ''' <param name="m"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PostMessage(ByRef m As Message) As Integer
        m.Result = PostMessageAPI(m.HWnd, m.Msg, m.WParam, m.LParam)
        Return m.Result
    End Function

    ''' <summary>
    ''' Instead of sending a message and waiting for it to be processed, posts a message on the window's message queue and immediately returns.
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="wparam"></param>
    ''' <param name="lparam"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PostMessage(ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer) As Integer
        Return PostMessageAPI(hWnd, msg, wparam, lparam)
    End Function

    Public Enum SendMessageTimeoutFlags As Integer
        Normal = &H0
        Block = &H1
        AbortIfHung = &H2
        ''' <summary>
        '''
        ''' </summary>
        ''' <remarks>Windows 2000 and later only.</remarks>
        NoTimeoutIfNotHung = &H8
        ''' <summary>
        '''
        ''' </summary>
        ''' <remarks>Windows Vista and later only.</remarks>
        ErrorOnExit = &H20
    End Enum

    ''' <summary>
    ''' Sends a message and if the application does not process the message in the specified timeout period, returns.
    ''' </summary>
    ''' <param name="m"></param>
    ''' <param name="flags"></param>
    ''' <param name="timeout">The timeout period in milliseconds. Note that if the message is being sent to multiple windows, the maximum timeout period is the number of windows times this.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendMessageTimeout(ByRef m As Message, ByVal flags As SendMessageTimeoutFlags, ByVal timeout As Integer) As Integer
        m.Result = SendMessageTimeout(m.Msg, m.WParam, m.LParam, flags, timeout)
        Return m.Result
    End Function

    ''' <summary>
    ''' Sends a message and if the application does not process the message in the specified timeout period, returns.
    ''' </summary>
    ''' <param name="flags"></param>
    ''' <param name="timeout">The timeout period in milliseconds. Note that if the message is being sent to multiple windows, the maximum timeout period is the number of windows times this.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendMessageTimeout(ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer, ByVal flags As SendMessageTimeoutFlags, ByVal timeout As Integer) As Integer
        Dim result As Integer
        SendMessageTimeoutAPI(hWnd, msg, wparam, lparam, flags, timeout, result)
        Return result
    End Function

    ''' <summary>
    ''' Returns or resizes the rectangle of the full window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Rectangle() As Rectangle
        Get
            Dim r As RECT
            If GetWindowRect(hWnd, r) = False Then
                Throw New Exception("Error retrieving window rectangle.")
            End If
            Return Drawing.Rectangle.FromLTRB(r.left, r.top, r.right, r.bottom)
        End Get
        Set(ByVal value As Rectangle)
            Move(value, True)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the width of this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Width() As Integer
        Get
            Return Me.Rectangle.Width
        End Get
        Set(ByVal value As Integer)
            Me.Rectangle = New Rectangle(Me.Rectangle.X, Me.Rectangle.Y, value, Me.Rectangle.Height)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the height of this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Height() As Integer
        Get
            Return Me.Rectangle.Height
        End Get
        Set(ByVal value As Integer)
            Me.Rectangle = New Rectangle(Me.Rectangle.X, Me.Rectangle.Y, Me.Rectangle.Width, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the size of this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Size() As Size
        Get
            Return Me.Rectangle.Size
        End Get
        Set(ByVal value As Size)
            Me.Rectangle = New Rectangle(Me.Rectangle.X, Me.Rectangle.Y, value.Width, value.Height)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the top Y position of this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Top() As Integer
        Get
            Return Me.Rectangle.Top
        End Get
        Set(ByVal value As Integer)
            Me.Rectangle = New Rectangle(Me.Rectangle.X, value, Me.Rectangle.Width, Me.Rectangle.Height)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the left X position of this window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Left() As Integer
        Get
            Return Me.Rectangle.Left
        End Get
        Set(ByVal value As Integer)
            Me.Rectangle = New Rectangle(value, Me.Rectangle.Y, Me.Rectangle.Width, Me.Rectangle.Height)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets this window's location.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Location() As Point
        Get
            Return Me.Rectangle.Location
        End Get
        Set(ByVal value As Point)
            Me.Rectangle = New Rectangle(value.X, value.Y, Me.Rectangle.Width, Me.Rectangle.Height)
        End Set
    End Property

    ''' <summary>
    ''' Returns whether this window is currently visible.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This uses SetWindowState with Hide and Show.</remarks>
    Public Property Visible() As Boolean
        Get
            Return IsWindowVisible(hWnd)
        End Get
        Set(ByVal value As Boolean)
            SetWindowState(IIf(value, WindowState.Show, WindowState.Hide))
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the enabled state of the window.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Enabled() As Boolean
        Get
            Return IsWindowEnabled(hWnd)
        End Get
        Set(ByVal value As Boolean)
            EnableWindow(hWnd, value)
        End Set
    End Property

    ''' <summary>
    ''' Used by SetWindowState and represents the different window states.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum WindowState
        ''' <summary>
        ''' Hides the window.
        ''' </summary>
        ''' <remarks></remarks>
        Hide = 0
        ''' <summary>
        ''' Unmaximizes, unminimizes, and activates the window.
        ''' </summary>
        ''' <remarks></remarks>
        ShowNormal = 1
        ''' <summary>
        ''' Shows, minimizes, and activates the window.
        ''' </summary>
        ''' <remarks></remarks>
        ShowMinimized = 2
        ''' <summary>
        ''' Shows, maximizes and activates the window.
        ''' </summary>
        ''' <remarks></remarks>
        ShowMaximized = 3
        ''' <summary>
        ''' Maximizes the window.
        ''' </summary>
        ''' <remarks></remarks>
        Maximized = 3
        ''' <summary>
        ''' Unmaximizes, unminimizes the window but does not activate it.
        ''' </summary>
        ''' <remarks></remarks>
        NormalNA = 4
        ''' <summary>
        ''' Brings the window back to the state it had previous to being hidden and activates it.
        ''' </summary>
        ''' <remarks></remarks>
        Show = 5
        ''' <summary>
        ''' Minimizes and activates the window.
        ''' </summary>
        ''' <remarks></remarks>
        Minimize = 6
        ''' <summary>
        ''' Shows, minimizes, but does not activate the window.
        ''' </summary>
        ''' <remarks></remarks>
        ShowMinimizeNA = 7
        ''' <summary>
        ''' Shows the window without activating it.
        ''' </summary>
        ''' <remarks></remarks>
        ShowNA = 8
        ''' <summary>
        ''' If the window is minimized, restores the window to the proper state, either maximized or normal.
        ''' </summary>
        ''' <remarks></remarks>
        Restore = 9
        ''' <summary>
        ''' Puts the window in the application's startup default state.
        ''' </summary>
        ''' <remarks></remarks>
        ShowDefault = 10
        ''' <summary>
        ''' Forces the window into a minimized state.
        ''' </summary>
        ''' <remarks></remarks>
        ForceMinimize = 11

    End Enum

    ''' <summary>
    ''' Sets the window state of the window using a WindowState enum.
    ''' </summary>
    ''' <param name="style"></param>
    ''' <remarks></remarks>
    Public Sub SetWindowState(ByVal style As WindowState)
        ShowWindow(hWnd, style)
    End Sub

    ''' <summary>
    ''' Refreshes the window.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Refresh()
        If InvalidateRect(hWnd, Nothing, True) = False Then Throw New System.ComponentModel.Win32Exception
        If UpdateWindow(hWnd) = False Then Throw New Exception("Error updating the window!")
    End Sub

    ''' <summary>
    ''' Moves and resizes the window.
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="shouldRefresh"></param>
    ''' <remarks></remarks>
    Public Sub Move(ByVal r As Rectangle, Optional ByVal shouldRefresh As Boolean = True)
        If MoveWindow(hWnd, r.X, r.Y, r.Width, r.Height, shouldRefresh) = False Then Throw New Exception("Error moving the window.")
    End Sub

    ''' <summary>
    ''' Provides options for the ChildFromPoint function.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ChildFromPointOptions
        All = 0
        SkipInvisible = 1
        SkipDisabled = 2
        SkipTransparent = 4
    End Enum

    ''' <summary>
    ''' Returns the topmost child window at the specified point with the given options.
    ''' </summary>
    ''' <param name="p">The position in client coordinates.</param>
    ''' <param name="options"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChildFromPoint(ByVal p As Point, Optional ByVal options As ChildFromPointOptions = ChildFromPointOptions.All) As WindowInfo
        Dim pt As PT
        pt.x = p.X
        pt.y = p.Y
        Return New WindowInfo(ChildWindowFromPointEx(hWnd, pt, options))
    End Function

    ''' <summary>
    ''' Flags used by GetRelatedWindow.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum WindowRelation As Integer
        ''' <summary>
        ''' The top-most child window.
        ''' </summary>
        ''' <remarks></remarks>
        Child = 5
        ''' <summary>
        ''' The top-most sibling window.
        ''' </summary>
        ''' <remarks></remarks>
        FirstSibling = 0
        ''' <summary>
        ''' The bottom-most sibling window.
        ''' </summary>
        ''' <remarks></remarks>
        LastSibling = 1
        ''' <summary>
        ''' The next sibling window below the current window.
        ''' </summary>
        ''' <remarks></remarks>
        NextSibling = 2
        ''' <summary>
        ''' The next sibling window above the current window.
        ''' </summary>
        ''' <remarks></remarks>
        PreviousSibling = 3
        ''' <summary>
        ''' The owner of the window.
        ''' </summary>
        ''' <remarks></remarks>
        Owner = 4
    End Enum

    ''' <summary>
    ''' Returns the window with the specified relationship to this window.
    ''' </summary>
    ''' <param name="relation"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRelatedWindow(ByVal relation As WindowRelation) As WindowInfo
        Return New WindowInfo(GetWindow(hWnd, relation))
    End Function

    ''' <summary>
    ''' Returns whether or not the application is responding.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NotResponding() As Boolean
        Get
            Return IsHungAppWindow(hWnd)
        End Get
    End Property

    ''' <summary>
    ''' Invokes the more friendly "end task".
    ''' </summary>
    ''' <param name="force"></param>
    ''' <remarks></remarks>
    Public Sub EndTask(ByVal force As Boolean)
        If EndTaskAPI(hWnd, False, force) = False Then Throw New Exception("Error ending the task.")
    End Sub

    Dim children As ObjectModel.Collection(Of WindowInfo)
    Dim childChooser As WindowChooseDelegate

    ''' <summary>
    ''' Returns all child windows.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChildWindows() As ObjectModel.Collection(Of WindowInfo)
        Return GetChildWindows(Nothing)
    End Function

    ''' <summary>
    ''' Returns all child windows who's selection delegate returns true.
    ''' </summary>
    ''' <param name="selector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChildWindows(ByVal selector As WindowChooseDelegate) As ObjectModel.Collection(Of WindowInfo)
        children.Clear()
        childChooser = selector
        If EnumChildWindows(hWnd, AddressOf ListChildWindow, 0) = False Then Throw New Exception("Error enumerating child windows.")
        childChooser = Nothing
        Return children
    End Function

    ''' <summary>
    ''' This function is called for every child window.
    ''' </summary>
    ''' <param name="hwnd"></param>
    ''' <param name="lparam"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ListChildWindow(ByVal hwnd As IntPtr, ByVal lparam As Integer) As Boolean
        If childChooser Is Nothing OrElse childChooser(New WindowInfo(hwnd)) Then
            children.Add(New WindowInfo(hwnd))
        End If
        Return True
    End Function

    ''' <summary>
    ''' Finds the child window with the specified class name and text.
    ''' </summary>
    ''' <param name="className"></param>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindChild(ByVal className As String, ByVal text As String) As WindowInfo
        Return New WindowInfo(FindWindowEx(hWnd, 0, className, text))
    End Function

    ''' <summary>
    ''' Finds the child window with the specified class name.
    ''' </summary>
    ''' <param name="className"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindChildByClass(ByVal className As String) As WindowInfo
        Return FindChild(className, Nothing)
    End Function

    ''' <summary>
    ''' Finds the child window with the specified text.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindChildByText(ByVal text As String) As WindowInfo
        Return FindChild(Nothing, text)
    End Function

    ''' <summary>
    ''' Finds all children with the specified class name and text.
    ''' </summary>
    ''' <param name="className"></param>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindChildren(ByVal className As String, ByVal text As String) As ObjectModel.Collection(Of WindowInfo)
        Dim canContinue As Boolean = True
        Dim children As New ObjectModel.Collection(Of WindowInfo)
        Do
            Dim lastHandle As IntPtr
            If children.Count <> 0 Then lastHandle = children(children.Count - 1).Handle
            Dim h As IntPtr = FindWindowEx(hWnd, lastHandle, className, text)
            If h = 0 Then
                canContinue = False
            Else
                children.Add(New WindowInfo(h))
            End If
        Loop Until canContinue = False
        Return children
    End Function

    ''' <summary>
    ''' Finds all children with the specified class name.
    ''' </summary>
    ''' <param name="className"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindChildrenByClass(ByVal className As String) As ObjectModel.Collection(Of WindowInfo)
        Return FindChildren(className, Nothing)
    End Function

    ''' <summary>
    ''' Finds all children with the specified text.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindChildrenByText(ByVal text As String) As ObjectModel.Collection(Of WindowInfo)
        Return FindChildren(Nothing, text)
    End Function

    Private Const GA_PARENT = 1
    Private Const GA_ROOT = 2
    Private Const GA_ROOTOWNER = 3

    ''' <summary>
    ''' Returns the root window of this window.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRoot() As WindowInfo
        Return New WindowInfo(GetAncestor(hWnd, GA_ROOT))
    End Function

    ''' <summary>
    ''' Returns the root owner window of this window.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRootOwner() As WindowInfo
        Return New WindowInfo(GetAncestor(hWnd, GA_ROOTOWNER))
    End Function

    ''' <summary>
    ''' Returns the parent, not the owner.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetParentNotOwner() As WindowInfo
        Return New WindowInfo(GetAncestor(hWnd, GA_PARENT))
    End Function

    ''' <summary>
    ''' Returns the parent, including the owner.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetParent() As WindowInfo
        Return New WindowInfo(GetParentAPI(hWnd))
    End Function

    ''' <summary>
    ''' Returns whether or not the specified window is a child of this window.
    ''' </summary>
    ''' <param name="parent"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsChild(ByVal parent As WindowInfo) As Boolean
        Return IsChildAPI(parent.Handle, hWnd)
    End Function

    ''' <summary>
    ''' Returns whether or not the specified window is a parent of this window.
    ''' </summary>
    ''' <param name="child"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsParent(ByVal child As WindowInfo) As Boolean
        Return IsChildAPI(hWnd, child.Handle)
    End Function

    ''' <summary>
    ''' Switches to this window using the SwitchToThisWindow API.
    ''' </summary>
    ''' <param name="altTab">Whether of not this window was switched to by alt-tab.</param>
    ''' <remarks></remarks>
    Public Sub SwitchTo(Optional ByVal altTab As Boolean = False)
        SwitchToThisWindow(hWnd, altTab)
    End Sub

    ''' <summary>
    ''' Returns whether this window is minimized.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Minimized() As Boolean
        Get
            Return IsIconic(hWnd)
        End Get
    End Property

    ''' <summary>
    ''' Returns whether this window is maximized.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Maximized() As Boolean
        Get
            Return IsZoomed(hWnd)
        End Get
    End Property

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> _
    Private Structure WININFO
        Public cbSize As Integer
        Public rcWindow As RECT
        Public rcClient As RECT
        Public dwStyle As Integer
        Public dwExStyle As Integer
        Public dwWindowStatus As Integer
        Public cxWindowBorders As UInteger
        Public cyWindowBorders As UInteger
        Public atomWindowType As Short
        Public wCreatorVersion As Short
    End Structure

    ''' <summary>
    ''' Sets this window as the foreground window.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BringToFront()
        If SetForegroundWindow(hWnd) = False Then Throw New Exception("Error bringing window to front.")
    End Sub

    ''' <summary>
    ''' Sets this window as the bottom-most window.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SendToBack()
        If SetWindowPos(hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE) = False Then Throw New Exception("Error sending window to back.")
    End Sub

    ''' <summary>
    ''' Sets this window as a topmost window who can't fall behind not topmost windows.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetAsTopMost()
        If SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE) = False Then Throw New Exception("Error making window topmost.")
    End Sub

    ''' <summary>
    ''' Sets this window as a non-topmost window.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetAsNotTopMost()
        If SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE) = False Then Throw New Exception("Error making window non-topmost.")
    End Sub

    Private Const HWND_TOP As Integer = 0
    Private Const HWND_BOTTOM As Integer = 1
    Private Const HWND_TOPMOST As Integer = -1
    Private Const HWND_NOTOPMOST As Integer = -2

    Private Const SWP_NOSIZE As Integer = &H1
    Private Const SWP_NOMOVE As Integer = &H2
    Private Const SWP_NOZORDER As Integer = &H4
    Private Const SWP_NOREDRAW As Integer = &H8
    Private Const SWP_NOACTIVATE As Integer = &H10
    Private Const SWP_FRAMECHANGED As Integer = &H20  ' The frame changed: send WM_NCCALCSIZE
    Private Const SWP_SHOWWINDOW As Integer = &H40
    Private Const SWP_HIDEWINDOW As Integer = &H80
    Private Const SWP_NOCOPYBITS As Integer = &H100
    Private Const SWP_NOOWNERZORDER As Integer = &H200  ' Don't do owner Z ordering
    Private Const SWP_NOSENDCHANGING As Integer = &H400  ' Don't send WM_WINDOWPOSCHANGING

    Public Const WM_PRINT As Integer = &H317
    Public Const PRF_NONCLIENT As Integer = &H2
    Public Const PRF_CLIENT As Integer = &H4
    Public Const PRF_ERASEBKGND As Integer = &H8
    Public Const PRF_CHILDREN As Integer = &H10
    Public Const PRF_OWNED As Integer = &H20

    ''' <summary>
    ''' Takes a snapshot of a window and its children and returns it as a bitmap.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CaptureBitmap() As Bitmap
        Dim b As New Bitmap(Me.Width, Me.Height)
        Dim gr As Graphics = Graphics.FromImage(b)
        SendMessage(WM_PRINT, gr.GetHdc, PRF_NONCLIENT Or PRF_CLIENT Or PRF_CHILDREN)
        gr.ReleaseHdc()
        Return b
    End Function

#Region "Animate Window"

    <Flags()> _
    Public Enum AnimationDirection As Integer
        HorizontalPositive = 1
        HorizontalNegative = 2
        VerticalPositive = 4
        VerticalNegative = 8
    End Enum

    Private Const AW_CENTER As Integer = 16
    Private Const AW_HIDE As Integer = 65536
    Private Const AW_ACTIVATE As Integer = 131072
    Private Const AW_BLEND As Integer = 524288

    ''' <summary>
    ''' Fades the window in or out.
    ''' </summary>
    ''' <param name="time">The number of milliseconds to run the animation.</param>
    ''' <param name="activate">Whether to activate the window.</param>
    ''' <param name="hide">Whether to fade the window in or out.</param>
    ''' <remarks></remarks>
    Public Sub AnimateFade(ByVal time As Integer, Optional ByVal activate As Boolean = True, Optional ByVal hide As Boolean = False)
        AnimateBase(time, activate, hide, AW_BLEND)
    End Sub

    ''' <summary>
    ''' Collapses or expands the window in or out.
    ''' </summary>
    ''' <param name="time">The number of milliseconds to run the animation.</param>
    ''' <param name="activate">Whether to activate the window.</param>
    ''' <param name="hide">Whether to fade the window in or out.</param>
    ''' <remarks></remarks>
    Public Sub AnimateExpandCollapse(ByVal time As Integer, Optional ByVal activate As Boolean = True, Optional ByVal hide As Boolean = False)
        AnimateBase(time, activate, hide, AW_CENTER)
    End Sub

    ''' <summary>
    ''' Whether to do a slide or roll animation.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum SlideRoll
        Roll = 0
        Slide = 262144
    End Enum

    ''' <summary>
    ''' Fades the window in or out.
    ''' </summary>
    ''' <param name="time">The number of milliseconds to run the animation.</param>
    ''' <param name="activate">Whether to activate the window.</param>
    ''' <param name="hide">Whether to fade the window in or out.</param>
    ''' <param name="direction">The direction in which to slide or roll. This may be more than one for diagonal animations.</param>
    ''' <param name="type">The type of animation to play.</param>
    ''' <remarks></remarks>
    Public Sub AnimateSlideRoll(ByVal type As SlideRoll, ByVal direction As AnimationDirection, ByVal time As Integer, Optional ByVal activate As Boolean = True, Optional ByVal hide As Boolean = False)
        AnimateBase(time, activate, hide, type Or direction)
    End Sub

    ''' <summary>
    ''' The function called by the other animation functions to animate the window.
    ''' </summary>
    ''' <param name="time"></param>
    ''' <param name="activate"></param>
    ''' <param name="hide"></param>
    ''' <param name="base"></param>
    ''' <remarks></remarks>
    Private Sub AnimateBase(ByVal time As Integer, ByVal activate As Boolean, ByVal hide As Boolean, ByVal base As Integer)
        Dim flags As Integer = base
        If hide Then
            flags = flags Or AW_HIDE
        Else
            If activate Then flags = flags Or AW_ACTIVATE
        End If
        If AnimateWindow(hWnd, time, flags) = False Then Throw New System.ComponentModel.Win32Exception
    End Sub

#End Region

    Public Enum WindowMessages As Integer
        WM_NULL = &H0
        WM_CREATE = &H1
        WM_DESTROY = &H2
        WM_MOVE = &H3
        WM_SIZE = &H5

        WM_ACTIVATE = &H6

        'WM_ACTIVATE state values

        WA_INACTIVE = 0
        WA_ACTIVE = 1
        WA_CLICKACTIVE = 2

        WM_SETFOCUS = &H7
        WM_KILLFOCUS = &H8
        WM_ENABLE = &HA
        WM_SETREDRAW = &HB
        WM_SETTEXT = &HC
        WM_GETTEXT = &HD
        WM_GETTEXTLENGTH = &HE
        WM_PAINT = &HF
        WM_CLOSE = &H10
        '#ifndef _WIN32_WCE
        WM_QUERYENDSESSION = &H11
        WM_QUERYOPEN = &H13
        WM_ENDSESSION = &H16
        '#endif
        WM_QUIT = &H12
        WM_ERASEBKGND = &H14
        WM_SYSCOLORCHANGE = &H15
        WM_SHOWWINDOW = &H18
        WM_WININICHANGE = &H1A
        '#if(WINVER >= =&h0400)
        WM_SETTINGCHANGE = WM_WININICHANGE
        '#endif /* WINVER >= =&h0400 */


        WM_DEVMODECHANGE = &H1B
        WM_ACTIVATEAPP = &H1C
        WM_FONTCHANGE = &H1D
        WM_TIMECHANGE = &H1E
        WM_CANCELMODE = &H1F
        WM_SETCURSOR = &H20
        WM_MOUSEACTIVATE = &H21
        WM_CHILDACTIVATE = &H22
        WM_QUEUESYNC = &H23

        WM_GETMINMAXINFO = &H24

        WM_PAINTICON = &H26
        WM_ICONERASEBKGND = &H27
        WM_NEXTDLGCTL = &H28
        WM_SPOOLERSTATUS = &H2A
        WM_DRAWITEM = &H2B
        WM_MEASUREITEM = &H2C
        WM_DELETEITEM = &H2D
        WM_VKEYTOITEM = &H2E
        WM_CHARTOITEM = &H2F
        WM_SETFONT = &H30
        WM_GETFONT = &H31
        WM_SETHOTKEY = &H32
        WM_GETHOTKEY = &H33
        WM_QUERYDRAGICON = &H37
        WM_COMPAREITEM = &H39
        '#if(WINVER >= =&h0500)
        '#ifndef _WIN32_WCE
        WM_GETOBJECT = &H3D
        '#End If
        '#endif /* WINVER >= =&h0500 */
        WM_COMPACTING = &H41
        WM_COMMNOTIFY = &H44  'no longer suported
        WM_WINDOWPOSCHANGING = &H46
        WM_WINDOWPOSCHANGED = &H47

        WM_POWER = &H48

        'wParam for WM_POWER window message and DRV_POWER driver notification

        PWR_OK = 1
        PWR_FAIL = (-1)
        PWR_SUSPENDREQUEST = 1
        PWR_SUSPENDRESUME = 2
        PWR_CRITICALRESUME = 3

        WM_COPYDATA = &H4A
        WM_CANCELJOURNAL = &H4B

        WM_NCCREATE = &H81
    End Enum

    'Struct pointed to by WM_GETMINMAXINFO lParam
    Private Structure MINMAXINFO
        Public ptReserved As APIPOINT
        Public ptMaxSize As APIPOINT
        Public ptMaxPosition As APIPOINT
        Public ptMinTrackSize As APIPOINT
        Public ptMaxTrackSize As APIPOINT
    End Structure

    Private Structure APIPOINT
        Public x As Integer
        Public y As Integer
    End Structure

    'lParam of WM_COPYDATA message points to...

    Private Structure COPYDATASTRUCT
        Public dwDataPtr As Integer 'ULONG_PTR
        Public cbData As Integer
        Public lpData As Integer
    End Structure

End Class