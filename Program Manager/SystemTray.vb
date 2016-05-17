Imports System.Runtime.InteropServices, Microsoft.Win32

Module SystemTray

    Private TrayWndProc As NativeWindowEx.WndProcDelegate
    Private RefreshTimerDelegate As System.Threading.TimerCallback
    Private RefreshTimer As System.Threading.Timer
    'WithEvents tmrRefresh As New Timer

    WithEvents trayWnd As New NativeWindowEx
    WithEvents trayMng As New NativeWindowEx

#Region "Init/Uninit"

    Public Sub InitSystemTray()
        Try
            'First, register the tray class
            Dim ccp As New ClassCreateParams
            ccp.Name = "Shell_TrayWnd"
            ccp.Style = ClassCreateParams.ClassStyles.DoubleClicks
            TrayWndProc = New NativeWindowEx.WndProcDelegate(AddressOf SysTrayProc)
            ccp.WndProc = TrayWndProc
            trayWnd.RegisterClass(ccp)

            'Now, create the tray window
            Dim cp As New CreateParamsEx
            cp.ClassName = "Shell_TrayWnd"
            cp.Caption = "Program Manager"
            cp.Style = &H80000000 ' WS_POPUP
            cp.ExStyle = &H8 'WS_EX_TOOLWINDOW or WS_EX_TOPMOST
            trayWnd.CreateHandle(cp)
        Catch ex As Exception
            'We don't want an error to occur in this try block.
            'If necessary, it can occur in the next one.
            Debug.Print("Error creating Shell_TrayWnd. " & ex.Message)
            'Stop
            MessageBox.Show("An error occurred while creating the system tray. The system tray will not be available because: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        If trayWnd.Handle <> 0 Then
            Try
                'Register the child window sometimes used by other apps
                Dim ccp As New ClassCreateParams
                ccp.Name = "TrayNotifyWnd"
                ccp.Style = &H8 'CS_DBLCLKS
                trayMng.RegisterClass(ccp)

                'Now, create the window
                Dim cp As New CreateParamsEx
                cp.ClassName = "TrayNotifyWnd"
                cp.Parent = trayWnd.Handle
                cp.Style = &H40000000 'WS_CHILD
                trayMng.CreateHandle(cp)
            Catch ex As Exception
                'An error here can be okay.
                Debug.Print("Error creating TrayNotifyWnd. " & ex.Message)
            Finally
                BroadcastTaskbarCreated()
            End Try
        End If
        LoadShellServiceObjects()
        RefreshTimerDelegate = New System.Threading.TimerCallback(AddressOf TimerRefreshTray)
        RefreshTimer = New System.Threading.Timer(RefreshTimerDelegate, Nothing, 1000, 250)
        'tmrRefresh.Interval = 2000
        'tmrRefresh.Enabled = True
    End Sub

    ''' <summary>
    ''' Call this after you have created a taskbar-like window.
    '''
    ''' According to MSDN: "With Internet Explorer 4.0 and later, the Shell notifies applications
    ''' that the taskbar has been created. When the taskbar is created, it registers a message with
    ''' the TaskbarCreated string and then broadcasts this message to all top-level windows.
    ''' When your taskbar application receives this message, it should assume that any taskbar icons
    ''' it added have been removed and add them again."
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub BroadcastTaskbarCreated()
        PostMessage(HWND_BROADCAST, RegisterWindowMessage("TaskbarCreated"), 0, 0)
    End Sub

    Public Sub UninitSystemTray()
        Try
            UnloadShellServiceObjects()

            TrayWndProc = Nothing

            If RefreshTimer IsNot Nothing Then
                RefreshTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                RefreshTimer.Dispose()
            End If
            If RefreshTimerDelegate IsNot Nothing Then RefreshTimerDelegate = Nothing

            Icons.Clear()

            trayMng.DestroyHandle()
            trayWnd.DestroyHandle()
            trayMng.UnRegisterClass("TrayNotifyWnd")
            trayWnd.UnRegisterClass("Shell_TrayWnd")
        Catch ex As Exception
            Debug.Print("Error while uninitializing systray: " + ex.Message)
        End Try
    End Sub

#End Region

#Region "Shell Service Objects"

    Dim shellServiceObjects As New ObjectModel.Collection(Of IOleCommandTarget)

    Dim CGID_ShellServiceObject As New Guid(&H214D2L, 0, 0, &HC0, 0, 0, 0, 0, 0, 0, &H46)

    Private Sub LoadShellServiceObjects()
        Try
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad")
            For Each n As String In key.GetValueNames
                If key.GetValueKind(n) = RegistryValueKind.String Then
                    Try
                        Dim g As New Guid(CStr(key.GetValue(n)))
                        Dim shellObj As IOleCommandTarget = CoCreateInstance(g, Nothing, CLSCTX.CLSCTX_INPROC_SERVER, New Guid("b722bccb-4e68-101b-a2bc-00aa00404770"))
                        shellObj.Exec(CGID_ShellServiceObject, 2, 0, Nothing, Nothing)
                        shellServiceObjects.Add(shellObj)
                    Catch ex As Exception
                        Debug.Print("Error loading shell service object: " & ex.Message)
                    End Try
                End If
            Next
        Catch ex As Exception
            Debug.Print("Error loading shell service objects: " & ex.Message)
        End Try
    End Sub

    Private Sub UnloadShellServiceObjects()
        For Each i As IOleCommandTarget In shellServiceObjects
            Try
                i.Exec(CGID_ShellServiceObject, 3, 0, Nothing, Nothing)
            Catch ex As Exception
                Debug.Print("Error unloading shell service object: " & ex.Message)
            End Try
        Next
        shellServiceObjects.Clear()     'Release all COM objects
    End Sub

    <Flags()> _
    Enum CLSCTX As UInteger
        CLSCTX_INPROC_SERVER = &H1
        CLSCTX_INPROC_HANDLER = &H2
        CLSCTX_LOCAL_SERVER = &H4
        CLSCTX_INPROC_SERVER16 = &H8
        CLSCTX_REMOTE_SERVER = &H10
        CLSCTX_INPROC_HANDLER16 = &H20
        CLSCTX_RESERVED1 = &H40
        CLSCTX_RESERVED2 = &H80
        CLSCTX_RESERVED3 = &H100
        CLSCTX_RESERVED4 = &H200
        CLSCTX_NO_CODE_DOWNLOAD = &H400
        CLSCTX_RESERVED5 = &H800
        CLSCTX_NO_CUSTOM_MARSHAL = &H1000
        CLSCTX_ENABLE_CODE_DOWNLOAD = &H2000
        CLSCTX_NO_FAILURE_LOG = &H4000
        CLSCTX_DISABLE_AAA = &H8000
        CLSCTX_ENABLE_AAA = &H10000
        CLSCTX_FROM_DEFAULT_CONTEXT = &H20000
        CLSCTX_INPROC = CLSCTX_INPROC_SERVER Or CLSCTX_INPROC_HANDLER
        CLSCTX_SERVER = CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER Or CLSCTX_REMOTE_SERVER
        CLSCTX_ALL = CLSCTX_SERVER Or CLSCTX_INPROC_HANDLER
    End Enum

    <DllImport("ole32.dll", ExactSpelling:=True, PreserveSig:=False)> _
    Private Function CoCreateInstance( _
<[In](), MarshalAs(UnmanagedType.LPStruct)> ByVal rclsid As Guid, _
<MarshalAs(UnmanagedType.IUnknown)> ByVal pUnkOuter As Object, _
ByVal dwClsContext As CLSCTX, _
<[In](), MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid) As <MarshalAs(UnmanagedType.Interface)> Object
    End Function

    <ComImport(), _
    Guid("b722bccb-4e68-101b-a2bc-00aa00404770"), _
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    Private Interface IOleCommandTarget
        'IMPORTANT: The order of the methods is critical here. You
        'perform early binding in most cases, so the order of the methods
        'here MUST match the order of their vtable layout (which is determined
        'by their layout in IDL). The interop calls key off the vtable
        'ordering, not the symbolic names. Therefore, if you switched these
        'method declarations and tried to call the Exec method on an
        'IOleCommandTarget interface from your application, it would
        'translate into a call to the QueryStatus method instead.
        Sub QueryStatus( _
         ByRef pguidCmdGroup As Guid, _
         ByVal cCmds As UInt32, _
         <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal prgCmds() As OLECMD, _
         ByRef CmdText As OLECMDTEXT)
        Sub Exec( _
         ByRef pguidCmdGroup As Guid, _
         ByVal nCmdId As UInteger, _
         ByVal nCmdExecOpt As UInteger, _
         ByRef pvaIn As Object, _
         ByRef pvaOut As Object)
    End Interface

    Private Structure OLECMD
        Public cmdID As ULong
        Public cmdf As Integer
    End Structure

    Private Structure OLECMDTEXT
        Public cmdtextf As Integer
        Public cwActual As ULong
        Public cwBuf As ULong
        Public rgwz() As Char
    End Structure

#End Region

#Region "Interop"

    Private Const HWND_BROADCAST As Integer = &HFFFF
    Private Declare Auto Function RegisterWindowMessage Lib "user32.dll" (ByVal lpString As String) As Integer
    Private Declare Auto Function PostMessage Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wparam As Integer, ByVal lparam As Integer) As Integer

    Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Private Declare Function CopyIcon Lib "user32.dll" (ByVal hIcon As IntPtr) As IntPtr
    Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As IntPtr) As Integer

    Private Declare Auto Function SHLockShared Lib "shell32.dll" (ByVal hData As IntPtr, ByVal dwSourceProcessId As Integer) As IntPtr
    Private Declare Auto Function SHUnlockShared Lib "shell32.dll" (ByVal data As IntPtr) As Boolean

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure NOTIFYICONDATAV5
        Public cbSize As Int32
        Public hwnd As IntPtr
        Public uID As Int32
        Public uFlags As Int32
        Public uCallbackMessage As IntPtr
        Public hIcon As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> _
        Public szTip As String
        Public dwState As Int32
        Public dwStateMask As Int32
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> _
        Public szInfo As String
        Public uTimeout As Int32
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> _
        Public szInfoTitle As String
        Public dwInfoFlags As Int32
    End Structure

    Private Structure COPYDATASTRUCT
        Public dwData As Int32
        Public cdData As Int32
        Public lpData As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Friend Structure NOTIFYICONDATA
        Dim cbSize As Integer
        Dim hWnd As IntPtr
        Dim uID As Integer
        Dim uFlags As Integer
        Dim uCallbackMessage As Integer
        Dim hIcon As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Dim szTip As String
        Dim dwState As Integer
        Dim dwStateMask As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Dim szInfo As String
        Dim uTimeout As Integer ' ignore the uVersion union
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Dim szInfoTitle As String
        Dim dwInfoFlags As Integer
        Dim guidItem As Guid
    End Structure

    Private Structure SHELLTRAYDATA
        Dim dwUnknown As Integer
        Dim dwMessage As Integer
        Dim iconData As NOTIFYICONDATAV5
    End Structure

    Const WM_COPYDATA As Integer = &H4A

    Const NIM_ADD As Integer = 0
    Const NIM_MODIFY As Integer = 1
    Const NIM_DELETE As Integer = 2
    Const NIM_SETFOCUS As Integer = 3
    Const NIM_SETVERSION As Integer = 4

    Const NIF_MESSAGE As Integer = 1
    Const NIF_ICON As Integer = 2
    Const NIF_TIP As Integer = 4
    Const NIF_STATE As Integer = 8
    Const NIF_INFO As Integer = 16
    Const NIF_GUID As Integer = 32

    Const NIS_HIDDEN As Integer = 1
    Const NIS_SHAREDICON As Integer = 2

    Const NOTIFYICON_VERSION As Integer = 3
    Const NOTIFYICON_VERSION_4 As Integer = 4

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SHELLABDATAXP
        Public abd As APPBARDATA
        Public dwOddThisShouldNotBeHere As Integer
        Public dwMessage As Integer
        Public hSharedHandle As IntPtr
        Public dwProcId As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SHELLABDATA2K
        Public abd As APPBARDATA
        Public dwMessage As Integer
        Public hSharedHandle As IntPtr
        Public dwProcId As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer

        Public Sub New(ByVal rect As Rectangle)
            left = rect.Left
            top = rect.Top
            right = rect.Right
            bottom = rect.Bottom
        End Sub

        Public Function ToRect() As Rectangle
            Return Rectangle.FromLTRB(left, top, right, bottom)
        End Function

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
    Private Enum ABEdge As Integer
        Left = 0
        Top
        Right
        Bottom
    End Enum

#End Region

    Private Function SysTrayProc(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wparam As IntPtr, ByVal lparam As IntPtr) As IntPtr
        Try
            Select Case msg
                Case WM_COPYDATA
                    Dim cds As New COPYDATASTRUCT
                    cds = CType(Marshal.PtrToStructure(lparam, GetType(COPYDATASTRUCT)), COPYDATASTRUCT)

                    'Debug.Print("SysTrayProc -> WM_COPYDATA -> cds.dwData = " + cds.dwData.ToString)
                    'Dim std As New SHELLTRAYDATA

                    Dim nid As New NOTIFYICONDATA
                    Select Case cds.dwData
                        Case 0
                            '----------------------------------
                            'AppBar Messages
                            '----------------------------------

                            Dim abData As SHELLABDATA2K

                            'We need to get the appbar data here, but it is different for different OS's
                            If (Environment.OSVersion.Platform = PlatformID.Win32NT AndAlso Environment.OSVersion.Version.Major >= 5) Then
                                If (Environment.OSVersion.Version.Minor = 0) Then
                                    'Windows 2000
                                    Dim pdata As SHELLABDATA2K = Marshal.PtrToStructure(cds.lpData, GetType(SHELLABDATA2K))

                                    abData.abd = pdata.abd
                                    abData.dwMessage = pdata.dwMessage
                                    abData.dwProcId = pdata.dwProcId
                                    abData.hSharedHandle = pdata.hSharedHandle
                                Else
                                    'Windows XP
                                    Dim pdata As SHELLABDATAXP = Marshal.PtrToStructure(cds.lpData, GetType(SHELLABDATAXP))

                                    abData.abd = pdata.abd
                                    abData.dwMessage = pdata.dwMessage
                                    abData.dwProcId = pdata.dwProcId
                                    abData.hSharedHandle = pdata.hSharedHandle
                                End If
                            Else
                                'Windows 9X
                                Dim pdata As SHELLABDATA2K = Marshal.PtrToStructure(cds.lpData, GetType(SHELLABDATA2K))

                                abData.abd = pdata.abd
                                abData.dwMessage = pdata.dwMessage
                                abData.dwProcId = pdata.dwProcId
                                abData.hSharedHandle = pdata.hSharedHandle
                            End If

                            'Debug.Print(CType(abData.dwMessage, ABMsg).ToString)

                            Select Case CType(abData.dwMessage, ABMsg)
                                Case ABMsg.ABM_GETTASKBARPOS, ABMsg.ABM_SETPOS, ABMsg.ABM_QUERYPOS
                                    Return AppBarTransferHelper(abData)
                                Case ABMsg.ABM_NEW
                                    Return AddAppBar(abData)
                                Case ABMsg.ABM_REMOVE
                                    Return RemoveAppBar(abData)
                                Case ABMsg.ABM_ACTIVATE
                                    Return True
                                Case ABMsg.ABM_WINDOWPOSCHANGED
                                    Return True
                                Case ABMsg.ABM_GETAUTOHIDEBAR
                                    Return False
                                Case ABMsg.ABM_SETAUTOHIDEBAR
                                    Return False
                                Case ABMsg.ABM_GETSTATE
                                    Return False
                                Case ABMsg.ABM_SETSTATE
                                    Return False
                                Case Else
                                    Debug.Print("Unknown AppBar Message")
                            End Select

                            Return True     'If it wasn't handled

                        Case 1
                            '----------------------------------
                            'Notify Icon Messages
                            '----------------------------------

                            nid = CType(Marshal.PtrToStructure(New IntPtr(cds.lpData.ToInt64 + 8), GetType(NOTIFYICONDATA)), NOTIFYICONDATA)
                            Dim traycmd As Int32 = Marshal.ReadInt32(New IntPtr(cds.lpData.ToInt64 + 4))
                            Select Case traycmd
                                Case NIM_ADD
                                    'Debug.Print("SysTrayProc -> WM_COPYDATA -> traycmd = NIM_ADD")
                                    Return AddNotifyIcon(nid)
                                Case NIM_MODIFY
                                    'Debug.Print("SysTrayProc -> WM_COPYDATA -> traycmd = NIM_MODIFY")
                                    Dim bHidden As Boolean = (NID_IsHidden(nid) = 1)
                                    If ContainsNotifyIcon(nid.hWnd, nid.uID) And bHidden Then
                                        Return RemoveNotifyIcon(GetNotifyIcon(nid.hWnd, nid.uID))
                                    End If
                                    'If Not ContainsTrayItem(nid.hWnd, nid.uID) Then
                                    'Return AddTrayItem(nid)
                                    'Else
                                    Return ModifyNotifyIcon(GetNotifyIcon(nid.hWnd, nid.uID), nid)
                                    'End If
                                Case NIM_DELETE
                                    'Debug.Print("SysTrayProc -> WM_COPYDATA -> traycmd = NIM_DELETE")
                                    If ContainsNotifyIcon(nid.hWnd, nid.uID) Then
                                        Return RemoveNotifyIcon(GetNotifyIcon(nid.hWnd, nid.uID))
                                    End If
                                Case NIM_SETVERSION
                                    'Debug.Print("SysTrayProc -> WM_COPYDATA -> traycmd = NIM_SETVERSION")
                                    If ContainsNotifyIcon(nid.hWnd, nid.uID) Then
                                        Return SetNotifyIconVersion(GetNotifyIcon(nid.hWnd, nid.uID), nid)
                                    End If
                                Case NIM_SETFOCUS
                                    'Debug.Print("SysTrayProc -> WM_COPYDATA -> traycmd = NIM_SETFOCUS")
                                    If ContainsNotifyIcon(nid.hWnd, nid.uID) Then
                                        RaiseEvent NotifyIconSetFocus(GetNotifyIcon(nid.hWnd, nid.uID))
                                        Return True
                                    End If
                            End Select
                    End Select
                    Return False
                Case RegisterWindowMessage("AppBarUpdateNotification")
                    For Each a As AppBar In AppBars
                        If CInt(a.Edge) = wparam AndAlso a.Handle <> lparam Then
                            PostMessage(a.Handle, a.CallbackMessage, ABNotify.ABN_POSCHANGED, 0)
                        End If
                    Next
                Case Else
                    Return DefWindowProc(hwnd, msg, wparam, lparam)
            End Select

        Catch ex As Exception
            'An error occurred! This is not good!
            Debug.Print("An error occurred in SysTrayProc: " & ex.Message)
        End Try
    End Function

#Region "Notify Icon Helpers"

    Private Function AddNotifyIcon(ByVal nid As NOTIFYICONDATA) As Boolean
        SyncLock Icons
            If ContainsNotifyIcon(nid) Then Return False

            If WindowInfo.IsValid(nid.hWnd) And (Not ((NID_IsShared(nid)) And (nid.hIcon = IntPtr.Zero))) And nid.uID >= 1 Then
                Dim newItem As New TrayNotifyIcon(nid)
                'Dim newInfo As New NOTIFYICONDATA

                'This can be removed since it was added automatically by the TrayNotifyIcon constructor
                'newInfo.hWnd = nid.hWnd
                'newInfo.cbSize = nid.cbSize
                'newInfo.uFlags = nid.uFlags
                'newInfo.uCallbackMessage = nid.uCallbackMessage
                'newInfo.uID = nid.uID

                CopyIconData(newItem, nid)
                CopyVersionSpecificData(newItem, nid)

                'newItem.Hidden = False
                'newItem.Item = newInfo

                Icons.Add(newItem)
                RaiseEvent NotifyIconAdded(newItem)
                Return True
            End If
            Return False
        End SyncLock
    End Function

    Private Function ModifyNotifyIcon(ByVal ti As TrayNotifyIcon, ByVal nid As NOTIFYICONDATA) As Boolean
        'Dim bChanged As Boolean = False
        Dim iHidden As Boolean = False
        Dim bRet As Boolean = True

        If ti Is Nothing Then Return False

        'update flags
        'ti.Item.uFlags = nid.uFlags


        'check for callback message validity
        If nid.uFlags And NIF_MESSAGE Then
            ti.CallbackMessage = nid.uCallbackMessage
            'bChanged = True
        End If

        'icon
        'If CopyIconData(ti, nid) Then
        'bChanged = True
        'End If
        bRet = CopyIconData(ti, nid)

        'tooltip
        If CopyVersionSpecificData(ti, nid) Then
            'bChanged = True
        End If

        iHidden = NID_IsHidden(nid)
        If ti.Hidden <> (iHidden = True) Then
            If Not iHidden Then
                ti.Hidden = False
            Else
                ti.Hidden = True
            End If
            'bChanged = False
        End If

        RaiseEvent NotifyIconUpdated(ti)

        Return bRet
    End Function

    Private Function RemoveNotifyIcon(ByVal ti As TrayNotifyIcon) As Boolean
        If Icons.Contains(ti) Then
            If ti.Icon IsNot Nothing Then
                DestroyIcon(ti.Icon.Handle)
                ti.Icon = Nothing
            End If
            RaiseEvent NotifyIconBeforeRemoved(ti)
            Icons.Remove(ti)
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsValidIcon(ByVal nid As NOTIFYICONDATA) As Boolean
        Return ((nid.uFlags And NIF_ICON) And (Not NID_IsHidden(nid)))
    End Function

    Private Function CopyIconData(ByVal ti As TrayNotifyIcon, ByVal nid As NOTIFYICONDATA) As Boolean
        Dim bRet As Boolean = True
        If (nid.uFlags And NIF_ICON) And (nid.hIcon <> IntPtr.Zero) Then
            If ti.Icon IsNot Nothing Then DestroyIcon(ti.Icon.Handle)

            ti.Icon = Icon.FromHandle(CopyIcon(nid.hIcon))

            'TODO: do visual icon update here?

            If ti.Icon Is Nothing Then
                bRet = False
            End If
        End If
        Return bRet
    End Function

    Private Function CopyVersionSpecificData(ByVal ti As TrayNotifyIcon, ByVal nid As NOTIFYICONDATA) As Boolean
        Dim ret As Boolean = False

        If nid.uFlags And NIF_STATE Then
            ti.State = nid.dwState
            ti.StateMask = nid.dwStateMask
            ret = True
        End If

        If nid.uFlags And NIF_INFO Then
            Dim sbTmp As New System.Text.StringBuilder
            sbTmp.Capacity = 256
            sbTmp.Append(nid.szInfo)
            ti.BalloonText = sbTmp.ToString 'nid.szInfo.Substring(0, 256)
            sbTmp = New System.Text.StringBuilder
            sbTmp.Capacity = 64
            sbTmp.Append(nid.szInfoTitle)
            ti.BalloonTitle = sbTmp.ToString 'nid.szInfoTitle.Substring(0, 64)
            'ti.InfoFlags = nid.dwInfoFlags
            ti.BalloonTimeout = nid.uTimeout
            ret = True
        End If

        If nid.uFlags And NIF_TIP Then
            If String.Compare(ti.Tooltip, nid.szTip) <> 0 Then
                Dim sbTmp As New System.Text.StringBuilder
                sbTmp.Capacity = 128
                sbTmp.Append(nid.szTip)
                ti.Tooltip = sbTmp.ToString     'nid.szTip.Substring(0, 128)
                ret = True
            End If
        End If

        Return ret
    End Function

    Private Function SetNotifyIconVersion(ByVal ti As TrayNotifyIcon, ByVal nid As NOTIFYICONDATA) As Boolean
        If nid.uTimeout = 0 Or nid.uTimeout = 3 Then
            ti.Version = nid.uTimeout 'Union with uVersion
            Return True
        Else
            Return False
        End If
    End Function

    Private Function NID_IsHidden(ByVal nid As NOTIFYICONDATA) As Boolean
        Dim Hidden As Boolean = False
        If nid.uFlags And NIF_STATE Then
            If nid.dwStateMask And NIS_HIDDEN Then
                If nid.dwState And NIS_HIDDEN Then
                    Hidden = True
                Else
                    Hidden = False
                End If
            End If
        End If
        Return Hidden
    End Function

    Private Function NID_IsShared(ByVal nid As NOTIFYICONDATA)
        Dim isShared As Boolean = False
        If nid.uFlags And NIF_STATE Then
            If nid.dwStateMask And NIS_SHAREDICON Then
                If nid.dwState And NIS_SHAREDICON Then
                    isShared = True
                Else
                    isShared = False
                End If
            End If
        End If
        Return isShared
    End Function

    Public Sub RefreshTray()
        Dim remitems As New ObjectModel.Collection(Of TrayNotifyIcon)
        SyncLock Icons
            For Each i As TrayNotifyIcon In Icons
                If Not WindowInfo.IsValid(i.WindowHandle) Then
                    remitems.Add(i)
                End If
            Next
        End SyncLock
        For Each i As TrayNotifyIcon In remitems
            RemoveNotifyIcon(i)
        Next
    End Sub

    Private Sub TimerRefreshTray(ByVal state As Object)
        RefreshTray()
    End Sub

#End Region

#Region "AppBar Helpers"

    Private Function AppBarTransferHelper(ByVal pData As SHELLABDATA2K) As Boolean
        If Environment.OSVersion.Platform = PlatformID.Win32NT OrElse Environment.OSVersion.Version.Minor >= 10 Then
            Dim appBarDataObj As APPBARDATA
            Dim appBarDataPointer As IntPtr = SHLockShared(pData.hSharedHandle, pData.dwProcId)
            appBarDataObj = Marshal.PtrToStructure(appBarDataPointer, GetType(APPBARDATA))

            'Do stuff with the APPBARDATA structure
            Select Case pData.dwMessage
                Case ABMsg.ABM_QUERYPOS
                    CalcAppBarPos(appBarDataObj)
                Case ABMsg.ABM_SETPOS
                    CalcAppBarPos(appBarDataObj)
                    If ContainsAppBar(appBarDataObj.hWnd) Then
                        Dim a As AppBar = GetAppBar(appBarDataObj.hWnd)
                        a.Edge = appBarDataObj.uEdge
                        a.Rectangle = Rectangle.FromLTRB(appBarDataObj.rc.left, appBarDataObj.rc.top, appBarDataObj.rc.right, appBarDataObj.rc.bottom)
                    End If
                    UpdateWorkingArea()
                Case ABMsg.ABM_GETTASKBARPOS
                    appBarDataObj.rc = New RECT(New Rectangle(0, SystemInformation.WorkingArea.Bottom, Screen.PrimaryScreen.Bounds.Width, Math.Max(1, Screen.PrimaryScreen.Bounds.Bottom - SystemInformation.WorkingArea.Bottom)))
                    appBarDataObj.uEdge = ABEdge.Bottom
            End Select

            Marshal.StructureToPtr(appBarDataObj, appBarDataPointer, False)
            SHUnlockShared(appBarDataPointer)
        End If

        Return True
    End Function

    Private Sub CalcAppBarPos(ByRef ab As APPBARDATA)
        Dim s As Screen = Screen.FromRectangle(ab.rc.ToRect)
        Select Case CType(ab.uEdge, ABEdge)
            Case ABEdge.Bottom, ABEdge.Top
                'Clip the height to a quarter of the screen if needed
                Dim height As Integer = ab.rc.bottom - ab.rc.top
                height = Math.Min(height, s.Bounds.Height / 4)

                'Stretch the appbar to fill the width of the screen
                ab.rc.left = s.Bounds.Left
                ab.rc.right = s.Bounds.Right

                'Calculate an offset based on the other appbars from the bottom or top of the screen
                Dim offset As Integer = 0
                For Each a As AppBar In AppBars
                    If Screen.FromRectangle(a.Rectangle).Bounds = s.Bounds AndAlso a.Edge = ab.uEdge Then
                        If a.Handle <> ab.hWnd Then
                            offset += a.Rectangle.Height
                        Else
                            Exit For
                        End If
                    End If
                Next

                'Finally, apply the offset from the top or bottom edge
                Select Case CType(ab.uEdge, ABEdge)
                    Case ABEdge.Bottom
                        ab.rc.bottom = s.Bounds.Bottom - offset
                        ab.rc.top = ab.rc.bottom - height
                    Case ABEdge.Top
                        ab.rc.top = s.Bounds.Top + offset
                        ab.rc.bottom = ab.rc.top + height
                End Select

                ScheduleUpdateEdge(ABEdge.Left, 0)
                ScheduleUpdateEdge(ABEdge.Right, 0)

            Case ABEdge.Left, ABEdge.Right
                'Clip the width to a quarter of the screen if needed
                Dim width As Integer = ab.rc.right - ab.rc.left
                width = Math.Min(width, s.Bounds.Width / 4)

                'Calculate offsets based on the other appbars from the bottom or top of the screen
                Dim topoffset As Integer = 0
                Dim bottomoffset As Integer = 0
                For Each a As AppBar In AppBars
                    If a.Handle <> ab.hWnd AndAlso Screen.FromRectangle(a.Rectangle).Bounds = s.Bounds Then
                        Select Case a.Edge
                            Case AppbarForm.ABEdge.Bottom
                                bottomoffset += a.Rectangle.Height
                            Case AppbarForm.ABEdge.Top
                                topoffset += a.Rectangle.Height
                        End Select
                    End If
                Next
                ab.rc.top = s.Bounds.Top + topoffset
                ab.rc.bottom = s.Bounds.Bottom - bottomoffset

                'Calculate an offset from the left or right side
                Dim offset As Integer = 0
                For Each a As AppBar In AppBars
                    If Screen.FromRectangle(a.Rectangle).Bounds = s.Bounds AndAlso a.Edge = ab.uEdge Then
                        If a.Handle <> ab.hWnd Then
                            offset += a.Rectangle.Width
                        Else
                            Exit For
                        End If
                    End If
                Next

                'Finally, apply the offset from the left or right edge
                Select Case CType(ab.uEdge, ABEdge)
                    Case ABEdge.Left
                        ab.rc.left = s.Bounds.Left + offset
                        ab.rc.right = ab.rc.left + width
                    Case ABEdge.Right
                        ab.rc.right = s.Bounds.Right - offset
                        ab.rc.left = ab.rc.right - width
                End Select
        End Select
    End Sub

    Private Sub UpdateWorkingArea()
        Dim area As New Padding
        For Each a As AppBar In AppBars
            Try
                Select Case a.Edge
                    Case AppbarForm.ABEdge.Bottom
                        area.Bottom += a.Rectangle.Height
                    Case AppbarForm.ABEdge.Left
                        area.Left += a.Rectangle.Width
                    Case AppbarForm.ABEdge.Right
                        area.Right += a.Rectangle.Width
                    Case AppbarForm.ABEdge.Top
                        area.Top += a.Rectangle.Height
                End Select
            Catch ex As Exception

            End Try
        Next
        Dim r As New Rectangle(area.Left, area.Top, Screen.PrimaryScreen.Bounds.Width - area.Horizontal, Screen.PrimaryScreen.Bounds.Height - area.Vertical)
        '        SystemInformationWritable.SetWorkingArea(r)
    End Sub

    Private Function AddAppBar(ByVal pData As SHELLABDATA2K) As Boolean
        If WindowInfo.IsValid(pData.abd.hWnd) AndAlso Not ContainsAppBar(pData.abd.hWnd) Then
            Dim a As New AppBar
            a.Handle = pData.abd.hWnd
            a.CallbackMessage = pData.abd.uCallbackMessage
            a.AutoHide = (pData.dwMessage = ABMsg.ABM_SETAUTOHIDEBAR)
            AppBars.Add(a)
            Return True
        Else
            Return False
        End If
    End Function

    Private Function RemoveAppBar(ByVal pData As SHELLABDATA2K) As Boolean
        If ContainsAppBar(pData.abd.hWnd) Then
            AppBars.Remove(GetAppBar(pData.abd.hWnd))
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub ScheduleUpdateEdge(ByVal edge As ABEdge, ByVal hwnd As IntPtr)
        PostMessage(trayWnd.Handle, RegisterWindowMessage("AppBarUpdateNotification"), CInt(edge), hwnd)
    End Sub

#End Region

#Region "Public Helper Functions"

    Public Function ContainsNotifyIcon(ByVal item As TrayNotifyIcon) As Boolean
        Return ContainsNotifyIcon(item.WindowHandle, item.Id)
    End Function

    Private Function ContainsNotifyIcon(ByVal item As NOTIFYICONDATA) As Boolean
        Return ContainsNotifyIcon(item.hWnd, item.uID)
    End Function

    Public Function ContainsNotifyIcon(ByVal winhandle As IntPtr, ByVal id As Integer) As Boolean
        For Each i As TrayNotifyIcon In Icons
            If i.WindowHandle = winhandle AndAlso i.Id = id Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetNotifyIcon(ByVal winhandle As IntPtr, ByVal id As Integer) As TrayNotifyIcon
        For Each i As TrayNotifyIcon In Icons
            If i.WindowHandle = winhandle AndAlso i.Id = id Then
                Return i
            End If
        Next
        Return Nothing
    End Function

    Public Function ContainsAppBar(ByVal item As AppBar)
        For Each i As AppBar In AppBars
            If i = item Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function ContainsAppBar(ByVal winhandle As IntPtr)
        For Each i As AppBar In AppBars
            If i.Handle = winhandle Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetAppBar(ByVal winhandle As IntPtr)
        For Each i As AppBar In AppBars
            If i.Handle = winhandle Then
                Return i
            End If
        Next
        Return Nothing
    End Function

#End Region

#Region "Events"

    Public Event NotifyIconAdded(ByVal item As TrayNotifyIcon)
    Public Event NotifyIconBeforeRemoved(ByVal item As TrayNotifyIcon)
    Public Event NotifyIconUpdated(ByVal item As TrayNotifyIcon)
    Public Event NotifyIconSetFocus(ByVal item As TrayNotifyIcon)

#End Region

#Region "Public Classes"

    Public Class TrayNotifyIcon

        Dim _hwnd As IntPtr

        Public Property WindowHandle() As IntPtr
            Get
                Return _hwnd
            End Get
            Set(ByVal value As IntPtr)
                _hwnd = value
            End Set
        End Property

        Dim _id As Integer

        Public Property Id() As Integer
            Get
                Return _id
            End Get
            Set(ByVal value As Integer)
                _id = value
            End Set
        End Property

        Dim _tooltip As String

        Public Property Tooltip() As String
            Get
                Return _tooltip
            End Get
            Set(ByVal value As String)
                _tooltip = value
            End Set
        End Property

        Dim _icon As Icon

        Public Property Icon() As Icon
            Get
                Return _icon
            End Get
            Set(ByVal value As Icon)
                _icon = value
            End Set
        End Property

        Dim _version As Integer

        Public Property Version() As Integer
            Get
                Return _version
            End Get
            Set(ByVal value As Integer)
                _version = value
            End Set
        End Property

        Dim _msg As Integer

        Public Property CallbackMessage()
            Get
                Return _msg
            End Get
            Set(ByVal value)
                _msg = value
            End Set
        End Property

        Dim _balltext

        Public Property BalloonText() As String
            Get
                Return _balltext
            End Get
            Set(ByVal value As String)
                _balltext = value
            End Set
        End Property

        Dim _balltitle As String

        Public Property BalloonTitle() As String
            Get
                Return _balltitle
            End Get
            Set(ByVal value As String)
                _balltitle = value
            End Set
        End Property

        Dim _state As Integer

        Public Property State() As Integer
            Get
                Return _state
            End Get
            Set(ByVal value As Integer)
                _state = value
            End Set
        End Property

        Dim _stateMask As Integer

        Public Property StateMask() As Integer
            Get
                Return _stateMask
            End Get
            Set(ByVal value As Integer)
                _stateMask = value
            End Set
        End Property

        Dim _timeout As Integer

        Public Property BalloonTimeout() As Integer
            Get
                Return _timeout
            End Get
            Set(ByVal value As Integer)
                If value < 10000 Then
                    _timeout = 10000
                ElseIf value > 30000 Then
                    _timeout = 30000
                Else
                    _timeout = value
                End If
            End Set
        End Property

        Dim _hidden As Boolean

        Public Property Hidden() As Boolean
            Get
                Return _hidden
            End Get
            Set(ByVal value As Boolean)
                _hidden = value
            End Set
        End Property


        Const WM_CONTEXTMENU As Integer = &H7B
        Const WM_MOUSEFIRST As Integer = &H200
        Const WM_MOUSEMOVE As Integer = &H200
        Const WM_LBUTTONDOWN As Integer = &H201
        Const WM_LBUTTONUP As Integer = &H202
        Const WM_LBUTTONDBLCLK As Integer = &H203
        Const WM_RBUTTONDOWN As Integer = &H204
        Const WM_RBUTTONUP As Integer = &H205
        Const WM_RBUTTONDBLCLK As Integer = &H206
        Const WM_MBUTTONDOWN As Integer = &H207
        Const WM_MBUTTONUP As Integer = &H208
        Const WM_MBUTTONDBLCLK As Integer = &H209

        <DllImport("user32.dll")> _
        Private Shared Function AllowSetForegroundWindow(ByVal dwProcessId As Integer) As Boolean
        End Function

        Public Sub SendContextMenu()
            If WindowInfo.IsValid(Me.WindowHandle) Then
                Dim w As New WindowInfo(Me.WindowHandle)
                AllowSetForegroundWindow(w.ProcessId)
                Select Case Me.Version
                    Case 3
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_RBUTTONDOWN, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_CONTEXTMENU, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                    Case 0
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_RBUTTONDOWN, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_RBUTTONUP, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                End Select
            End If
        End Sub

        Public Sub SendMouseMove()
            If WindowInfo.IsValid(Me.WindowHandle) Then
                Dim w As New WindowInfo(Me.WindowHandle)
                AllowSetForegroundWindow(w.ProcessId)
                Select Case Me.Version
                    Case 3, 0
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_MOUSEMOVE, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                End Select
            End If
        End Sub

        Public Sub SendLeftMouseClick()
            If WindowInfo.IsValid(Me.WindowHandle) Then
                Dim w As New WindowInfo(Me.WindowHandle)
                AllowSetForegroundWindow(w.ProcessId)
                Select Case Me.Version
                    Case 3, 0
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_LBUTTONDOWN, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_LBUTTONUP, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                End Select
            End If
        End Sub

        Public Sub SendLeftMouseDoubleClick()
            If WindowInfo.IsValid(Me.WindowHandle) Then
                Dim w As New WindowInfo(Me.WindowHandle)
                AllowSetForegroundWindow(w.ProcessId)
                Select Case Me.Version
                    Case 3, 0
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_LBUTTONDBLCLK, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                End Select
            End If
        End Sub

        Public Sub SendRightMouseDoubleClick()
            If WindowInfo.IsValid(Me.WindowHandle) Then
                Dim w As New WindowInfo(Me.WindowHandle)
                AllowSetForegroundWindow(w.ProcessId)
                Select Case Me.Version
                    Case 3, 0
                        w.SendMessageTimeout(Me.CallbackMessage, Me.Id, WM_RBUTTONDBLCLK, WindowInfo.SendMessageTimeoutFlags.AbortIfHung, 50)
                End Select
            End If
        End Sub

        Friend Sub New(ByVal item As NOTIFYICONDATA)
            Me.WindowHandle = item.hWnd
            Me.Id = item.uID
            'If item.uFlags And NIF_ICON Then Me.Icon = Icon.FromHandle(item.hIcon)
            If item.uFlags And NIF_MESSAGE Then Me.CallbackMessage = item.uCallbackMessage
            'If item.uFlags And NIF_TIP Then Me.Tooltip = item.szTip
        End Sub

        Public Shared Operator =(ByVal obj1 As TrayNotifyIcon, ByVal obj2 As TrayNotifyIcon) As Boolean
            Return obj1.WindowHandle = obj2.WindowHandle AndAlso obj1.Id = obj2.Id
        End Operator

        Public Shared Operator <>(ByVal obj1 As TrayNotifyIcon, ByVal obj2 As TrayNotifyIcon) As Boolean
            Return Not obj1 = obj2
        End Operator

    End Class

    Public Class AppBar

        Dim _hwnd As IntPtr

        Public Property Handle() As IntPtr
            Get
                Return _hwnd
            End Get
            Set(ByVal value As IntPtr)
                _hwnd = value
            End Set
        End Property

        Dim _callbackMsg As Integer

        Public Property CallbackMessage() As Integer
            Get
                Return _callbackMsg
            End Get
            Set(ByVal value As Integer)
                _callbackMsg = value
            End Set
        End Property

        Dim _edge As AppbarForm.ABEdge

        Public Property Edge() As AppbarForm.ABEdge
            Get
                Return _edge
            End Get
            Set(ByVal value As AppbarForm.ABEdge)
                _edge = value
            End Set
        End Property

        Dim _autohide As Boolean

        Public Property AutoHide() As Boolean
            Get
                Return _autohide
            End Get
            Set(ByVal value As Boolean)
                _autohide = value
            End Set
        End Property

        Dim _topmost As Boolean

        Public Property TopMost() As Boolean
            Get
                Return _topmost
            End Get
            Set(ByVal value As Boolean)
                _topmost = value
            End Set
        End Property

        Dim _rect As Rectangle

        Public Property Rectangle() As Rectangle
            Get
                Return _rect
            End Get
            Set(ByVal value As Rectangle)
                _rect = value
            End Set
        End Property

        Public Shared Operator =(ByVal obj1 As AppBar, ByVal obj2 As AppBar) As Boolean
            Return (obj1 Is Nothing AndAlso obj2 Is Nothing) OrElse (obj1 IsNot Nothing AndAlso obj2 IsNot Nothing AndAlso obj1.Handle = obj2.Handle)
        End Operator

        Public Shared Operator <>(ByVal obj1 As AppBar, ByVal obj2 As AppBar) As Boolean
            Return Not obj1 = obj2
        End Operator

    End Class

#End Region

#Region "Public Properties"

    Dim _trayicons As New ObjectModel.Collection(Of TrayNotifyIcon)

    Public ReadOnly Property Icons() As ObjectModel.Collection(Of TrayNotifyIcon)
        Get
            Return _trayicons
        End Get
    End Property

    Dim _appbars As New ObjectModel.Collection(Of AppBar)

    Public ReadOnly Property AppBars() As ObjectModel.Collection(Of AppBar)
        Get
            Return _appbars
        End Get
    End Property

#End Region

End Module
