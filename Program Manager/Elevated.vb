Imports System
Imports System.Runtime.InteropServices
Namespace Win32
    <StructLayout(LayoutKind.Sequential)> _
 Public Class SHELLEXECUTEINFO
        Public cbSize As Integer
        Public fMask As Integer
        Public hwnd As IntPtr
        Public lpVerb As String
        Public lpFile As String
        Public lpParameters As String
        Public lpDirectory As String
        Public nShow As Integer
        Public hInstApp As Integer
        Public lpIDList As Integer
        Public lpClass As String
        Public hkeyClass As Integer
        Public dwHotKey As Integer
        Public hIcon As Integer
        Public hProcess As Integer
    End Class
    Public Class Shell
        Public Const SEE_MASK_INVOKEIDLIST As Integer = 12
        <DllImport("shell32.dll")> _
        Public Shared Function ShellExecuteEx(<[In](), Out()> ByVal execInfo As SHELLEXECUTEINFO) As Boolean
        End Function
    End Class
End Namespace