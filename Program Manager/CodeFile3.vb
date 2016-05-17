Imports System.ComponentModel

Public Interface IShellControl
    Inherits IAlphaPaint

    Property Config() As ShellControlConfig
    Property Background() As ShellRenderer

    Sub BeginUpdate()
    Sub EndUpdate()
    Sub Refresh()

    Function ToString() As String

End Interface

Public Module ShellControlHelper

    Public Function GetConfig(Of configtype As {ShellControlConfig, New})(ByVal control As IShellControl) As configtype
        Dim ctrlprops As PropertyDescriptorCollection = TypeDescriptor.GetProperties(control)
        Dim c As New configtype
        Dim cfgprops As PropertyDescriptorCollection = TypeDescriptor.GetProperties(c)
        For Each p As PropertyDescriptor In cfgprops
            For Each q As PropertyDescriptor In ctrlprops
                If p.Name.Replace("_", "") = q.Name AndAlso p.PropertyType.FullName = q.PropertyType.FullName Then
                    p.SetValue(c, q.GetValue(control))
                End If
            Next
        Next
        Return c
    End Function

    Public Sub SetConfig(ByVal control As IShellControl, ByVal config As ShellControlConfig)
        control.BeginUpdate()
        Dim ctrlprops As PropertyDescriptorCollection = TypeDescriptor.GetProperties(control)
        Dim cfgprops As PropertyDescriptorCollection = TypeDescriptor.GetProperties(config)
        For Each p As PropertyDescriptor In cfgprops
            For Each q As PropertyDescriptor In ctrlprops
                If p.Name.Replace("_", "") = q.Name AndAlso p.PropertyType.FullName = q.PropertyType.FullName Then
                    q.SetValue(control, p.GetValue(config))
                End If
            Next
        Next
        control.EndUpdate()
    End Sub

End Module