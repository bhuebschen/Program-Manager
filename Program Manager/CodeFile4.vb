Imports System.Windows.Forms.Design, System.ComponentModel

<System.ComponentModel.Editor(GetType(IShellRendererSelector), GetType(Drawing.Design.UITypeEditor))> _
Public MustInherit Class ShellRenderer
    Implements ICustomTypeDescriptor

    Public Sub New()

    End Sub

    MustOverride Sub Render(ByVal dest As Graphics, ByVal bounds As Rectangle)

#Region "ICustomTypeDescriptor"

    Public Function GetAttributes() As System.ComponentModel.AttributeCollection Implements System.ComponentModel.ICustomTypeDescriptor.GetAttributes
        Return TypeDescriptor.GetAttributes(Me, True)
    End Function

    Public Function GetClassName() As String Implements System.ComponentModel.ICustomTypeDescriptor.GetClassName
        Return TypeDescriptor.GetClassName(Me, True)
    End Function

    Public Function GetComponentName() As String Implements System.ComponentModel.ICustomTypeDescriptor.GetComponentName
        Return TypeDescriptor.GetComponentName(Me, True)
    End Function

    Public Function GetConverter() As System.ComponentModel.TypeConverter Implements System.ComponentModel.ICustomTypeDescriptor.GetConverter
        Return TypeDescriptor.GetConverter(Me, True)
    End Function

    Public Function GetDefaultEvent() As System.ComponentModel.EventDescriptor Implements System.ComponentModel.ICustomTypeDescriptor.GetDefaultEvent
        Return TypeDescriptor.GetDefaultEvent(Me, True)
    End Function

    Public Function GetDefaultProperty() As System.ComponentModel.PropertyDescriptor Implements System.ComponentModel.ICustomTypeDescriptor.GetDefaultProperty
        Return TypeDescriptor.GetDefaultProperty(Me, True)
    End Function

    Public Function GetEditor(ByVal editorBaseType As System.Type) As Object Implements System.ComponentModel.ICustomTypeDescriptor.GetEditor
        Return TypeDescriptor.GetEditor(Me, editorBaseType, True)
    End Function

    Public Function GetEvents() As System.ComponentModel.EventDescriptorCollection Implements System.ComponentModel.ICustomTypeDescriptor.GetEvents
        Return TypeDescriptor.GetEvents(Me, True)
    End Function

    Public Function GetEvents(ByVal attributes() As System.Attribute) As System.ComponentModel.EventDescriptorCollection Implements System.ComponentModel.ICustomTypeDescriptor.GetEvents
        Return TypeDescriptor.GetEvents(Me, attributes, True)
    End Function

    Public Function GetProperties() As System.ComponentModel.PropertyDescriptorCollection Implements System.ComponentModel.ICustomTypeDescriptor.GetProperties
        Return TypeDescriptor.GetProperties(Me, True)
    End Function

    Public Function GetProperties(ByVal attributes() As System.Attribute) As System.ComponentModel.PropertyDescriptorCollection Implements System.ComponentModel.ICustomTypeDescriptor.GetProperties
        Dim oldprops As PropertyDescriptorCollection = TypeDescriptor.GetProperties(Me.GetType, attributes)
        Dim newprops(oldprops.Count - 1) As PropertyDescriptor
        For i As Integer = 0 To oldprops.Count - 1
            newprops(i) = New ConfigPropertyDescriptor(oldprops(i), attributes)
        Next
        Return New PropertyDescriptorCollection(newprops)
    End Function

    Public Function GetPropertyOwner(ByVal pd As System.ComponentModel.PropertyDescriptor) As Object Implements System.ComponentModel.ICustomTypeDescriptor.GetPropertyOwner
        Return Me
    End Function

#End Region

End Class

Public Class IShellRendererSelector
    Inherits Drawing.Design.UITypeEditor

    Public Overrides Function GetPaintValueSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overrides Sub PaintValue(ByVal e As System.Drawing.Design.PaintValueEventArgs)
        MyBase.PaintValue(e)
        If e.Value IsNot Nothing AndAlso TypeOf (e.Value) Is ShellRenderer Then
            CType(e.Value, ShellRenderer).Render(e.Graphics, e.Bounds)
        End If
    End Sub

    Public Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
        Return Drawing.Design.UITypeEditorEditStyle.Modal
    End Function

    Public Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object) As Object
        'Dim frmsvr As IWindowsFormsEditorService = provider.GetService(GetType(IWindowsFormsEditorService))
        'Dim sre As New ShellRendererEditor
        'sre.propRend.SelectedObject = value
        'If frmsvr.ShowDialog(sre) = DialogResult.OK Then
        '    Return sre.propRend.SelectedObject
        'Else
        '    Return value
        'End If
    End Function

End Class

<TypeConverter(GetType(ShowPropsTypeConverter))> _
Public Class StatesRenderCollection

    Dim _normal As ShellRenderer

    Public Property Normal() As ShellRenderer
        Get
            Return _normal
        End Get
        Set(ByVal value As ShellRenderer)
            _normal = value
        End Set
    End Property

    Dim _hot As ShellRenderer

    Public Property Hot() As ShellRenderer
        Get
            Return _hot
        End Get
        Set(ByVal value As ShellRenderer)
            _hot = value
        End Set
    End Property

    Dim _down As ShellRenderer

    Public Property Down() As ShellRenderer
        Get
            Return _down
        End Get
        Set(ByVal value As ShellRenderer)
            _down = value
        End Set
    End Property

    Public Sub New(ByVal norm As ShellRenderer, ByVal hot As ShellRenderer, ByVal down As ShellRenderer)
        Me.Normal = norm
        Me.Hot = hot
        Me.Down = down
    End Sub

    Private Sub New()

    End Sub

    Public Enum RenderStateEnum
        Normal
        Hot
        Down
    End Enum

    Public Overrides Function ToString() As String
        Return "States"
    End Function

    Public Sub RenderState(ByVal gr As Graphics, ByVal bounds As Rectangle, ByVal state As RenderStateEnum)
        Select Case state
            Case RenderStateEnum.Normal
                If Me.Normal IsNot Nothing Then Me.Normal.Render(gr, bounds)
            Case RenderStateEnum.Hot
                If Me.Hot IsNot Nothing Then Me.Hot.Render(gr, bounds)
            Case RenderStateEnum.Down
                If Me.Down IsNot Nothing Then Me.Down.Render(gr, bounds)
        End Select
    End Sub

End Class

Public Class ShowPropsTypeConverter
    Inherits TypeConverter

    Public Overrides Function GetPropertiesSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overrides Function GetProperties(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal value As Object, ByVal attributes() As System.Attribute) As System.ComponentModel.PropertyDescriptorCollection
        Return TypeDescriptor.GetProperties(value)
    End Function

End Class

<TypeConverter(GetType(ShowPropsTypeConverter))> _
Public Class StatesColorCollection

    Dim _normal As Color

    <Editor(GetType(ColorEditorEx), GetType(Drawing.Design.UITypeEditor))> _
    Public Property Normal() As Color
        Get
            Return _normal
        End Get
        Set(ByVal value As Color)
            _normal = value
        End Set
    End Property

    Dim _hot As Color

    <Editor(GetType(ColorEditorEx), GetType(Drawing.Design.UITypeEditor))> _
    Public Property Hot() As Color
        Get
            Return _hot
        End Get
        Set(ByVal value As Color)
            _hot = value
        End Set
    End Property

    Dim _down As Color

    <Editor(GetType(ColorEditorEx), GetType(Drawing.Design.UITypeEditor))> _
    Public Property Down() As Color
        Get
            Return _down
        End Get
        Set(ByVal value As Color)
            _down = value
        End Set
    End Property

    Public Sub New(ByVal norm As Color, ByVal hot As Color, ByVal down As Color)
        Me.Normal = norm
        Me.Hot = hot
        Me.Down = down
    End Sub

    Private Sub New()

    End Sub

    Public Enum ColorStateEnum
        Normal
        Hot
        Down
    End Enum

    Public Overrides Function ToString() As String
        Return "States"
    End Function

    Public Function GetState(ByVal state As ColorStateEnum) As Color
        Select Case state
            Case ColorStateEnum.Normal
                Return Me.Normal
            Case ColorStateEnum.Hot
                Return Me.Hot
            Case ColorStateEnum.Down
                Return Me.Down
        End Select
    End Function

End Class

<System.ComponentModel.Editor(GetType(FontEditor), GetType(Drawing.Design.UITypeEditor))> _
Public Class Font

    Private _style As FontStyle

    Public Property Style() As FontStyle
        Get
            Return _style
        End Get
        Set(ByVal value As FontStyle)
            _style = value
        End Set
    End Property

    Private _size As Single

    Public Property Size() As Single
        Get
            Return _size
        End Get
        Set(ByVal value As Single)
            _size = value
        End Set
    End Property

    Private _name As String

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Function ToFont() As Drawing.Font
        Return New Drawing.Font(Name, Size, Style)
    End Function

    Public Shared Function FromFont(ByVal f As Drawing.Font) As Font
        Dim f2 As New Font
        f2.Name = f.Name
        f2.Style = f.Style
        f2.Size = f.Size
        Return f2
    End Function

    Public Sub New()

    End Sub

    Public Sub New(ByVal n As String, ByVal s As Single)
        Me.Name = n
        Me.Size = s
    End Sub

    Public Sub New(ByVal n As String, ByVal s As Single, ByVal style As FontStyle)
        Me.Name = n
        Me.Size = s
        Me.Style = style
    End Sub
End Class

Public Class FontEditor
    Inherits Drawing.Design.UITypeEditor

    Public Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object) As Object
        Dim fd As New FontDialog
        fd.AllowScriptChange = False
        fd.ShowColor = False
        fd.ShowApply = False
        fd.Font = CType(value, Font).ToFont
        If fd.ShowDialog = DialogResult.OK Then
            Return Font.FromFont(fd.Font)
        Else
            Return value
        End If
    End Function

    Public Overrides Function GetPaintValueSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return False
    End Function

    Public Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
        Return Drawing.Design.UITypeEditorEditStyle.Modal
    End Function

End Class