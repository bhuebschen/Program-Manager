Imports System.ComponentModel

'<TypeConverter(GetType(ShowPropsTypeConverter))> _
Public MustInherit Class ShellControlConfig
    Implements ICustomTypeDescriptor

    Public Sub New()

    End Sub

    Dim _backgrnd As ShellRenderer = New VisualStyleRenderer

    <Description("Used to render the background of the control.")> _
    Public Property Background() As ShellRenderer
        Get
            Return _backgrnd
        End Get
        Set(ByVal value As ShellRenderer)
            _backgrnd = value
        End Set
    End Property

    Dim _dock As DockStyle

    <Description("A docking style used to control the control's position on it's panel.")> _
    Public Property Dock() As DockStyle
        Get
            Return _dock
        End Get
        Set(ByVal value As DockStyle)
            _dock = value
        End Set
    End Property

    Dim _anchor As AnchorStyles

    <Description("Used to anchor sides of the control the sides of it's shell panel during resizing.")> _
    Public Property Anchor() As AnchorStyles
        Get
            Return _anchor
        End Get
        Set(ByVal value As AnchorStyles)
            _anchor = value
        End Set
    End Property

    Dim _location As Point

    <Description("The location of the control in pixels.")> _
    Public Property Location() As Point
        Get
            Return _location
        End Get
        Set(ByVal value As Point)
            _location = value
        End Set
    End Property

    Dim _size As Size

    <Description("The size of the control in pixels.")> _
    Public Property Size() As Size
        Get
            Return _size
        End Get
        Set(ByVal value As Size)
            _size = value
        End Set
    End Property

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

Public Class ConfigPropertyDescriptor
    Inherits PropertyDescriptor

    Private baseProp As PropertyDescriptor

    Private friendlyName As String

    Public Sub New(ByVal baseProp As PropertyDescriptor, ByVal filter() As Attribute)
        MyBase.New(baseProp)
        Me.baseProp = baseProp

    End Sub

    Public Overrides ReadOnly Property Name() As String
        Get
            Return Me.baseProp.Name.Replace("_", " ")
        End Get
    End Property

    Public Overrides ReadOnly Property DisplayName() As String
        Get
            Return Me.baseProp.DisplayName.Replace("_", " ")
        End Get
    End Property

    Public Overrides ReadOnly Property IsReadOnly() As Boolean
        Get
            If Me.PropertyType.IsInterface Then
                Return True
            Else
                Return baseProp.IsReadOnly
            End If
        End Get
    End Property

    Public Overrides Function CanResetValue(ByVal component As Object) As Boolean
        Return Me.baseProp.CanResetValue(component)
    End Function

    Public Overrides ReadOnly Property ComponentType() As Type
        Get
            Return baseProp.ComponentType
        End Get
    End Property

    Public Overrides Function GetValue(ByVal component As Object) As Object
        Return Me.baseProp.GetValue(component)
    End Function

    Public Overrides ReadOnly Property PropertyType() As Type
        Get
            Return Me.baseProp.PropertyType
        End Get
    End Property

    Public Overrides Sub ResetValue(ByVal component As Object)
        baseProp.ResetValue(component)
    End Sub

    Public Overrides Sub SetValue(ByVal component As Object, ByVal Value As Object)
        Me.baseProp.SetValue(component, Value)
    End Sub

    Public Overrides Function ShouldSerializeValue(ByVal component As Object) As Boolean
        Return Me.baseProp.ShouldSerializeValue(component)
    End Function

    Public Overrides Function GetChildProperties(ByVal instance As Object, ByVal filter() As System.Attribute) As System.ComponentModel.PropertyDescriptorCollection
        Dim oldprops As PropertyDescriptorCollection = TypeDescriptor.GetProperties(instance.GetType, filter)
        Dim newprops(oldprops.Count - 1) As PropertyDescriptor
        For i As Integer = 0 To oldprops.Count - 1
            newprops(i) = New ConfigPropertyDescriptor(oldprops(i), filter)
        Next
        Return New PropertyDescriptorCollection(newprops)
    End Function

End Class