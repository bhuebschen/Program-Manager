Imports System.Drawing.Design, System.Drawing.Drawing2D, System.Windows.Forms.Design, System.Reflection, System.ComponentModel

Public Class ImageMarginRenderer
    Inherits ShellRenderer

    Public Overrides Function ToString() As String
        Return "Image with Margins"
    End Function

    Public Overrides Sub Render(ByVal dest As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle)
        If Image IsNot Nothing Then
            'TODO: Draw the upper-left corner

            'TODO: Draw the left side of the image

            'TODO: Draw the top of the image

            'TODO: Draw the right side of the image

            'TODO: Draw the bottom side of the image

            'TODO: Draw the upper-right of the image

            'TODO: Draw the lower-left corner of the image

            'TODO: Draw the lower-right corner of the image

            'TODO: Draw the center of the image
            If Margins.Top + Margins.Bottom < Image.Height AndAlso Margins.Left + Margins.Right < Image.Width Then
                dest.DrawImage(Image, New Rectangle(Margins.Left, Margins.Top, bounds.Width - Margins.Left - Margins.Right, bounds.Height - Margins.Top - Margins.Bottom), New Rectangle(Margins.Left, Margins.Top, Image.Width - Margins.Left - Margins.Right, Image.Height - Margins.Top - Margins.Bottom), GraphicsUnit.Pixel)
            End If
        End If
    End Sub

    Dim _image As Image

    Public Property Image() As Image
        Get
            Return _image
        End Get
        Set(ByVal value As Image)
            _image = value
        End Set
    End Property

    Dim _m As New Padding(3, 3, 3, 3)

    Public Property Margins() As Padding
        Get
            Return _m
        End Get
        Set(ByVal value As Padding)
            _m = value
        End Set
    End Property

    Public Enum CenterModeEnum
        Stretch
        Tile
    End Enum

    Dim _mode As CenterModeEnum

    Public Property Center_Mode() As CenterModeEnum
        Get
            Return _mode
        End Get
        Set(ByVal value As CenterModeEnum)
            _mode = value
        End Set
    End Property

End Class

Public Class TextureRenderer
    Inherits ShellRenderer

    Public Overrides Function ToString() As String
        Return "Texture"
    End Function

    Public Overrides Sub Render(ByVal dest As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle)
        If Image IsNot Nothing Then
            Using t As New TextureBrush(Image, Wrap_Mode)
                t.RotateTransform(Rotation)
                t.TranslateTransform(Offset.X, Offset.Y)
                dest.FillRectangle(t, bounds)
            End Using
        End If
    End Sub

    Dim _img As Image

    ''' <summary>
    ''' The image to use as the texture.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("The image to use as the texture.")> _
    Public Property Image() As Image
        Get
            Return _img
        End Get
        Set(ByVal value As Image)
            _img = value
        End Set
    End Property

    Dim _wrap As Drawing2D.WrapMode

    ''' <summary>
    ''' How to handle textures smaller than the painted area.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("How to handle textures smaller than the painted area.")> _
    Public Property Wrap_Mode() As Drawing2D.WrapMode
        Get
            Return _wrap
        End Get
        Set(ByVal value As Drawing2D.WrapMode)
            _wrap = value
        End Set
    End Property

    Dim _offset As Point

    ''' <summary>
    ''' The offset of the texture from the upper left corner in pixels.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("The offset of the texture from the upper left corner in pixels.")> _
    Public Property Offset() As Point
        Get
            Return _offset
        End Get
        Set(ByVal value As Point)
            _offset = value
        End Set
    End Property

    Dim _rotate As Double

    ''' <summary>
    ''' The rotation of the texture in degrees.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("The rotation of the texture in degrees.")> _
    Public Property Rotation() As Double
        Get
            Return _rotate
        End Get
        Set(ByVal value As Double)
            _rotate = value
        End Set
    End Property

End Class

Public Class HatchRenderer
    Inherits ShellRenderer

    Public Overrides Function ToString() As String
        Return "Hatch Pattern"
    End Function

    Public Overrides Sub Render(ByVal dest As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle)
        dest.FillRectangle(New HatchBrush(_style, _fore, _bck), bounds)
    End Sub

    Dim _fore As Color = Color.Black

    <Editor(GetType(ColorEditorEx), GetType(UITypeEditor))> _
    Public Property Fore_Color() As Color
        Get
            Return _fore
        End Get
        Set(ByVal value As Color)
            _fore = value
        End Set
    End Property

    Dim _bck As Color = Color.White

    <Editor(GetType(ColorEditorEx), GetType(UITypeEditor))> _
    Public Property Back_Color() As Color
        Get
            Return _bck
        End Get
        Set(ByVal value As Color)
            _bck = value
        End Set
    End Property

    Dim _style As HatchStyle

    Public Property Hatch_Style() As HatchStyle
        Get
            Return _style
        End Get
        Set(ByVal value As HatchStyle)
            _style = value
        End Set
    End Property

End Class

Public Class ColorRenderer
    Inherits ShellRenderer

    Public Overrides Function ToString() As String
        Return "Color"
    End Function

    Public Overrides Sub Render(ByVal dest As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle)
        dest.FillRectangle(New SolidBrush(_clr), bounds)
    End Sub

    Dim _clr As Color = Drawing.Color.White

    <Editor(GetType(ColorEditorEx), GetType(UITypeEditor))> _
    Public Property Color() As Color
        Get
            Return _clr
        End Get
        Set(ByVal value As Color)
            _clr = value
        End Set
    End Property

End Class

Public Class LinearGradientRenderer
    Inherits ShellRenderer

    Public Overrides Function ToString() As String
        Return "Linear Gradient"
    End Function

    Public Overrides Sub Render(ByVal dest As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle)
        Using l As New Drawing2D.LinearGradientBrush(bounds, Color_1, Color_2, Direction)
            dest.FillRectangle(l, bounds)
        End Using
    End Sub

    Dim _color1 As Color = Color.White

    <Editor(GetType(ColorEditorEx), GetType(UITypeEditor))> _
    Public Property Color_1() As Color
        Get
            Return _color1
        End Get
        Set(ByVal value As Color)
            _color1 = value
        End Set
    End Property

    Dim _color2 As Color = Color.Black

    <Editor(GetType(ColorEditorEx), GetType(UITypeEditor))> _
    Public Property Color_2() As Color
        Get
            Return _color2
        End Get
        Set(ByVal value As Color)
            _color2 = value
        End Set
    End Property

    Dim _dir As Drawing2D.LinearGradientMode

    Public Property Direction() As Drawing2D.LinearGradientMode
        Get
            Return _dir
        End Get
        Set(ByVal value As Drawing2D.LinearGradientMode)
            _dir = value
        End Set
    End Property

End Class

Public Class VisualStyleRenderer
    Inherits ShellRenderer

    Public Overrides Function ToString() As String
        Return "Visual Style Element"
    End Function

    Public Overrides Sub Render(ByVal dest As System.Drawing.Graphics, ByVal bounds As System.Drawing.Rectangle)
        'Try

        'Catch ex As Exception

        'End Try
        'If VisualStyles.VisualStyleRenderer.IsElementDefined(_ele) Then
        '    Dim r As New VisualStyles.VisualStyleRenderer(_ele)
        '    r.DrawBackground(dest, bounds)
        'End If
    End Sub

    Dim _ele As VisualStyles.VisualStyleElement

    <System.ComponentModel.Editor(GetType(VisualStyleSelector), GetType(UITypeEditor)), Xml.Serialization.XmlIgnore()> _
    Public Property Element() As VisualStyles.VisualStyleElement
        Get
            Return _ele
        End Get
        Set(ByVal value As VisualStyles.VisualStyleElement)
            _ele = value
        End Set
    End Property

    <Browsable(False)> _
    Public Property Part() As Integer
        Get
            Return Element.Part
        End Get
        Set(ByVal value As Integer)
            Element = VisualStyles.VisualStyleElement.CreateElement(Classname, value, State)
        End Set
    End Property

    <Browsable(False)> _
    Public Property Classname() As String
        Get
            Return Element.ClassName
        End Get
        Set(ByVal value As String)
            Element = VisualStyles.VisualStyleElement.CreateElement(value, Part, State)
        End Set
    End Property

    <Browsable(False)> _
    Public Property State() As Integer
        Get
            Return Element.State
        End Get
        Set(ByVal value As Integer)
            Element = VisualStyles.VisualStyleElement.CreateElement(Classname, Part, value)
        End Set
    End Property

    Public Sub New()
        _ele = VisualStyles.VisualStyleElement.Button.PushButton.Default
    End Sub

    Public Sub New(ByVal e As VisualStyles.VisualStyleElement)
        _ele = e
    End Sub

End Class

Public Class VisualStyleSelector
    Inherits Drawing.Design.UITypeEditor

    Public Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
        Return Drawing.Design.UITypeEditorEditStyle.DropDown
    End Function

    Public Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object) As Object
        Dim frmsvr As IWindowsFormsEditorService = provider.GetService(GetType(IWindowsFormsEditorService))
        Dim vse As New VisualStyleEditor(value)
        vse.Width = 200
        vse.Height = 300
        frmsvr.DropDownControl(vse)
        Return vse.SelectedVisualStyleElement
    End Function

    Public Overrides ReadOnly Property IsDropDownResizable() As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return "Visual Style Element"
    End Function

    Public Overrides Function GetPaintValueSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overrides Sub PaintValue(ByVal e As System.Drawing.Design.PaintValueEventArgs)
        MyBase.PaintValue(e)
        If e.Value IsNot Nothing AndAlso TypeOf (e.Value) Is VisualStyles.VisualStyleElement Then
            Dim i As VisualStyles.VisualStyleElement = e.Value
            If VisualStyles.VisualStyleRenderer.IsElementDefined(i) Then
                Dim r As New VisualStyles.VisualStyleRenderer(i)
                r.DrawBackground(e.Graphics, e.Bounds)
            End If
        End If
    End Sub

End Class

Public Class VisualStyleDescriptor
    Inherits TypeDescriptionProvider

    Public Overrides Function GetFullComponentName(ByVal component As Object) As String
        If component IsNot Nothing AndAlso TypeOf (component) Is VisualStyles.VisualStyleElement Then
            Return String.Format("{0} Part={1} State={2}", CType(component, VisualStyles.VisualStyleElement).ClassName, CType(component, VisualStyles.VisualStyleElement).Part, CType(component, VisualStyles.VisualStyleElement).State)
        Else
            Return MyBase.GetFullComponentName(component)
        End If
    End Function

End Class

Public Class VisualStyleEditor
    Inherits TreeView

    Public Sub New(ByVal e As VisualStyles.VisualStyleElement)
        _ele = e

        'Populate the tree
        Dim t As Type = GetType(VisualStyles.VisualStyleElement)
        Dim root As New TreeNode("Visual Style Elements")
        root.Tag = t
        root.ForeColor = SystemColors.GrayText
        Me.Nodes.Add(root)
        AddNestedTypes(t, root)
        root.Expand()

        If enode IsNot Nothing Then Me.SelectedNode = enode
    End Sub

    Dim enode As TreeNode

    Private Sub AddNestedTypes(ByVal t As Type, ByVal n As TreeNode)
        For Each i As Type In t.GetNestedTypes
            Dim j As New TreeNode(i.Name)
            j.Tag = i
            j.ForeColor = SystemColors.GrayText
            n.Nodes.Add(j)
            AddNestedTypes(i, j)
            For Each p As Reflection.PropertyInfo In i.GetProperties
                If GetType(VisualStyles.VisualStyleElement).IsAssignableFrom(p.PropertyType) Then
                    Dim k As New TreeNode(p.Name)
                    k.Tag = p.GetValue(Nothing, New Object() {})
                    If AreElementsEqual(_ele, k.Tag) Then enode = k
                    j.Nodes.Add(k)
                End If
            Next
        Next
    End Sub

    Private Function AreElementsEqual(ByVal e1 As VisualStyles.VisualStyleElement, ByVal e2 As VisualStyles.VisualStyleElement) As Boolean
        Return e1.ClassName = e2.ClassName AndAlso e1.Part = e2.Part AndAlso e1.State = e2.State
    End Function

    Protected Overrides Sub OnAfterSelect(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        MyBase.OnAfterSelect(e)
        Try
            If e.Node.Tag IsNot Nothing AndAlso TypeOf (e.Node.Tag) Is VisualStyles.VisualStyleElement Then
                _ele = CType(e.Node.Tag, VisualStyles.VisualStyleElement)
                Me.FindForm.Close()
            End If
        Catch ex As Exception
            Debug.Print("Error selecting item.")
        End Try
    End Sub

    Dim _ele As VisualStyles.VisualStyleElement

    Public ReadOnly Property SelectedVisualStyleElement() As VisualStyles.VisualStyleElement
        Get
            Return _ele
        End Get
    End Property

End Class

''' <summary>
''' Class extending the <see cref="ColorEditor"/> which adds the
''' capability to also change the alpha value of the color.
''' </summary>
Public Class ColorEditorEx
    Inherits ColorEditor

#Region "Class ColorUIWrapper"

    ''' <summary>
    ''' Wrapper for the private ColorUI class nested within <see cref="ColorEditor"/>.
    ''' It publishes its internals via reflection and adds a <see cref="TrackBar"/> to
    ''' adjust teh alpha value.
    ''' </summary>
    Public Class ColorUIWrapper

#Region "Fields"

        Private _control As Control
        Private _startMethodInfo As MethodInfo
        Private _endMethodInfo As MethodInfo
        Private _valuePropertyInfo As PropertyInfo
        Private _tbAlpha As TrackBar
        Private _lblAlpha As Label
        Private _inSizeChange As Boolean = False

#End Region

#Region "Constructors"

        ''' <summary>
        ''' Creates a new instance.
        ''' </summary>
        ''' <param name="colorEditor">The editor this instance belongs to.</param>
        Public Sub New(ByVal colorEditor As ColorEditorEx)
            Dim colorUiType As Type = GetType(ColorEditor).GetNestedType("ColorUI", BindingFlags.CreateInstance Or BindingFlags.NonPublic)
            Dim constructorInfo As ConstructorInfo = colorUiType.GetConstructor(New Type() {GetType(ColorEditor)})
            _control = constructorInfo.Invoke(New Object() {colorEditor})

            Dim alphaPanel As Panel = New Panel()
            alphaPanel.BackColor = SystemColors.Control
            alphaPanel.Dock = DockStyle.Right
            alphaPanel.Width = 28
            _control.Controls.Add(alphaPanel)
            _tbAlpha = New TrackBar()
            _tbAlpha.Orientation = Orientation.Vertical
            _tbAlpha.Dock = DockStyle.Fill
            _tbAlpha.TickStyle = TickStyle.None
            _tbAlpha.Maximum = Byte.MaxValue
            _tbAlpha.Minimum = Byte.MinValue
            AddHandler _tbAlpha.ValueChanged, AddressOf OnTrackBarAlphaValueChanged
            alphaPanel.Controls.Add(_tbAlpha)
            _lblAlpha = New Label()
            _lblAlpha.Text = "0"
            _lblAlpha.Dock = DockStyle.Bottom
            _lblAlpha.TextAlign = ContentAlignment.MiddleCenter
            alphaPanel.Controls.Add(_lblAlpha)

            _startMethodInfo = _control.GetType().GetMethod("Start")
            _endMethodInfo = _control.GetType().GetMethod("End")
            _valuePropertyInfo = _control.GetType().GetProperty("Value")

            AddHandler _control.SizeChanged, AddressOf OnControlSizeChanged
        End Sub

#End Region

#Region "Public interface"

        ''' <summary>
        ''' The control to be shown when a color is edited.
        ''' The concrete type is ColorUI which is privately hidden
        ''' within System.Drawing.Design.
        ''' </summary>
        Public ReadOnly Property Control() As Control
            Get
                Return _control
            End Get
        End Property

        ''' <summary>
        ''' Gets the edited color with applied alpha value.
        ''' </summary>
        Public ReadOnly Property Value() As Object
            Get
                Dim result As Object = _valuePropertyInfo.GetValue(_control, New Object(-1) {})
                If (TypeOf (result) Is Color) Then result = Color.FromArgb(_tbAlpha.Value, result)
                Return result
            End Get
        End Property

        ''' <summary>
        ''' Starts the editing process.
        ''' </summary>
        ''' <param name="service">The editor service.</param>
        ''' <param name="value">The value to be edited.</param>
        Public Sub Start(ByVal service As IWindowsFormsEditorService, ByVal value As Object)
            If (TypeOf (value) Is Color) Then _tbAlpha.Value = CType(value, Color).A

            _startMethodInfo.Invoke(_control, New Object() {service, value})
        End Sub

        ''' <summary>
        ''' End the editing process.
        ''' </summary>
        Public Sub [End]()
            _endMethodInfo.Invoke(_control, New Object(-1) {})
        End Sub

#End Region

#Region "Privates"

        Private Sub OnControlSizeChanged(ByVal sender As Object, ByVal e As EventArgs)
            If (_inSizeChange) Then Return

            Try
                _inSizeChange = True

                Dim tabControl As TabControl = _control.Controls(0)

                Dim size As Size = tabControl.TabPages(0).Controls(0).Size
                Dim rectangle As Rectangle = tabControl.GetTabRect(0)
                _control.Size = New Size(_tbAlpha.Width + size.Width, size.Height + rectangle.Height)
            Finally
                _inSizeChange = False
            End Try
        End Sub

#End Region

        Private Sub OnTrackBarAlphaValueChanged(ByVal sender As Object, ByVal e As EventArgs)
            _lblAlpha.Text = _tbAlpha.Value.ToString()
        End Sub

    End Class

#End Region

#Region "Fields"

    Private _colorUI As ColorUIWrapper

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new instance.
    ''' </summary>
    Public Sub New()

    End Sub

#End Region

#Region "Overridden from ColorEditor"

    ''' <summary>
    ''' Edits the given value.
    ''' </summary>
    ''' <param name="context">Context infromation.</param>
    ''' <param name="provider">Service provider.</param>
    ''' <param name="value">Value to be edited.</param>
    ''' <returns>An edited value.</returns>
    Public Overrides Function EditValue(ByVal context As ITypeDescriptorContext, ByVal provider As IServiceProvider, ByVal value As Object) As Object
        If (provider IsNot Nothing) Then

            Dim service As IWindowsFormsEditorService = provider.GetService(GetType(IWindowsFormsEditorService))
            If (service Is Nothing) Then Return value

            If (_colorUI Is Nothing) Then _colorUI = New ColorUIWrapper(Me)

            _colorUI.Start(service, value)
            service.DropDownControl(_colorUI.Control)
            If ((_colorUI.Value IsNot Nothing) AndAlso (_colorUI.Value <> Color.Empty)) Then
                value = _colorUI.Value
            End If
            _colorUI.End()
        End If
        Return value
    End Function

    Public Overrides Sub PaintValue(ByVal e As PaintValueEventArgs)
        If (TypeOf (e.Value) Is Color AndAlso CType(e.Value, Color).A < Byte.MaxValue) Then
            Dim oneThird As Integer = e.Bounds.Width / 3
            Using brush As New SolidBrush(Color.White)
                e.Graphics.FillRectangle(brush, New Rectangle(e.Bounds.X, e.Bounds.Y, oneThird, e.Bounds.Height - 1))
            End Using
            Using brush As New SolidBrush(Color.DarkGray)
                e.Graphics.FillRectangle(brush, New Rectangle(e.Bounds.X + oneThird, e.Bounds.Y, oneThird, e.Bounds.Height - 1))
            End Using
            Using brush As New SolidBrush(Color.Black)
                e.Graphics.FillRectangle(brush, New Rectangle(e.Bounds.X + oneThird * 2, e.Bounds.Y, e.Bounds.Width - oneThird * 2, e.Bounds.Height - 1))
            End Using
        End If

        MyBase.PaintValue(e)
    End Sub

#End Region

End Class
