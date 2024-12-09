Imports System.Xml.Serialization ' Does XML serializing for a class.
Imports System.Drawing ' Required for storing a Bitmap.
Imports System.IO ' Required for using Memory stream objects.
Imports System.ComponentModel ' Required for conversion of Bitmap objects.

' Set this 'Ruler' class as the root node of any XML file its serialized to.
' added <System.Reflection.ObfuscationAttribute(Feature:="renaming")> to turn of obsfucation for ruler class - as it cause problems
<System.Reflection.ObfuscationAttribute(Feature:="renaming")> <XmlRootAttribute("Ruler", [Namespace]:="", IsNullable:=False)> _
Public Class Ruler

    Implements IDisposable

    Private m_HorizontalBackground, m_VerticalBackground As Bitmap
    Private mLineBoxes, mLineTicksAndSides, mLineLength, mLineMeasuring, mLineMidpoint, mLineThirds, mLineGoldenRatio, mLineTicks, mNumbers, mTransparencyColour As Color

    Public Sub New()
    End Sub

    ' Set this 'DateTimeValue' field to be an attribute of the root node.
    ' Note date type must be "dateTime" in that odd mixture of upper and lower case to work
    <XmlAttributeAttribute(DataType:="dateTime")> _
    Public DateTimeValue As System.DateTime = Now


    ' By NOT specifing any custom Metadata Attributes, fields will be created as an element by default.

    '*******************************************************************************************************
    Public Name As String










    '*******************************************************************************************************
    Public SkinEnabled As Boolean
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineBoxes() As Color
        Get
            Return mLineBoxes
        End Get
        Set(ByVal value As Color)
            mLineBoxes = value
        End Set
    End Property
    Public Property Line_Boxes() As String
        Get
            Return mLineBoxes.ToArgb
        End Get
        Set(ByVal value As String)
            mLineBoxes = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineLength() As Color
        Get
            Return mLineLength
        End Get
        Set(ByVal value As Color)
            mLineLength = value
        End Set
    End Property
    Public Property Line_Length() As String
        Get
            Return mLineLength.ToArgb
        End Get
        Set(ByVal value As String)
            mLineLength = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineMeasuring() As Color
        Get
            Return mLineMeasuring
        End Get
        Set(ByVal value As Color)
            mLineMeasuring = value
        End Set
    End Property
    Public Property Line_Measuring() As String
        Get
            Return mLineMeasuring.ToArgb
        End Get
        Set(ByVal value As String)
            mLineMeasuring = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineMidpoint() As Color
        Get
            Return mLineMidpoint
        End Get
        Set(ByVal value As Color)
            mLineMidpoint = value
        End Set
    End Property
    Public Property Line_Midpoint() As String
        Get
            Return mLineMidpoint.ToArgb
        End Get
        Set(ByVal value As String)
            mLineMidpoint = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineThirds() As Color
        Get
            Return mLineThirds
        End Get
        Set(ByVal value As Color)
            mLineThirds = value
        End Set
    End Property
    Public Property Line_Thirds() As String
        Get
            Return mLineThirds.ToArgb
        End Get
        Set(ByVal value As String)
            mLineThirds = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineGoldenRatio() As Color
        Get
            Return mLineGoldenRatio
        End Get
        Set(ByVal value As Color)
            mLineGoldenRatio = value
        End Set
    End Property
    Public Property Line_GoldenRatio() As String
        Get
            Return mLineGoldenRatio.ToArgb
        End Get
        Set(ByVal value As String)
            mLineGoldenRatio = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property LineTicksAndSides() As Color
        Get
            Return mLineTicks
        End Get
        Set(ByVal value As Color)
            mLineTicks = value
        End Set
    End Property
    Public Property Line_Ticks() As String
        Get
            Return mLineTicks.ToArgb
        End Get
        Set(ByVal value As String)
            mLineTicks = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property Numbers() As Color
        Get
            Return mNumbers
        End Get
        Set(ByVal value As Color)
            mNumbers = value
        End Set
    End Property
    Public Property Line_Numbers() As String
        Get
            Return mNumbers.ToArgb
        End Get
        Set(ByVal value As String)
            mNumbers = Color.FromArgb(value)
        End Set
    End Property
    '*******************************************************************************************************
    <System.Xml.Serialization.XmlIgnore()> _
    Public Property TransparencyColour() As Color
        Get
            Return mTransparencyColour
        End Get
        Set(ByVal value As Color)
            mTransparencyColour = value
        End Set
    End Property
    Public Property Background_TransparencyColour() As String
        Get
            Return mTransparencyColour.ToArgb
        End Get
        Set(ByVal value As String)
            mTransparencyColour = Color.FromArgb(value)
        End Set
    End Property

    '*******************************************************************************************************
    Public Opacity As Single
    '*******************************************************************************************************
    Public Tiled As Boolean
    '********************************************************************************************************
    Public HorizontalRotation As Integer
    '*******************************************************************************************************
    Public VerticalRotation As Integer
    '*******************************************************************************************************
    Public BackgroundImageFileName As String
    '*******************************************************************************************************
    ' Set serialization to IGNORE this property because the 'BackgroundByteArray' property
    ' is used instead to serialize the 'Background' Bitmap as an array of bytes.
    <XmlIgnoreAttribute()> _
    Public Property HorizontalBackground() As Bitmap
        Get
            Return m_HorizontalBackground
        End Get
        Set(ByVal value As Bitmap)
            m_HorizontalBackground = value
        End Set
    End Property
    '********************************************************************************************************
    ' Serializes the 'HorizontalBackground' Bitmap to XML.
    <XmlElementAttribute("HorizontalBackground")> _
    Public Property HorizontalBackgroundByteArray() As Byte()
        Get
            If m_HorizontalBackground IsNot Nothing Then
                Dim BitmapConverter As TypeConverter = TypeDescriptor.GetConverter(m_HorizontalBackground.[GetType]())
                Return DirectCast(BitmapConverter.ConvertTo(m_HorizontalBackground, GetType(Byte())), Byte())
            Else
                Return Nothing
            End If
        End Get

        Set(ByVal value As Byte())
            If value IsNot Nothing Then
                m_HorizontalBackground = New Bitmap(New MemoryStream(value))
            Else
                m_HorizontalBackground = Nothing
            End If
        End Set
    End Property

    '********************************************************************************************************
    ' Set serialization to IGNORE this property because the 'BackgroundByteArray' property
    ' is used instead to serialize the 'Background' Bitmap as an array of bytes.
    <XmlIgnoreAttribute()> _
    Public Property VerticalBackground() As Bitmap
        Get
            Return m_VerticalBackground
        End Get
        Set(ByVal value As Bitmap)
            m_VerticalBackground = value
        End Set
    End Property
    '********************************************************************************************************
    ' Serializes the 'HorizontalBackground' Bitmap to XML.
    <XmlElementAttribute("VerticalBackground")> _
    Public Property VerticalBackgroundByteArray() As Byte()
        Get
            If m_VerticalBackground IsNot Nothing Then
                Dim BitmapConverter As TypeConverter = TypeDescriptor.GetConverter(m_VerticalBackground.[GetType]())
                Return DirectCast(BitmapConverter.ConvertTo(m_VerticalBackground, GetType(Byte())), Byte())
            Else
                Return Nothing
            End If
        End Get

        Set(ByVal value As Byte())
            If value IsNot Nothing Then
                m_VerticalBackground = New Bitmap(New MemoryStream(value))
            Else
                m_VerticalBackground = Nothing
            End If
        End Set
    End Property

    'some additional examples
    ''*******************************************************************************************************
    '' Set serialization to IGNORE this field (ie. not add it to the XML).
    '<XmlIgnoreAttribute()> _
    'Public Example As Boolean

    ''*******************************************************************************************************
    '' Serializes an ArrayList as a "Hobbies" array of XML elements of type string named "Hobby".
    '<XmlArray("Hobbies"), XmlArrayItem("Hobby", GetType(String))> _
    'Public Hobbies As New System.Collections.ArrayList()

    ''*******************************************************************************************************
    '' Serializes an ArrayList as a "EmailAddresses" array of XML elements of custom type EmailAddress named "EmailAddress".
    '<XmlArray("EmailAddresses"), XmlArrayItem("EmailAddress", GetType(EmailAddress))> _
    'Public EmailAddresses As New System.Collections.ArrayList()

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                m_HorizontalBackground.Dispose()
                m_VerticalBackground.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

