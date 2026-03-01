Imports System.Drawing

Namespace Domain
    Public Class Song
        Public Property Title As String = String.Empty
        Public Property Credits As String = String.Empty
        Public Property CcliNumber As String = String.Empty
        Public Property Copyright As String = String.Empty
        Public Property SourceFile As String = String.Empty
        Public ReadOnly Property Slides As List(Of SongSlide) = New List(Of SongSlide)()
    End Class

    Public Class SongSlide
        Public Property TextLines As List(Of String) = New List(Of String)()
        Public Property NotesLines As List(Of String) = New List(Of String)()
    End Class

    Public Class SlideStyle
        Public Property FontName As String = "Arial"
        Public Property FontColorHtml As String = "#FFFFFF"
        Public Property BackgroundColorHtml As String = "#000000"
        Public Property FontSizePoints As Integer = 42
    End Class

    Public Class TitleStyle
        Public Property FontName As String = "Arial"
        Public Property FontColorHtml As String = "#FFFFFF"
        Public Property BackgroundColorHtml As String = "#000000"
        Public Property FontSizePoints As Integer = 46
    End Class

    Public Class AppSettings
        Public Property SongsFolder As String = String.Empty
        Public Property OutputFolder As String = String.Empty
        Public Property SlideStyle As SlideStyle = New SlideStyle()
        Public Property TitleStyle As TitleStyle = New TitleStyle()
    End Class

    Public Class SongGroup
        Public Property GroupName As String = String.Empty
        Public Property SongFileNames As List(Of String) = New List(Of String)()
    End Class
End Namespace
