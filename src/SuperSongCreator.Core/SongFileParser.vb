Imports System.IO
Imports SuperSongCreator.Core.Domain

Namespace Parsing
    Public Class SongFileParser
        Public Function ParseFile(path As String) As Song
            Dim lines As List(Of String) = File.ReadAllLines(path).
                Select(Function(line) line.TrimEnd()).
                ToList()

            If lines.Count < 3 Then
                Throw New InvalidDataException($"Song file '{path}' must contain at least 3 lines.")
            End If

            Dim song As Song = New Song() With {
                .Title = lines(0).Trim(),
                .Credits = lines(1).Trim(),
                .CcliNumber = lines(2).Trim(),
                .SourceFile = path
            }

            Dim currentSlide As SongSlide = New SongSlide()

            For i As Integer = 3 To lines.Count - 1
                Dim line As String = lines(i).Trim()

                If line.StartsWith("##", StringComparison.Ordinal) Then
                    song.Copyright = line.Substring(2).Trim()
                    Continue For
                End If

                If line.Equals("#", StringComparison.Ordinal) Then
                    AppendSlideIfNotEmpty(song, currentSlide)
                    currentSlide = New SongSlide()
                    Continue For
                End If

                If line.StartsWith("**", StringComparison.Ordinal) Then
                    currentSlide.NotesLines.Add(line.Substring(2).Trim())
                    Continue For
                End If

                If line.Length > 0 Then
                    currentSlide.TextLines.Add(line)
                End If
            Next

            AppendSlideIfNotEmpty(song, currentSlide)

            If song.Slides.Count = 0 Then
                Throw New InvalidDataException($"Song file '{path}' did not contain any slide content.")
            End If

            Return song
        End Function

        Private Shared Sub AppendSlideIfNotEmpty(song As Song, slide As SongSlide)
            If slide.TextLines.Count = 0 AndAlso slide.NotesLines.Count = 0 Then
                Return
            End If

            song.Slides.Add(slide)
        End Sub
    End Class
End Namespace
