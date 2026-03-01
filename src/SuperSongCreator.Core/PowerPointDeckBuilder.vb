Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports A = DocumentFormat.OpenXml.Drawing
Imports SuperSongCreator.Core.Domain
Imports System.IO

Namespace PowerPoint
    Public Class PowerPointDeckBuilder
        Private Const SlideWidthEmu As Int32 = 12192000
        Private Const SlideHeightEmu As Int32 = 6858000

        Public Sub BuildDeck(songs As IEnumerable(Of Song), settings As AppSettings, outputPath As String)
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath))

            Using presentation As PresentationDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation)
                Dim presentationPart As PresentationPart = presentation.AddPresentationPart()
                presentationPart.Presentation = New Presentation()

                Dim slideMasterPart As SlideMasterPart = presentationPart.AddNewPart(Of SlideMasterPart)("rId1")
                slideMasterPart.SlideMaster = BuildSlideMaster()

                Dim slideLayoutPart As SlideLayoutPart = slideMasterPart.AddNewPart(Of SlideLayoutPart)("rId1")
                slideLayoutPart.SlideLayout = BuildSlideLayout()
                slideLayoutPart.AddPart(slideMasterPart)

                slideMasterPart.SlideMaster.SlideLayoutIdList = New SlideLayoutIdList(
                    New SlideLayoutId() With {.Id = 1UI, .RelationshipId = "rId1"}
                )
                slideMasterPart.SlideMaster.Save()

                presentationPart.Presentation.SlideMasterIdList = New SlideMasterIdList(
                    New SlideMasterId() With {.Id = 2147483648UI, .RelationshipId = "rId1"}
                )
                presentationPart.Presentation.SlideIdList = New SlideIdList()

                Dim slideId As UInteger = 256UI
                For Each song As Song In songs
                    Dim titleText As String = song.Title
                    Dim subtitle As String = $"{song.Credits} | CCLI {song.CcliNumber}".Trim()

                    AddSlide(presentationPart, slideLayoutPart, slideId, titleText & Environment.NewLine & subtitle, String.Empty, settings.TitleStyle)
                    slideId = 1UI

                    For Each lyricSlide As SongSlide In song.Slides
                        Dim body As String = String.Join(Environment.NewLine, lyricSlide.TextLines)
                        Dim notes As String = String.Join(Environment.NewLine, lyricSlide.NotesLines)
                        AddSlide(presentationPart, slideLayoutPart, slideId, body, notes, settings.SlideStyle)
                        slideId = 1UI
                    Next

                    If song.Copyright.Length > 0 Then
                        AddSlide(presentationPart, slideLayoutPart, slideId, song.Copyright, "Copyright", settings.SlideStyle)
                        slideId = 1UI
                    End If
                Next

                presentationPart.Presentation.SlideSize = New SlideSize() With {
                    .Cx = CUInt(SlideWidthEmu),
                    .Cy = CUInt(SlideHeightEmu),
                    .Type = SlideSizeValues.Screen4x3
                }
                presentationPart.Presentation.Save()
            End Using
        End Sub

        Private Shared Sub AddSlide(presentationPart As PresentationPart, slideLayoutPart As SlideLayoutPart, slideId As UInteger, slideText As String, notesText As String, style As Object)
            Dim relId As String = $"rIdSlide{slideId}"
            Dim slidePart As SlidePart = presentationPart.AddNewPart(Of SlidePart)(relId)
            slidePart.AddPart(slideLayoutPart)
            slidePart.Slide = BuildSlideContent(slideText, style)
            slidePart.Slide.Save()

            If notesText.Length > 0 Then
                AddNotes(slidePart, notesText)
            End If

            Dim idList As SlideIdList = presentationPart.Presentation.SlideIdList
            idList.Append(New SlideId() With {.Id = slideId, .RelationshipId = relId})
        End Sub

        Private Shared Function BuildSlideContent(text As String, style As Object) As Slide
            Dim fontName As String
            Dim fontSize As Integer
            Dim fontColor As String
            Dim backgroundColor As String

            If TypeOf style Is SlideStyle Then
                Dim s As SlideStyle = CType(style, SlideStyle)
                fontName = s.FontName
                fontSize = s.FontSizePoints
                fontColor = s.FontColorHtml.TrimStart("#"c)
                backgroundColor = s.BackgroundColorHtml.TrimStart("#"c)
            Else
                'Dim s As TitleStyle = CType(style, TitleStyle)
                'fontName = s.FontName
                'fontSize = s.FontSizePoints
                'fontColor = s.FontColorHtml.TrimStart("#"c)
                'backgroundColor = s.BackgroundColorHtml.TrimStart("#"c)
            End If

            Dim slide As Slide = New Slide(
                New CommonSlideData(
                    New ShapeTree(
                        New NonVisualGroupShapeProperties(
                            New NonVisualDrawingProperties() With {.Id = 1UI, .Name = ""},
                            New NonVisualGroupShapeDrawingProperties(),
                            New ApplicationNonVisualDrawingProperties()
                        ),
                        New GroupShapeProperties(New A.TransformGroup()),
                        New Shape(
                            New NonVisualShapeProperties(
                                New NonVisualDrawingProperties() With {.Id = 2UI, .Name = "TextBox"},
                                New NonVisualShapeDrawingProperties(New A.ShapeLocks() With {.NoGrouping = True}),
                                New ApplicationNonVisualDrawingProperties()
                            ),
                            New ShapeProperties(
                                New A.Transform2D(
                                    New A.Offset() With {.X = 0L, .Y = 0L},
                                    New A.Extents() With {.Cx = SlideWidthEmu, .Cy = SlideHeightEmu}
                                )
                            ),
                            New TextBody(
                                New A.BodyProperties() With {.Anchor = A.TextAnchoringTypeValues.Center},
                                New A.ListStyle(),
                                BuildParagraph(text, fontName, fontSize, fontColor)
                            )
                        )
                    )
                ),
                New ColorMapOverride(New A.MasterColorMapping())
            )

            slide.CommonSlideData.Background = New Background(
                New BackgroundStyleReference() With {.Index = 1001UI}
            )

            slide.CommonSlideData.Background.BackgroundProperties = New BackgroundProperties(
                New A.SolidFill(New A.RgbColorModelHex() With {.Val = backgroundColor})
            )

            Return slide
        End Function

        Private Shared Function BuildParagraph(text As String, fontName As String, fontSizePoints As Integer, fontColor As String) As A.Paragraph
            Dim paragraph As A.Paragraph = New A.Paragraph(
                New A.ParagraphProperties() With {
                    .Alignment = A.TextAlignmentTypeValues.Center
                }
            )

            For Each line As String In text.Split({Environment.NewLine}, StringSplitOptions.None)
                paragraph.Append(New A.Run(
                    New A.RunProperties() With {
                        .Language = "en-US",
                        .FontSize = fontSizePoints * 100,
                        .Bold = True,
                        .Dirty = False
                    },
                    New A.RunProperties(
                        New A.SolidFill(New A.RgbColorModelHex() With {.Val = fontColor}),
                        New A.LatinFont() With {.Typeface = fontName}
                    ),
                    New A.Text(line)
                ))
                paragraph.Append(New A.Break())
            Next

            paragraph.Append(New A.EndParagraphRunProperties() With {.Language = "en-US"})
            Return paragraph
        End Function

        Private Shared Sub AddNotes(slidePart As SlidePart, notesText As String)
            Dim notesSlidePart As NotesSlidePart = slidePart.AddNewPart(Of NotesSlidePart)()
            Dim notesMasterPart As NotesMasterPart = slidePart.AddNewPart(Of NotesMasterPart)()
            notesMasterPart.NotesMaster = New NotesMaster(
                New CommonSlideData(New ShapeTree(
                    New NonVisualGroupShapeProperties(
                        New NonVisualDrawingProperties() With {.Id = 1UI, .Name = String.Empty},
                        New NonVisualGroupShapeDrawingProperties(),
                        New ApplicationNonVisualDrawingProperties()),
                    New GroupShapeProperties(New A.TransformGroup())))
            )
            notesMasterPart.NotesMaster.Save()

            notesSlidePart.AddPart(notesMasterPart)
            notesSlidePart.NotesSlide = New NotesSlide(
                New CommonSlideData(
                    New ShapeTree(
                        New NonVisualGroupShapeProperties(
                            New NonVisualDrawingProperties() With {.Id = 1UI, .Name = String.Empty},
                            New NonVisualGroupShapeDrawingProperties(),
                            New ApplicationNonVisualDrawingProperties()
                        ),
                        New GroupShapeProperties(New A.TransformGroup()),
                        New Shape(
                            New NonVisualShapeProperties(
                                New NonVisualDrawingProperties() With {.Id = 2UI, .Name = "Notes"},
                                New NonVisualShapeDrawingProperties(),
                                New ApplicationNonVisualDrawingProperties()
                            ),
                            New ShapeProperties(),
                            New TextBody(
                                New A.BodyProperties(),
                                New A.ListStyle(),
                                New A.Paragraph(New A.Run(New A.Text(notesText)))
                            )
                        )
                    )
                )
            )
            notesSlidePart.NotesSlide.Save()
        End Sub

        Private Shared Function BuildSlideLayout() As SlideLayout
            Return New SlideLayout(
                New CommonSlideData(New ShapeTree(
                    New NonVisualGroupShapeProperties(
                        New NonVisualDrawingProperties() With {.Id = 1UI, .Name = String.Empty},
                        New NonVisualGroupShapeDrawingProperties(),
                        New ApplicationNonVisualDrawingProperties()),
                    New GroupShapeProperties(New A.TransformGroup())
                )),
                New ColorMapOverride(New A.MasterColorMapping())
            )
        End Function

        Private Shared Function BuildSlideMaster() As SlideMaster
            Return New SlideMaster(
                New CommonSlideData(New ShapeTree(
                    New NonVisualGroupShapeProperties(
                        New NonVisualDrawingProperties() With {.Id = 1UI, .Name = String.Empty},
                        New NonVisualGroupShapeDrawingProperties(),
                        New ApplicationNonVisualDrawingProperties()),
                    New GroupShapeProperties(New A.TransformGroup())
                )),
                New ColorMap() With {
                    .Background1 = A.ColorSchemeIndexValues.Light1,
                    .Text1 = A.ColorSchemeIndexValues.Dark1,
                    .Background2 = A.ColorSchemeIndexValues.Light2,
                    .Text2 = A.ColorSchemeIndexValues.Dark2,
                    .Accent1 = A.ColorSchemeIndexValues.Accent1,
                    .Accent2 = A.ColorSchemeIndexValues.Accent2,
                    .Accent3 = A.ColorSchemeIndexValues.Accent3,
                    .Accent4 = A.ColorSchemeIndexValues.Accent4,
                    .Accent5 = A.ColorSchemeIndexValues.Accent5,
                    .Accent6 = A.ColorSchemeIndexValues.Accent6,
                    .Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                    .FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                }
            )
        End Function
    End Class
End Namespace
