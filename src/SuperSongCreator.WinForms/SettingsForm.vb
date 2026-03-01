Imports SuperSongCreator.Core.Domain

Public Class SettingsForm
    Inherits Form

    Private ReadOnly _songsFolderText As TextBox = New TextBox()
    Private ReadOnly _outputFolderText As TextBox = New TextBox()
    Private ReadOnly _fontNameText As TextBox = New TextBox()
    Private ReadOnly _fontSizeNumeric As NumericUpDown = New NumericUpDown()
    Private ReadOnly _fontColorText As TextBox = New TextBox()
    Private ReadOnly _backgroundColorText As TextBox = New TextBox()

    Public Property CurrentSettings As AppSettings

    Public Sub New(existingSettings As AppSettings)
        CurrentSettings = existingSettings

        Text = "Settings"
        Width = 600
        Height = 360
        StartPosition = FormStartPosition.CenterParent

        Dim table As TableLayoutPanel = New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 8
        }
        table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))
        table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60.0F))

        AddRow(table, 0, "Songs Folder", _songsFolderText)
        AddRow(table, 1, "Output Folder", _outputFolderText)
        AddRow(table, 2, "Slide Font", _fontNameText)
        AddRow(table, 3, "Slide Font Size", _fontSizeNumeric)
        AddRow(table, 4, "Font Color (#RRGGBB)", _fontColorText)
        AddRow(table, 5, "Background Color (#RRGGBB)", _backgroundColorText)

        Dim buttonFlow As FlowLayoutPanel = New FlowLayoutPanel() With {.Dock = DockStyle.Fill}
        Dim okButton As Button = New Button() With {.Text = "Save", .DialogResult = DialogResult.OK}
        Dim cancelButton As Button = New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel}
        AddHandler okButton.Click, AddressOf SaveAndClose
        buttonFlow.Controls.Add(okButton)
        buttonFlow.Controls.Add(cancelButton)
        table.Controls.Add(buttonFlow, 1, 7)

        Controls.Add(table)

        _fontSizeNumeric.Minimum = 12
        _fontSizeNumeric.Maximum = 120

        LoadFromSettings()
    End Sub

    Private Shared Sub AddRow(table As TableLayoutPanel, row As Integer, labelText As String, editor As Control)
        table.Controls.Add(New Label() With {.Text = labelText, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, row)
        editor.Dock = DockStyle.Fill
        table.Controls.Add(editor, 1, row)
    End Sub

    Private Sub LoadFromSettings()
        _songsFolderText.Text = CurrentSettings.SongsFolder
        _outputFolderText.Text = CurrentSettings.OutputFolder
        _fontNameText.Text = CurrentSettings.SlideStyle.FontName
        _fontSizeNumeric.Value = CurrentSettings.SlideStyle.FontSizePoints
        _fontColorText.Text = CurrentSettings.SlideStyle.FontColorHtml
        _backgroundColorText.Text = CurrentSettings.SlideStyle.BackgroundColorHtml
    End Sub

    Private Sub SaveAndClose(sender As Object, e As EventArgs)
        CurrentSettings.SongsFolder = _songsFolderText.Text.Trim()
        CurrentSettings.OutputFolder = _outputFolderText.Text.Trim()
        CurrentSettings.SlideStyle.FontName = _fontNameText.Text.Trim()
        CurrentSettings.SlideStyle.FontSizePoints = CInt(_fontSizeNumeric.Value)
        CurrentSettings.SlideStyle.FontColorHtml = _fontColorText.Text.Trim()
        CurrentSettings.SlideStyle.BackgroundColorHtml = _backgroundColorText.Text.Trim()

        If CurrentSettings.TitleStyle Is Nothing Then
            CurrentSettings.TitleStyle = New TitleStyle()
        End If

        CurrentSettings.TitleStyle.FontName = CurrentSettings.SlideStyle.FontName
        CurrentSettings.TitleStyle.FontColorHtml = CurrentSettings.SlideStyle.FontColorHtml
        CurrentSettings.TitleStyle.BackgroundColorHtml = CurrentSettings.SlideStyle.BackgroundColorHtml
        CurrentSettings.TitleStyle.FontSizePoints = CurrentSettings.SlideStyle.FontSizePoints + 4

        DialogResult = DialogResult.OK
        Close()
    End Sub
End Class