Imports System.Windows.Forms
Imports SuperSongCreator.Core.Domain
Imports SuperSongCreator.Core.Logging
Imports SuperSongCreator.Core.Parsing
Imports SuperSongCreator.Core.PowerPoint
Imports SuperSongCreator.Core.Storage

Public Class MainForm
    Inherits Form

    Private ReadOnly _logger As AppLogger
    Private ReadOnly _parser As SongFileParser = New SongFileParser()
    Private ReadOnly _settingsRepository As SettingsRepository = New SettingsRepository()
    Private ReadOnly _groupRepository As SongGroupRepository = New SongGroupRepository()
    Private ReadOnly _deckBuilder As PowerPointDeckBuilder = New PowerPointDeckBuilder()

    Private _settings As AppSettings = New AppSettings()
    Private _groups As List(Of SongGroup) = New List(Of SongGroup)()

    Private ReadOnly _songsList As ListBox = New ListBox()
    Private ReadOnly _selectedSongsList As ListBox = New ListBox()
    Private ReadOnly _groupsList As ListBox = New ListBox()

    Public Sub New(logger As AppLogger)
        _logger = logger
        AddHandler _logger.AlarmRaised, AddressOf OnAlarmRaised

        Text = "SuperSongCreator"
        Width = 1100
        Height = 700
        StartPosition = FormStartPosition.CenterScreen

        BuildLayout()
        LoadState()
        ReloadSongs()
        ReloadGroups()
    End Sub

    Private Sub BuildLayout()
        Dim root As TableLayoutPanel = New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 3,
            .RowCount = 1
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35.0F))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35.0F))

        Dim songsPanel As GroupBox = New GroupBox() With {.Dock = DockStyle.Fill, .Text = "Songs"}
        _songsList.Dock = DockStyle.Fill
        songsPanel.Controls.Add(_songsList)

        Dim buttonsPanel As FlowLayoutPanel = New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown}
        Dim addButton As Button = New Button() With {.Text = ">", .Width = 80}
        AddHandler addButton.Click, Sub() MoveSelectedSong(addToSelection:=True)

        Dim removeButton As Button = New Button() With {.Text = "<", .Width = 80}
        AddHandler removeButton.Click, Sub() MoveSelectedSong(addToSelection:=False)

        Dim settingsButton As Button = New Button() With {.Text = "Settings", .Width = 100}
        AddHandler settingsButton.Click, AddressOf OpenSettings

        Dim generateButton As Button = New Button() With {.Text = "Generate Slides", .Width = 130}
        AddHandler generateButton.Click, AddressOf GenerateSlides

        buttonsPanel.Controls.Add(addButton)
        buttonsPanel.Controls.Add(removeButton)
        buttonsPanel.Controls.Add(settingsButton)
        buttonsPanel.Controls.Add(generateButton)

        Dim groupsPanel As GroupBox = New GroupBox() With {.Dock = DockStyle.Fill, .Text = "Song Group / Service"}
        Dim groupLayout As TableLayoutPanel = New TableLayoutPanel() With {.Dock = DockStyle.Fill, .RowCount = 3, .ColumnCount = 1}
        groupLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 40.0F))
        groupLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 35.0F))
        groupLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 60.0F))

        _groupsList.Dock = DockStyle.Fill

        Dim groupsButtons As FlowLayoutPanel = New FlowLayoutPanel() With {.Dock = DockStyle.Fill}
        Dim saveGroupButton As Button = New Button() With {.Text = "Save Group"}
        AddHandler saveGroupButton.Click, AddressOf SaveGroup
        Dim openGroupButton As Button = New Button() With {.Text = "Open Group"}
        AddHandler openGroupButton.Click, AddressOf OpenGroup
        groupsButtons.Controls.Add(saveGroupButton)
        groupsButtons.Controls.Add(openGroupButton)

        _selectedSongsList.Dock = DockStyle.Fill

        groupLayout.Controls.Add(_groupsList, 0, 0)
        groupLayout.Controls.Add(groupsButtons, 0, 1)
        groupLayout.Controls.Add(_selectedSongsList, 0, 2)

        groupsPanel.Controls.Add(groupLayout)

        root.Controls.Add(songsPanel, 0, 0)
        root.Controls.Add(buttonsPanel, 1, 0)
        root.Controls.Add(groupsPanel, 2, 0)

        Controls.Add(root)
    End Sub

    Private Sub LoadState()
        Try
            _settings = _settingsRepository.LoadSettings()
            _groups = _groupRepository.LoadGroups()
        Catch ex As Exception
            _logger.Error("Failed to load application state.", ex)
            _settings = New AppSettings()
            _groups = New List(Of SongGroup)()
        End Try
    End Sub

    Private Sub ReloadSongs()
        _songsList.Items.Clear()

        If String.IsNullOrWhiteSpace(_settings.SongsFolder) OrElse Not Directory.Exists(_settings.SongsFolder) Then
            Return
        End If

        For Each filePath As String In Directory.GetFiles(_settings.SongsFolder, "*.txt", SearchOption.TopDirectoryOnly)
            _songsList.Items.Add(Path.GetFileName(filePath))
        Next
    End Sub

    Private Sub ReloadGroups()
        _groupsList.Items.Clear()
        For Each group As SongGroup In _groups
            _groupsList.Items.Add(group.GroupName)
        Next
    End Sub

    Private Sub MoveSelectedSong(addToSelection As Boolean)
        If addToSelection Then
            If _songsList.SelectedItem IsNot Nothing Then
                _selectedSongsList.Items.Add(_songsList.SelectedItem.ToString())
            End If
        Else
            If _selectedSongsList.SelectedItem IsNot Nothing Then
                _selectedSongsList.Items.Remove(_selectedSongsList.SelectedItem)
            End If
        End If
    End Sub

    Private Sub SaveGroup(sender As Object, e As EventArgs)
        Dim name As String = InputBox("Group name:", "Save Song Group")
        If String.IsNullOrWhiteSpace(name) Then
            Return
        End If

        Dim selected As List(Of String) = _selectedSongsList.Items.Cast(Of Object)().Select(Function(item) item.ToString()).ToList()
        Dim existing As SongGroup = _groups.FirstOrDefault(Function(group) group.GroupName.Equals(name, StringComparison.OrdinalIgnoreCase))

        If existing Is Nothing Then
            _groups.Add(New SongGroup() With {.GroupName = name, .SongFileNames = selected})
        Else
            existing.SongFileNames = selected
        End If

        Try
            _groupRepository.SaveGroups(_groups)
            ReloadGroups()
        Catch ex As Exception
            _logger.Error("Failed to save song groups.", ex)
        End Try
    End Sub

    Private Sub OpenGroup(sender As Object, e As EventArgs)
        If _groupsList.SelectedItem Is Nothing Then
            Return
        End If

        Dim selectedName As String = _groupsList.SelectedItem.ToString()
        Dim group As SongGroup = _groups.FirstOrDefault(Function(item) item.GroupName.Equals(selectedName, StringComparison.OrdinalIgnoreCase))
        If group Is Nothing Then
            Return
        End If

        _selectedSongsList.Items.Clear()
        For Each songFile As String In group.SongFileNames
            _selectedSongsList.Items.Add(songFile)
        Next
    End Sub

    Private Sub OpenSettings(sender As Object, e As EventArgs)
        Using settingsDialog As SettingsForm = New SettingsForm(_settings)
            If settingsDialog.ShowDialog(Me) = DialogResult.OK Then
                _settings = settingsDialog.CurrentSettings
                Try
                    _settingsRepository.SaveSettings(_settings)
                    ReloadSongs()
                Catch ex As Exception
                    _logger.Error("Failed to save settings.", ex)
                End Try
            End If
        End Using
    End Sub

    Private Sub GenerateSlides(sender As Object, e As EventArgs)
        Try
            If _selectedSongsList.Items.Count = 0 Then
                MessageBox.Show(Me, "Choose at least one song.", "No Songs", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If String.IsNullOrWhiteSpace(_settings.OutputFolder) Then
                MessageBox.Show(Me, "Set output folder in Settings first.", "Missing Output Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim songs As List(Of Song) = New List(Of Song)()
            For Each item As Object In _selectedSongsList.Items
                Dim filePath As String = Path.Combine(_settings.SongsFolder, item.ToString())
                songs.Add(_parser.ParseFile(filePath))
            Next

            Dim outputPath As String = Path.Combine(_settings.OutputFolder, $"Service-{DateTime.Now:yyyyMMdd-HHmmss}.pptx")
            _deckBuilder.BuildDeck(songs, _settings, outputPath)

            Process.Start(New ProcessStartInfo(outputPath) With {.UseShellExecute = True})
        Catch ex As Exception
            _logger.Error("Slide generation failed.", ex)
        End Try
    End Sub

    Private Sub OnAlarmRaised(message As String)
        If InvokeRequired Then
            Invoke(Sub() OnAlarmRaised(message))
            Return
        End If

        MessageBox.Show(Me, message & Environment.NewLine & "See log file under AppData\\SuperSongCreator\\logs", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub
End Class
