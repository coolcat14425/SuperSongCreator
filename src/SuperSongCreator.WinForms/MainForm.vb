Imports System.Text.Json
Imports SuperSongCreator.Core.Domain

Namespace Storage
    Public Class AppPaths
        Public Shared Function RootFolder() As String
            Dim appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Return Path.Combine(appData, "SuperSongCreator")
        End Function

        Public Shared Function EnsureRootFolder() As String
            Dim folder As String = RootFolder()
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Public Shared Function SettingsPath() As String
            Return Path.Combine(EnsureRootFolder(), "settings.json")
        End Function

        Public Shared Function GroupsPath() As String
            Return Path.Combine(EnsureRootFolder(), "songgroups.json")
        End Function

        Public Shared Function LogsFolder() As String
            Dim logs As String = Path.Combine(EnsureRootFolder(), "logs")
            Directory.CreateDirectory(logs)
            Return logs
        End Function
    End Class

    Public Class SettingsRepository
        Private Shared ReadOnly Options As JsonSerializerOptions = New JsonSerializerOptions() With {
            .WriteIndented = True
        }

        Public Function LoadSettings() As AppSettings
            Dim path As String = AppPaths.SettingsPath()
            If Not File.Exists(path) Then
                Return New AppSettings()
            End If

            Dim json As String = File.ReadAllText(path)
            Return JsonSerializer.Deserialize(Of AppSettings)(json, Options)
        End Function

        Public Sub SaveSettings(settings As AppSettings)
            Dim json As String = JsonSerializer.Serialize(settings, Options)
            File.WriteAllText(AppPaths.SettingsPath(), json)
        End Sub
    End Class

    Public Class SongGroupRepository
        Private Shared ReadOnly Options As JsonSerializerOptions = New JsonSerializerOptions() With {
            .WriteIndented = True
        }

        Public Function LoadGroups() As List(Of SongGroup)
            Dim path As String = AppPaths.GroupsPath()
            If Not File.Exists(path) Then
                Return New List(Of SongGroup)()
            End If

            Dim json As String = File.ReadAllText(path)
            Return JsonSerializer.Deserialize(Of List(Of SongGroup))(json, Options)
        End Function

        Public Sub SaveGroups(groups As List(Of SongGroup))
            Dim json As String = JsonSerializer.Serialize(groups, Options)
            File.WriteAllText(AppPaths.GroupsPath(), json)
        End Sub
    End Class
End Namespace
