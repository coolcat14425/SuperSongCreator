Imports System.IO
Imports System.Text
Imports SuperSongCreator.Core.Storage

Namespace Logging
    Public Class AppLogger
        Public Event AlarmRaised(message As String)

        Public Sub Info(message As String)
            Write("INFO", message)
        End Sub

        Public Sub [Error](message As String, ex As Exception)
            Dim fullMessage As String = $"{message}{Environment.NewLine}{ex}"
            Write("ERROR", fullMessage)
            RaiseEvent AlarmRaised(message)
        End Sub

        Private Sub Write(level As String, message As String)
            Dim logFile As String = Path.Combine(AppPaths.LogsFolder(), $"app-{DateTime.UtcNow:yyyyMMdd}.log")
            Dim line As String = $"{DateTime.UtcNow:O} [{level}] {message}{Environment.NewLine}"
            File.AppendAllText(logFile, line, Encoding.UTF8)
        End Sub
    End Class
End Namespace
