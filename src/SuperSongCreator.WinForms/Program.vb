Imports System.Windows.Forms
Imports SuperSongCreator.Core.Logging

Namespace My
    Friend Module Program
        <STAThread>
        Public Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)

            Dim logger As AppLogger = New AppLogger()
            Dim form As MainForm = New MainForm(logger)
            Application.Run(form)
        End Sub
    End Module
End Namespace