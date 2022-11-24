Imports System.ComponentModel

Public Class MainWindow







    Private Sub Close_Click(sender As Object, e As RoutedEventArgs)

        GetWindow(Me).Close()
    End Sub


    Public Sub OpenWin()
        Me.Show()
    End Sub
    Public Sub CloseWin()
        Me.Close()
    End Sub

    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        WinColl.Add(Me, Me.Name)
        iCurrentView = New CurrentView(New vmHomeBase)
        Dim iMainWindowViewModel As New MainWindowViewModel
        Me.DataContext = iMainWindowViewModel
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        WinColl.Remove(Me.Name)
    End Sub
End Class
