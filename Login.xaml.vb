

Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.Remoting


Class Login



    Private Sub Login_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

        WinColl.Add(Me, Me.Name)
        iCurrentView = New CurrentView(New LoginViewBaseViewModel)
        Dim iLoginWindowViewModel As New MainWindowViewModel
        Me.DataContext = iLoginWindowViewModel
    End Sub


    Private Sub Login_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        WinColl.Remove(Me.Name)
    End Sub

    Public Sub OpenWin()
        Me.Show()
    End Sub
    Public Sub CloseWin()
        Me.Close()
    End Sub

    Private Sub btn_Click(sender As Object, e As RoutedEventArgs) Handles btn.Click
        Dim OH As ObjectHandle
        Dim obj As Object
        obj = Activator.CreateInstance(vbNullString, "POS1.aclsDataObject")

        Dim obj2 As Object
        obj2 = obj.Unwrap
        obj2.GetNewMatchObj()
    End Sub
End Class

