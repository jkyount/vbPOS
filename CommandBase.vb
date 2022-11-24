Imports System.Windows.Input
Imports System.Threading.Tasks

Namespace Commands

    Public MustInherit Class CommandBase : Implements ICommand

        Private pViewModel As Object

        Public Property ViewModel As Object
            Set(value As Object)
                pViewModel = value
            End Set
            Get
                Return pViewModel
            End Get
        End Property

        Public Sub New(ViewModel As Object)
            Me.ViewModel = ViewModel
        End Sub

        Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

        Public MustOverride Sub Execute(parameter As Object) Implements ICommand.Execute



        Public Overridable Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
            Return True
        End Function

        'Public Sub OnCanExecuteChanged()
        '    RaiseEvent CanExecuteChanged(Me, New EventArgs)
        'End Sub
    End Class

End Namespace