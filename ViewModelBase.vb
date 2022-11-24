Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Threading.Tasks
Imports System
Imports System.Linq
Imports System.Collections.ObjectModel

Public MustInherit Class ViewModelBase
    Implements INotifyPropertyChanged

    Private pViewModelInstance As ViewModelBase
    Public iViewModelInstance As ViewModelBase



    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged



    Protected Sub OnPropertyChanged(propertyName As String)

        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub


    Public Property ViewModelInstance() As ViewModelBase
        Set(value As ViewModelBase)
            iViewModelInstance = value
            OnPropertyChanged(NameOf(ViewModelInstance))
        End Set
        Get
            Return iViewModelInstance
        End Get
    End Property


    Public Overridable Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)

    End Sub

End Class
