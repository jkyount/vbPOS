Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public MustInherit Class IButtonCanvas
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged



    Protected Sub OnPropertyChanged(propertyName As String)

        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

End Class
Public Class vmButtonCanvasBase
    Inherits IButtonCanvas


End Class




Public Class CondensedMenuButtonCanvasViewModel
    Inherits vmButtonCanvasBase



End Class

Public Class ExpandedMenuButtonCanvasViewModel
    Inherits vmButtonCanvasBase


End Class

Public Class BlankMenuButtonCanvasViewModel
    Inherits IButtonCanvas



End Class
