Imports System.Collections.ObjectModel
Imports Microsoft.Vbe.Interop.Forms

Public Class vmCheckView

    Inherits ViewModelBase


    Private pCollCheckData As ObservableCollection(Of Object)
    Private pButtonClickCommand As ButtonClickCommand
    Private pRecalledCheckLines As Collection

    Public Property CollCheckData As ObservableCollection(Of Object)
        Set(value As ObservableCollection(Of Object))
            pCollCheckData = value
            OnPropertyChanged(NameOf(CollCheckData))
        End Set
        Get
            Return pCollCheckData
        End Get
    End Property

    Public Property ButtonClickCommand As ButtonClickCommand
        Set(value As ButtonClickCommand)
            pButtonClickCommand = value
        End Set
        Get
            Return pButtonClickCommand
        End Get
    End Property
    Public Property RecalledCheckLines As Collection
        Set(value As Collection)
            pRecalledCheckLines = value
            OnPropertyChanged(NameOf(RecalledCheckLines))
        End Set
        Get
            Return pRecalledCheckLines
        End Get
    End Property

    Public Sub New()
        Me.CollCheckData = New ObservableCollection(Of Object)
        Me.RecalledCheckLines = IncorportateSeatLines(RecallCheckLines(iState.CurrentCheck))
        For Each member As Object In RecalledCheckLines()
            CollCheckData.Add(member)
        Next

        'Me.CollCheckData.Add(RecallCheckLines(CurrentCheck))
        Me.ButtonClickCommand = New ButtonClickCommand(Me)

    End Sub

    Private Function IncorportateSeatLines(CheckLines As Collection) As Collection
        Dim i As Integer, seat As Integer
        Dim coll As New Collection

        If CheckLines.Count = 0 Then
            Return CheckLines
            Exit Function
        End If


        seat = CheckLines("Line1").seat
        Dim initialseatarray(0 To 3) As String
        initialseatarray(1) = "- - - - - - -"
        initialseatarray(2) = "- - - - - - - Seat " & seat & "- - - - - - - - -"
        initialseatarray(3) = "- - - - - - -"
        coll.Add(New zclsCheckLines(initialseatarray),,, coll.Count)
        For i = 1 To CheckLines.Count
            Dim seatarray(0 To 3) As String
            seatarray(1) = "- - - - - - -"
            seatarray(3) = "- - - - - - -"


            If Not CheckLines("Line" & i).seat = seat Then
                seat = CheckLines("Line" & i).seat
                seatarray(2) = "- - - - - - - Seat " & seat & "- - - - - - - - -"
                coll.Add(New zclsCheckLines(seatarray),,, coll.Count)

            End If
            coll.Add(CheckLines("Line" & i),,, coll.Count)
        Next i

        Return coll
    End Function

    Public Sub ButtonClickCommandEvent()
        CollCheckData.RemoveAt(1)
    End Sub



    'Public Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
    '    Throw New NotImplementedException()
    'End Sub
End Class
