Attribute VB_Name = "FormControl"
Private List() As Control
Private curr_obj As Object
Private iHeight As Integer
Private iWidth As Integer
Private x_size As Double
Private y_size As Double

Private Type Control
    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
End Type
Public Sub ResizeControls(frm As Form)
Dim i As Integer
x_size = frm.height / iHeight
y_size = frm.width / iWidth

For i = 0 To UBound(List)
    For Each curr_obj In frm
        If curr_obj.TabIndex = List(i).Index Then
             With curr_obj
                .Left = List(i).Left * y_size
                .width = List(i).width * y_size
                .height = List(i).height * x_size
                .Top = List(i).Top * x_size
             End With
        End If
    Next curr_obj
Next i
End Sub
Public Function SetFontSize() As Integer
    If Int(x_size) > 0 Then
        SetFontSize = Int(x_size * 8)
    End If
End Function
Public Sub GetLocation(frm As Form)
Dim i As Integer
For Each curr_obj In frm
    ReDim Preserve List(i)
    With List(i)
        .Name = curr_obj
        .Index = curr_obj.TabIndex
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
    End With
    i = i + 1
Next curr_obj

    iHeight = frm.height
    iWidth = frm.width
End Sub
Public Sub CenterForm(frm As Form)
    frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
End Sub
Public Sub ResizeForm(frm As Form)
    frm.height = Screen.height / 2
    frm.width = Screen.width / 2
    ResizeControls frm
End Sub
