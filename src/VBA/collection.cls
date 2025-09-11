VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim inputCell As Range
    Set inputCell = Me.Range("J2") ' change to your input cell

    ' Clear placeholder when clicked
    If Not Intersect(Target, inputCell) Is Nothing Then
        If inputCell.Value = "Search" Then
            inputCell.Value = ""
            inputCell.Font.Italic = False
            inputCell.Font.Color = vbBlack
        End If
    Else
        ' Restore placeholder if cell left empty
        If inputCell.Value = "" Then
            inputCell.Value = "Search"
            inputCell.Font.Italic = True
            inputCell.Font.Color = RGB(150, 150, 150)
        End If
    End If
End Sub

