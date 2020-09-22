Attribute VB_Name = "Module3"
Public Write_Player As Integer
Public Rank_Value As Integer
Public WriteName As Integer
Public WriteScore As Integer
Public WriteRank As Integer

Public Function WriteNameModule()
Write_Player = Rank_Value
Select Case Write_Player
Case 1
Form35.Label1.Caption = Form35.Label1.Caption & left(Form35.Text1, 1)
Case 2
Form35.Label3.Caption = Form35.Label3.Caption & left(Form35.Text1, 1)
Case 3
Form35.Label5.Caption = Form35.Label5.Caption & left(Form35.Text1, 1)
Case 4
Form35.Label7.Caption = Form35.Label7.Caption & left(Form35.Text1, 1)
Case 5
Form35.Label9.Caption = Form35.Label9.Caption & left(Form35.Text1, 1)
Case 6
Form35.Label11.Caption = Form35.Label11.Caption & left(Form35.Text1, 1)
Case 7
Form35.Label13.Caption = Form35.Label13.Caption & left(Form35.Text1, 1)
Case 8
Form35.Label15.Caption = Form35.Label15.Caption & left(Form35.Text1, 1)
Case 9
Form35.Label17.Caption = Form35.Label17.Caption & left(Form35.Text1, 1)
Case 0
Form35.Label19.Caption = Form35.Label19.Caption & left(Form35.Text1, 1)
End Select
End Function
Public Function WriteScoreModule()
Write_Player = Rank_Value
Select Case Write_Player
Case 1
Form35.Label2.Caption = Form35.Label2.Caption & left(Form35.Text1, 1)
Case 2
Form35.Label4.Caption = Form35.Label4.Caption & left(Form35.Text1, 1)
Case 3
Form35.Label6.Caption = Form35.Label6.Caption & left(Form35.Text1, 1)
Case 4
Form35.Label8.Caption = Form35.Label8.Caption & left(Form35.Text1, 1)
Case 5
Form35.Label10.Caption = Form35.Label10.Caption & left(Form35.Text1, 1)
Case 6
Form35.Label12.Caption = Form35.Label12.Caption & left(Form35.Text1, 1)
Case 7
Form35.Label14.Caption = Form35.Label14.Caption & left(Form35.Text1, 1)
Case 8
Form35.Label16.Caption = Form35.Label16.Caption & left(Form35.Text1, 1)
Case 9
Form35.Label18.Caption = Form35.Label18.Caption & left(Form35.Text1, 1)
Case 0
Form35.Label20.Caption = Form35.Label20.Caption & left(Form35.Text1, 1)
End Select
End Function

Public Function WhatsMyRank(scorevalue As Long)
If Val(scorevalue) < Val(Form35.Label20.Caption) Then
WhatsMyRank = 88
Exit Function
End If


If Val(scorevalue) >= Val(Form35.Label2.Caption) Then
WhatsMyRank = 1
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label4.Caption) Then
WhatsMyRank = 2
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label6.Caption) Then
WhatsMyRank = 3
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label8.Caption) Then
WhatsMyRank = 4
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label10.Caption) Then
WhatsMyRank = 5
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label12.Caption) Then
WhatsMyRank = 6
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label14.Caption) Then
WhatsMyRank = 7
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label16.Caption) Then
WhatsMyRank = 8
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label18.Caption) Then
WhatsMyRank = 9
Exit Function
End If
If Val(scorevalue) >= Val(Form35.Label20.Caption) Then
WhatsMyRank = 0
Exit Function
End If
End Function
