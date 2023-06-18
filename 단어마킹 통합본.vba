Private Sub CommandButton1_Click()

Dim ar(14)
Dim i


ar(0) = TextBox1
ar(1) = TextBox2
ar(2) = TextBox3
ar(3) = TextBox4
ar(4) = TextBox5
ar(5) = TextBox6
ar(6) = TextBox7
ar(7) = TextBox8
ar(8) = TextBox9
ar(9) = TextBox10
ar(10) = TextBox11
ar(11) = TextBox12
ar(12) = TextBox13
ar(13) = TextBox14
ar(14) = TextBox15

Selection.Replace " ", "┃"

For i = 0 To 14

    If ar(i) <> "" Then
    Selection.Replace ar(i), "┃"
    End If
    
Next

If CheckBox1.Value = True Then
Selection.Replace "~*", "┃"
End If

Selection.Replace "┃┃┃┃", "┃"
Selection.Replace "┃┃┃", "┃"
Selection.Replace "┃┃", "┃"

MsgBox "끝!"

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub




Private Sub CommandButton1_Click()
        
    Call Highlighting_Word
        
End Sub

Private Sub CommandButton2_Click()

    Unload UserForm1

End Sub

Private Sub Highlighting_Word()


Dim intRow As Long
Dim strTemp As String
Dim strWord() As String
Dim intNum As Long
Dim intWhat As Integer, intFrom As Integer
Dim i As Long, j As Long, r As Long
Dim rngCell As Range, rngTarget As Range
Dim intAddr As Integer
Dim intLen As Integer

Dim intAddr2 As Integer
Dim intLctn2 As Integer

Dim intClrNo As Integer                 '// ColorIndex Number

Application.ScreenUpdating = False


    If OptionButton1 Then
        intClrNo = 3
    ElseIf OptionButton2 Then
        intClrNo = 5
    ElseIf OptionButton3 Then
        intClrNo = 4
    Else
        intClrNo = 15
    End If


Unload UserForm1

intRow = Range(Cells(3, 1), Cells(1048576, 1).End(xlUp)).Count + 2

showIE = intRow - 2

For i = 3 To intRow
    
    If Cells(i - 1, 1) = Cells(i, 1) Then GoTo NT
    
    Set rngCell = Cells(i, 1)
    
    intNum = 0
    intWhat = 1
    On Error Resume Next

    Do While Err = 0
    
        intFrom = Application.WorksheetFunction.Find("┃", rngCell, intWhat)
        intNum = intNum + 1
        ReDim Preserve strWord(intNum)
        strWord(intNum) = Mid(rngCell, intWhat, IIf(Err, Len(rngCell) + 1, intFrom) - intWhat)
        intWhat = intFrom + 1

    Loop

NT:
    
    For r = 1 To intNum
        
        Set rngTarget = Cells(i, 2)
    
        intAddr = InStr(1, rngTarget, strWord(r), 1)
        intLen = Len(strWord(r))
        
        If intAddr = 0 Then GoTo EX
        If intLen = 0 Then GoTo EX
        
        With rngTarget.Characters(Start:=intAddr, Length:=intLen).Font
            .FontStyle = "굵게"
            .ColorIndex = intClrNo
        End With
        
        
'//////////////////////////////////////// 2번째 중복단어 검색
        intLctn2 = intAddr + intLen
        intAddr2 = InStr(intLctn2, rngTarget, strWord(r), 1)
        
        If intAddr2 = 0 Then GoTo EX
        
        With rngTarget.Characters(Start:=intAddr2, Length:=intLen).Font
            .FontStyle = "굵게"
            .ColorIndex = intClrNo
        End With
'////////////////////////////////////////////////

''//////////////////////////////////////// 3번째 중복단어 검색
'        intLctn3 = intAddr2 + intLen
'        intAddr3 = InStr(intLctn3, rngTarget, strWord(r), 1)
'
'        If intAddr3 = 0 Then GoTo EX
'
'        With rngTarget.Characters(Start:=intAddr3, Length:=intLen).Font
'            .FontStyle = "굵게"
'            .ColorIndex = intClrNo
'        End With
''////////////////////////////////////////////////
'
''//////////////////////////////////////// 4번째 중복단어 검색
'        intLctn4 = intAddr3 + intLen
'        intAddr4 = InStr(intLctn4, rngTarget, strWord(r), 1)
'
'        If intAddr4 = 0 Then GoTo EX
'
'        With rngTarget.Characters(Start:=intAddr4, Length:=intLen).Font
'            .FontStyle = "굵게"
'            .ColorIndex = intClrNo
'        End With
''////////////////////////////////////////////////
EX:
        If intAddr = 0 Then Cells(i, 3) = Cells(i, 3).Value & " " & strWord(r)
        
    Next r


Next i

'Unload UserForm1

Application.ScreenUpdating = True

MsgBox "작업완료"

End Sub


Private Sub CommandButton3_Click()

seosic.Show

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub




Private Sub CommandButton1_Click()
        
    Call Highlighting_Word
        
End Sub

Private Sub CommandButton2_Click()

    Unload UserForm1

End Sub

Private Sub Highlighting_Word()


Dim intRow As Long
Dim strTemp As String
Dim strWord() As String
Dim intNum As Long
Dim intWhat As Integer, intFrom As Integer
Dim i As Long, j As Long, r As Long
Dim rngCell As Range, rngTarget As Range
Dim intAddr As Integer
Dim intLen As Integer

Dim intAddr2 As Integer
Dim intLctn2 As Integer

Dim intClrNo As Integer                 '// ColorIndex Number

Application.ScreenUpdating = False


    If OptionButton1 Then
        intClrNo = 3
    ElseIf OptionButton2 Then
        intClrNo = 5
    ElseIf OptionButton3 Then
        intClrNo = 4
    Else
        intClrNo = 15
    End If


Unload UserForm1

intRow = Range(Cells(3, 1), Cells(1048576, 1).End(xlUp)).Count + 2

showIE = intRow - 2

For i = 3 To intRow
    
    If Cells(i - 1, 1) = Cells(i, 1) Then GoTo NT
    
    Set rngCell = Cells(i, 1)
    
    intNum = 0
    intWhat = 1
    On Error Resume Next

    Do While Err = 0
    
        intFrom = Application.WorksheetFunction.Find("┃", rngCell, intWhat)
        intNum = intNum + 1
        ReDim Preserve strWord(intNum)
        strWord(intNum) = Mid(rngCell, intWhat, IIf(Err, Len(rngCell) + 1, intFrom) - intWhat)
        intWhat = intFrom + 1

    Loop

NT:
    
    For r = 1 To intNum
        
        Set rngTarget = Cells(i, 2)
    
        intAddr = InStr(1, rngTarget, strWord(r), 1)
        intLen = Len(strWord(r))
        
        If intAddr = 0 Then GoTo EX
        If intLen = 0 Then GoTo EX
        
        With rngTarget.Characters(Start:=intAddr, Length:=intLen).Font
            .FontStyle = "굵게"
            .ColorIndex = intClrNo
        End With
        
        
'//////////////////////////////////////// 2번째 중복단어 검색
        intLctn2 = intAddr + intLen
        intAddr2 = InStr(intLctn2, rngTarget, strWord(r), 1)
        
        If intAddr2 = 0 Then GoTo EX
        
        With rngTarget.Characters(Start:=intAddr2, Length:=intLen).Font
            .FontStyle = "굵게"
            .ColorIndex = intClrNo
        End With
'////////////////////////////////////////////////

''//////////////////////////////////////// 3번째 중복단어 검색
'        intLctn3 = intAddr2 + intLen
'        intAddr3 = InStr(intLctn3, rngTarget, strWord(r), 1)
'
'        If intAddr3 = 0 Then GoTo EX
'
'        With rngTarget.Characters(Start:=intAddr3, Length:=intLen).Font
'            .FontStyle = "굵게"
'            .ColorIndex = intClrNo
'        End With
''////////////////////////////////////////////////
'
''//////////////////////////////////////// 4번째 중복단어 검색
'        intLctn4 = intAddr3 + intLen
'        intAddr4 = InStr(intLctn4, rngTarget, strWord(r), 1)
'
'        If intAddr4 = 0 Then GoTo EX
'
'        With rngTarget.Characters(Start:=intAddr4, Length:=intLen).Font
'            .FontStyle = "굵게"
'            .ColorIndex = intClrNo
'        End With
''////////////////////////////////////////////////
EX:
        If intAddr = 0 Then Cells(i, 3) = Cells(i, 3).Value & " " & strWord(r)
        
    Next r


Next i

'Unload UserForm1

Application.ScreenUpdating = True

MsgBox "작업완료"

End Sub


Private Sub CommandButton3_Click()

seosic.Show

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub



Private Sub CommandButton1_Click()
        
    Call Highlighting_Word
        
End Sub

Private Sub CommandButton2_Click()

    Unload UserForm1

End Sub

Private Sub Highlighting_Word()


Dim intRow As Long
Dim strTemp As String
Dim strWord() As String
Dim intNum As Long
Dim intWhat As Integer, intFrom As Integer
Dim i As Long, j As Long, r As Long
Dim rngCell As Range, rngTarget As Range
Dim intAddr As Integer
Dim intLen As Integer

Dim intAddr2 As Integer
Dim intLctn2 As Integer

Dim intClrNo As Integer                 '// ColorIndex Number

Application.ScreenUpdating = False


    If OptionButton1 Then
        intClrNo = 3
    ElseIf OptionButton2 Then
        intClrNo = 5
    ElseIf OptionButton3 Then
        intClrNo = 4
    Else
        intClrNo = 15
    End If


Unload UserForm1

intRow = Range(Cells(3, 1), Cells(1048576, 1).End(xlUp)).Count + 2

showIE = intRow - 2

For i = 3 To intRow
    
    If Cells(i - 1, 1) = Cells(i, 1) Then GoTo NT
    
    Set rngCell = Cells(i, 1)
    
    intNum = 0
    intWhat = 1
    On Error Resume Next

    Do While Err = 0
    
        intFrom = Application.WorksheetFunction.Find("┃", rngCell, intWhat)
        intNum = intNum + 1
        ReDim Preserve strWord(intNum)
        strWord(intNum) = Mid(rngCell, intWhat, IIf(Err, Len(rngCell) + 1, intFrom) - intWhat)
        intWhat = intFrom + 1

    Loop

NT:
    
    For r = 1 To intNum
        
        Set rngTarget = Cells(i, 2)
    
        intAddr = InStr(1, rngTarget, strWord(r), 1)
        intLen = Len(strWord(r))
        
        If intAddr = 0 Then GoTo EX
        If intLen = 0 Then GoTo EX
        
        With rngTarget.Characters(Start:=intAddr, Length:=intLen).Font
            .FontStyle = "굵게"
            .ColorIndex = intClrNo
        End With
        
        
'//////////////////////////////////////// 2번째 중복단어 검색
        intLctn2 = intAddr + intLen
        intAddr2 = InStr(intLctn2, rngTarget, strWord(r), 1)
        
        If intAddr2 = 0 Then GoTo EX
        
        With rngTarget.Characters(Start:=intAddr2, Length:=intLen).Font
            .FontStyle = "굵게"
            .ColorIndex = intClrNo
        End With
'////////////////////////////////////////////////

''//////////////////////////////////////// 3번째 중복단어 검색
'        intLctn3 = intAddr2 + intLen
'        intAddr3 = InStr(intLctn3, rngTarget, strWord(r), 1)
'
'        If intAddr3 = 0 Then GoTo EX
'
'        With rngTarget.Characters(Start:=intAddr3, Length:=intLen).Font
'            .FontStyle = "굵게"
'            .ColorIndex = intClrNo
'        End With
''////////////////////////////////////////////////
'
''//////////////////////////////////////// 4번째 중복단어 검색
'        intLctn4 = intAddr3 + intLen
'        intAddr4 = InStr(intLctn4, rngTarget, strWord(r), 1)
'
'        If intAddr4 = 0 Then GoTo EX
'
'        With rngTarget.Characters(Start:=intAddr4, Length:=intLen).Font
'            .FontStyle = "굵게"
'            .ColorIndex = intClrNo
'        End With
''////////////////////////////////////////////////
EX:
        If intAddr = 0 Then Cells(i, 3) = Cells(i, 3).Value & " " & strWord(r)
        
    Next r


Next i

'Unload UserForm1

Application.ScreenUpdating = True

MsgBox "작업완료"

End Sub


Private Sub CommandButton3_Click()

seosic.Show

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub



Sub Highlighting()

    With UserForm1
        .Show
    End With

End Sub


Function ConcatText(ByVal 범위 As Variant) As String
 Dim T, Rev() As Variant
 Dim i As Integer
  For Each T In 범위
    If Len(T) Then
     ReDim Preserve Rev(i)
       Rev(i) = T
        i = i + 1
     End If
   Next T
 ConcatText = Join(Rev, "┃")
End Function



Sub 단추2_Click()

Selection.Replace " ", ""
Selection.Replace "~┃┃┃┃~", "┃"
Selection.Replace "~┃┃┃~", "┃"
Selection.Replace "~┃┃~", "┃"

Selection.Replace "~,~", "┃"
Selection.Replace "~/~", "┃"
Selection.Replace "~(~", "┃"
Selection.Replace "~)~", "┃"
Selection.Replace "~[~", "┃"
Selection.Replace "~]~", "┃"
Selection.Replace "~_~", "┃"
Selection.Replace "~:~", "┃"


End Sub
