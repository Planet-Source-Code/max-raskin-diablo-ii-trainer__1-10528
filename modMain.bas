Attribute VB_Name = "modMain"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Enum D2Status
    Normal = 0
    Sir = &H708
    Lord = &H908
    Baron = &HC08
End Enum

Enum D2Attr
    Strength = 566
    Dexterity = 574
    Vitality = 578
    Energy = 570
    Life = 583
    Mana = 591
    Stamina = 599
    CurLife = 587
    CurMana = 595
    CurStamina = 603
End Enum

'Here is the whole magic trick =) :
Function SetStatus(SaveFile As String, Status As D2Status)
    Open SaveFile For Binary As #1 'Open a save file a binary
    Put #1, 25, Status  'Set the value
    Close #1  'close the file
End Function


Public Function EnumFilesByExt(Path As String, ListBox As ListBox, Extension As String)
ListBox.Clear
    Dim XDir() As String
    Dim TmpDir As String
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If


    DoEvents
        TmpDir = Dir(Path, vbDirectory + vbHidden + vbSystem + vbArchive + vbReadOnly)


        Do While TmpDir <> ""


            If TmpDir <> "." And TmpDir <> ".." Then


                If (GetAttr(Path & TmpDir)) <> vbDirectory Then
                    If Right(TmpDir, Len(Extension)) = Extension Then ListBox.AddItem TmpDir
                    ReDim Preserve XDir(DirCount) As String
                End If
            End If
            TmpDir = Dir
            
        Loop
End Function

Function GetLevel(SaveFile As String) As String
    Dim vRetVal, nLVL As Integer, lPos As Long
    lPos = 37 'The position where the value stands
    Open SaveFile For Binary As #1  'open a save file as binary
    Get #1, lPos, nLVL 'Now, get the value
    Close #1   'Close the file
    vRetVal = Hex(nLVL)
    vRetVal = "&H" & CStr(vRetVal) 'convert it to a vb hex value because the clng function does not know the diffrence between a number and a hex without the &H
    If vRetVal = 0 Then
        GetLevel = "1" 'Get the level, and we're all done ! :-)
    Else
        GetLevel = CStr(CLng(vRetVal)) 'Get the level, and we're all done ! :-)
    End If
End Function

Function GetStatus(SaveFile As String) As String
    Dim nStatus As Integer, str As String
    Open SaveFile For Binary As #1  'open a save file a binary
    Get #1, 26, nStatus 'now, get the value
    Close #1   'close the file
    str = GetClass(SaveFile)
    If str = "Barbarian" Then GoTo SetMan
    If str = "Necromancer" Then GoTo SetMan
    If str = "Paladin" Then GoTo SetMan
    If str = "Amazon" Then GoTo SetWomen
    If str = "Sorceress" Then GoTo SetWomen
SetMan:         If Hex(nStatus) = 7 Then GetStatus = "Sir"
                If Hex(nStatus) = 5 Then GetStatus = "Sir"
                If Hex(nStatus) = 9 Then GetStatus = "Lord"
                If CStr(Hex(nStatus)) = "C" Then GetStatus = "Baron"
                Exit Function
SetWomen:   If Hex(nStatus) = 7 Then GetStatus = "Dame"
            If Hex(nStatus) = 5 Then GetStatus = "Dame"
            If Hex(nStatus) = 9 Then GetStatus = "Lady"
            If CStr(Hex(nStatus)) = "C" Then GetStatus = "Baroness"
    If Hex(nStatus) = 0 Then GetStatus = "" 'None (Not killed Diablo yet)
End Function

'Gets the character class out of a save file
Function GetClass(SaveFile As String) As String
    Dim vRetVal As Integer, nClass As Integer
    Open SaveFile For Binary As #1  'open a save file as binary
    Get #1, 35, nClass 'now, get the value
    Close #1   'close the file
    Select Case nClass 'Returned cases:
    Case 0
        GetClass = "Amazon"
    Case 1
        GetClass = "Sorceress"
    Case 2
        GetClass = "Necromancer"
    Case 3
        GetClass = "Paladin"
    Case 4
        GetClass = "Barbarian"
    End Select
End Function

'Converts a text box to numeric values can be typed only
Function Text2Numeric(Textbox As Textbox)
    Dim wndLong As Long
    wndLong = GetWindowLong(Textbox.hwnd, (-16))
    SetWindowLong Textbox.hwnd, (-16), wndLong Or &H2000&
End Function

'Get a specific character attribute
Function GetAttrib(SaveFile As String, Attrib As D2Attr) As String
    Dim vRetVal, nAttr As Integer, lPos As Integer
    lPos = Attrib   'The position where the value stands
    Open SaveFile For Binary As #1  'open a save file as binary
    Get #1, lPos, nAttr 'now, get the value
    Close #1   'close the file
    GetAttrib = nAttr 'set the value
End Function

'Set a specific character attribute
Function SetAttrib(SaveFile As String, Attr As D2Attr, NewVal As Integer)
    Open SaveFile For Binary As #1 'Open a save file a binary
    NewVal = "&H" & Hex(NewVal)
    Put #1, Attr, NewVal 'Set the value
    Close #1  'close the file
End Function
