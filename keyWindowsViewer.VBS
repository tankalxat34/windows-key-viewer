'                       Script windows-key-viewer
'   This VBS script can help you know your Windows key, version and product ID
'
'                  tankalxat34 (Alexander Podstrechnyy)
'              https://github.com/tankalxat34/windows-key-viewer
'
' License:
' MIT License
'
' Copyright (c) 2022 Alexander Podstrechnyy
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Set WshShell = CreateObject("WScript.Shell")
regKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
DigitalProductId = WshShell.RegRead(regKey & "DigitalProductId")
Win8ProductName = "Windows Product Name: " & WshShell.RegRead(regKey & "ProductName") & vbNewLine
Win8ProductID = "Windows Product ID: " & WshShell.RegRead(regKey & "ProductID") & vbNewLine
Win8ProductKey = ConvertToKey(DigitalProductId)
strProductKey ="Windows Key: " & Win8ProductKey
Win8ProductID = Win8ProductName & Win8ProductID & strProductKey
' MsgBox Win8ProductKey, vbInformation, "Windows Key Viewer"
MsgBox Win8ProductID, vbInformation, "Windows Key Viewer"
Function ConvertToKey(regKey)
    Const KeyOffset = 52
    isWin8 = (regKey(66) \ 6) And 1
    regKey(66) = (regKey(66) And &HF7) Or ((isWin8 And 2) * 4)
    j = 24
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    Do
        Cur = 0
        y = 14
        Do
            Cur = Cur * 256
            Cur = regKey(y + KeyOffset) + Cur
            regKey(y + KeyOffset) = (Cur \ 24)
            Cur = Cur Mod 24
            y = y -1
        Loop While y >= 0
        j = j -1
        winKeyOutput = Mid(Chars, Cur + 1, 1) & winKeyOutput
        Last = Cur
    Loop While j >= 0
    If (isWin8 = 1) Then
        keypart1 = Mid(winKeyOutput, 2, Last)
        insert = "N"
        winKeyOutput = Replace(winKeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
        If Last = 0 Then winKeyOutput = insert & winKeyOutput
    End If
    a = Mid(winKeyOutput, 1, 5)
    b = Mid(winKeyOutput, 6, 5)
    c = Mid(winKeyOutput, 11, 5)
    d = Mid(winKeyOutput, 16, 5)
    e = Mid(winKeyOutput, 21, 5)
    ConvertToKey = a & "-" & b & "-" & c & "-" & d & "-" & e
End Function