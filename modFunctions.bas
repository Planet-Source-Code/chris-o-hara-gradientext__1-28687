Attribute VB_Name = "modFunctions"
Option Explicit

Public Function CreateGradientFX(strHTMLText As String, strPath As String, strTextSize As String, strFont As String)

    Dim strSeparatorpos1            As String
    Dim strSeparatorpos2            As String
    Dim strColorRGB(0 To 10)        As String
    Dim strColorr(0 To 10)          As String
    Dim strColorg(0 To 10)          As String
    Dim strColorb(0 To 10)          As String
    Dim strTemp                     As String
    Dim strCreateTag                As String
    Dim strBackColour               As String
    Dim dblStringLength             As Double
    Dim dblValueBetweenR(0 To 10)   As Double
    Dim dblValueBetweenG(0 To 10)   As Double
    Dim dblValueBetweenB(0 To 10)   As Double
    Dim dblColours                  As Double
    Dim dblIntervalr(0 To 10)       As Double
    Dim dblIntervalg(0 To 10)       As Double
    Dim dblIntervalb(0 To 10)       As Double
    Dim strHTMLtextSectionLength    As Double
    Dim varColorRValue              As Variant
    Dim varColorGValue              As Variant
    Dim varColorBValue              As Variant
    Dim x                           As Double
    Dim y                           As Double

    dblStringLength = Len(strHTMLText)
    dblColours = frmMain.pboxColour.UBound
    strHTMLtextSectionLength = dblStringLength / dblColours
    strCreateTag = "<!--/This gradient text was created using Gradient Text by Chris O'Hara/-->"
    
    frmMain.pbProgress.Min = 0
    frmMain.pbProgress.Max = (strHTMLtextSectionLength + (dblColours * Int(dblStringLength / (dblColours))))
    frmMain.pbProgress.Value = 0
    
    strColorRGB(0) = GetRGBVals(frmMain.pboxBColour)
        
    strSeparatorpos1 = InStr(1, strColorRGB(0), ",")
    strSeparatorpos2 = InStr(strSeparatorpos1 + 1, strColorRGB(0), ",")
    strColorr(0) = Mid(strColorRGB(0), 1, strSeparatorpos1 - 1)
    strColorg(0) = Mid(strColorRGB(0), strSeparatorpos1 + 1, (strSeparatorpos2 - 1) - strSeparatorpos1)
    strColorb(0) = Mid(strColorRGB(0), strSeparatorpos2 + 1, Len(strColorRGB(0)) - strSeparatorpos2)
    
    varColorRValue = strColorr(0)
    varColorGValue = strColorg(0)
    varColorBValue = strColorb(0)
    
    If Val(varColorRValue) = varColorRValue Then varColorRValue = Hex(varColorRValue)
    If Val(varColorGValue) = varColorGValue Then varColorGValue = Hex(varColorGValue)
    If Val(varColorBValue) = varColorBValue Then varColorBValue = Hex(varColorBValue)
                    
    If Len(varColorRValue) = 1 Then varColorRValue = CStr("0" & varColorRValue)
    If Len(varColorGValue) = 1 Then varColorGValue = CStr("0" & varColorGValue)
    If Len(varColorBValue) = 1 Then varColorBValue = CStr("0" & varColorBValue)
    
    strBackColour = "#" & varColorRValue & varColorGValue & varColorBValue
    
    For x = 0 To dblColours Step 1
    
        strColorRGB(x) = GetRGBVals(frmMain.pboxColour(x))
        
        strSeparatorpos1 = InStr(1, strColorRGB(x), ",")
        strSeparatorpos2 = InStr(strSeparatorpos1 + 1, strColorRGB(x), ",")
        strColorr(x) = Mid(strColorRGB(x), 1, strSeparatorpos1 - 1)
        strColorg(x) = Mid(strColorRGB(x), strSeparatorpos1 + 1, (strSeparatorpos2 - 1) - strSeparatorpos1)
        strColorb(x) = Mid(strColorRGB(x), strSeparatorpos2 + 1, Len(strColorRGB(x)) - strSeparatorpos2)
        
    Next x
    
    For x = 0 To dblColours Step 1
        
        dblValueBetweenR(x) = Abs(Val(strColorr(x)) - Val(strColorr(x + 1)))
        dblValueBetweenG(x) = Abs(Val(strColorg(x)) - Val(strColorg(x + 1)))
        dblValueBetweenB(x) = Abs(Val(strColorb(x)) - Val(strColorb(x + 1)))
        
        dblIntervalr(x) = dblValueBetweenR(x) / Int(dblStringLength / (dblColours))
        dblIntervalg(x) = dblValueBetweenG(x) / Int(dblStringLength / (dblColours))
        dblIntervalb(x) = dblValueBetweenB(x) / Int(dblStringLength / (dblColours))
        
    Next x
    
    Dim tempr, tempg, tempb
    
        Open strPath For Output As #1
    
        Print #1, "<html>" & vbCrLf & "<body bgcolor=" & strBackColour & ">" & vbCrLf & "" & vbCrLf & "<table width=100% height=100%>" & vbCrLf & "" & vbCrLf & "<tr>" & vbCrLf & "<td>" & vbCrLf & "<center>" & vbCrLf & "" & vbCrLf & "<font face='" & strFont & "' size=" & strTextSize & ">" & vbCrLf & "" & vbCrLf & "" & vbCrLf
    
    For y = 0 To dblColours Step 1
            For x = 1 To strHTMLtextSectionLength Step 1
            
                If Mid(strHTMLText, (x + (y * Int(dblStringLength / (dblColours)))), 1) = "" Then Exit For
            
                If strColorr(y) >= strColorr(y + 1) Then varColorRValue = Int(strColorr(y) - (dblIntervalr(y) * x))
                If strColorr(y + 1) >= strColorr(y) Then varColorRValue = Int(strColorr(y) + (dblIntervalr(y) * x))
                
                If strColorg(y) >= strColorg(y + 1) Then varColorGValue = Int(strColorg(y) - (dblIntervalg(y) * x))
                If strColorg(y + 1) >= strColorg(y) Then varColorGValue = Int(strColorg(y) + (dblIntervalg(y) * x))
                
                If strColorb(y) >= strColorb(y + 1) Then varColorBValue = Int(strColorb(y) - (dblIntervalb(y) * x))
                If strColorb(y + 1) >= strColorb(y) Then varColorBValue = Int(strColorb(y) + (dblIntervalb(y) * x))
            
                If varColorRValue < 0 Then varColorRValue = 0
                If varColorGValue < 0 Then varColorGValue = 0
                If varColorBValue < 0 Then varColorBValue = 0
                
                If varColorRValue > 255 Then varColorRValue = 255
                If varColorGValue > 255 Then varColorGValue = 255
                If varColorBValue > 255 Then varColorBValue = 255
                
                            
                If Val(varColorRValue) = varColorRValue Then varColorRValue = Hex(varColorRValue)
                If Val(varColorGValue) = varColorGValue Then varColorGValue = Hex(varColorGValue)
                If Val(varColorBValue) = varColorBValue Then varColorBValue = Hex(varColorBValue)
            
                If Len(varColorRValue) = 1 Then varColorRValue = CStr("0" & varColorRValue)
                If Len(varColorGValue) = 1 Then varColorGValue = CStr("0" & varColorGValue)
                If Len(varColorBValue) = 1 Then varColorBValue = CStr("0" & varColorBValue)
                
                If Mid(strHTMLText, (x + (y * Int(dblStringLength / (dblColours)))), 1) = " " Then
                    Print #1, strTemp
                    strTemp = ""
                ElseIf Mid(strHTMLText, (x + (y * Int(dblStringLength / (dblColours)))), 1) = Chr(13) Then
                    Print #1, strTemp
                    Print #1, "<br>"
                    strTemp = ""
                Else
                    strTemp = strTemp & "<font color=#" & varColorRValue & varColorGValue & varColorBValue & ">" & Mid(strHTMLText, (x + (y * Int(dblStringLength / (dblColours)))), 1) & "</font>"
                End If
                
                frmMain.pbProgress.Value = (x + (y * Int(dblStringLength / (dblColours))))
                
            Next x
    Next y

    frmMain.pbProgress.Value = frmMain.pbProgress.Max

    Print #1, strTemp & vbCrLf
    strTemp = ""

    Print #1, vbCrLf & vbCrLf & "</font>" & vbCrLf & "" & vbCrLf & "</center>" & vbCrLf & "</td>" & vbCrLf & "</tr>" & vbCrLf & "" & vbCrLf & "</table>" & vbCrLf & "" & vbCrLf & "</body>" & vbCrLf & "</html>"

    Close #1
    
    strTemp = ""

End Function

Public Function GetRGBVals(picbox As PictureBox) As String

'Declare Variables
    Dim lngCol  As Long
    Dim lngR    As Long
    Dim lngG    As Long
    Dim lngB    As Long
    Dim lngX    As Long
    
    'Set values
    lngR = 0
    lngG = 0
    lngB = 0
    
    'Get RGB Values
    lngCol = picbox.BackColor
    
    If lngCol = 0 And picbox.BackColor = -2147483643 Then
        lngCol = 255
        lngR = 255
        lngG = 255
        lngB = 255
    End If
    
    For lngX = 1 To 513 Step 1
        
        If lngCol >= 65536 Then
                lngCol = lngCol - 65536
                lngB = lngB + 1
            ElseIf lngCol >= 256 Then
                lngCol = lngCol - 256
                lngG = lngG + 1
            Else
                lngR = lngCol
        End If
        
    Next lngX
    
    GetRGBVals = lngR & "," & lngG & "," & lngB

End Function

Public Function AddPicBox(colour As ColorConstants)

    'Add a picturebox to the exact right of the previous one
    Load frmMain.pboxColour(frmMain.pboxColour.UBound + 1)
    frmMain.pboxColour(frmMain.pboxColour.UBound).BackColor = colour
    frmMain.pboxColour(frmMain.pboxColour.UBound).Left = frmMain.pboxColour(frmMain.pboxColour.UBound - 1).Left + frmMain.pboxColour(frmMain.pboxColour.UBound - 1).Width
    frmMain.pboxColour(frmMain.pboxColour.UBound).Top = frmMain.pboxColour(frmMain.pboxColour.UBound - 1).Top
    frmMain.pboxColour(frmMain.pboxColour.UBound).Visible = True

End Function
