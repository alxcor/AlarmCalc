VERSION 5.00
Begin VB.Form frmAlCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Calculator"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDB2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Bit number in DB2 (e.g. 78.2 or DB2 DBX78.2)"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtAlarm 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Alarm Nr. (e.g. 700001)"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "i"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "DB2 bit"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Alarm"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblAlType 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Alarm description (see 840D manual for details)"
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frmAlCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblInfo_Click()
Dim txtInfo As String
    txtInfo = "Alarm Calculator v0.1 / 03.2010"
    txtInfo = txtInfo & vbCrLf & "alxcor@yahoo.com"
    MsgBox txtInfo, vbInformation, "About AlarmCalc..."
End Sub

Private Sub txtAlarm_KeyPress(KeyAscii As Integer)
On Error GoTo ErrAlarm
    If KeyAscii <> 13 Then
        Exit Sub
    End If
'    Dim iAlarma As Long
    Dim iTemp, iByte, iBit, iChan, iAxis As Integer
    Dim sTemp, sAlarm6, sAlarm5, sAlarm4, sAlarm3, sAlarm2, sAlarm1  As String
    
    frmAlCalc.Caption = "Alarm Calculator"
    txtDB2.Text = ""
    iByte = 999
    iBit = 999
    
    If (Len(txtAlarm.Text) > 6) Then
        txtDB2.Text = "--//--"
        lblAlType.Caption = ">6"
        Exit Sub
    End If
    
    iTemp = Int(Val(txtAlarm.Text))
    txtAlarm.Text = Trim$(Str(iTemp))
    sTemp = txtAlarm.Text
    
    sAlarm1 = "0"
    sAlarm2 = "0"
    sAlarm3 = "0"
    sAlarm4 = "0"
    sAlarm5 = "0"
    sAlarm6 = "0"

    If Len(sTemp) > 0 Then
        sAlarm1 = Right(sTemp, 1)
    End If
    If Len(sTemp) > 1 Then
        sTemp = Left(sTemp, Len(sTemp) - 1)
        sAlarm2 = Right(sTemp, 1)
    End If
    If Len(sTemp) > 1 Then
        sTemp = Left(sTemp, Len(sTemp) - 1)
        sAlarm3 = Right(sTemp, 1)
    End If
    If Len(sTemp) > 1 Then
        sTemp = Left(sTemp, Len(sTemp) - 1)
        sAlarm4 = Right(sTemp, 1)
    End If
    If Len(sTemp) > 1 Then
        sTemp = Left(sTemp, Len(sTemp) - 1)
        sAlarm5 = Right(sTemp, 1)
    End If
    If Len(sTemp) > 1 Then
        sTemp = Left(sTemp, Len(sTemp) - 1)
        sAlarm6 = Right(sTemp, 1)
    End If
    
'    MsgBox sAlarm6 & "." & sAlarm5 & "." & sAlarm4 & "." & sAlarm3 & "." & sAlarm2 & "." & sAlarm1

    Select Case sAlarm6
        Case "0"
            Select Case sAlarm5
                Case "0"
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "NCK Alarm; General> Diagnostics Manual"
                    Exit Sub
                Case "1"
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "NCK Alarm; Channel> Diagnostics Manual"
                    Exit Sub
                Case "2"
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "NCK Alarm; Axis/Sp.> Diagnostics Manual"
                    Exit Sub
                Case "3"
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "NCK Alarm; Functional> Diagnostics Manual"
                    Exit Sub
                Case "6"
                    txtDB2.Text = "--//--"
                    iTemp = Int(Val(sAlarm4 & sAlarm3 & sAlarm2 & sAlarm1))
                    If (iTemp >= 0) And (iTemp < 5000) Then
                        lblAlType.Caption = "NCK Alarm; Cycle SIE.> Diagnostics Manual"
                    Else
                        lblAlType.Caption = "NCK Alarm; Cycle USER> Manuf. Cycles!"
                    End If
                    Exit Sub
                Case "7"
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "NCK Alarm; Compile Cycles, manuf. and OEM"
                    Exit Sub
                Case Else
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "N.A."
                    Exit Sub
            End Select
        Case "5"
            iChan = Int(Val(sAlarm5))
            If iChan > 0 And iChan < 9 Then
                iByte = 18 * (iChan - 1)
                lblAlType.Caption = "Chan. " & sAlarm5 & ": "
                sTemp = sAlarm4 & sAlarm3 & sAlarm2 & sAlarm1
                iTemp = Int(Val(sTemp))
                Select Case iTemp
                    Case 0 To 15
                        iByte = iByte + Int(iTemp / 8)
                        iBit = iTemp Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Disable"
                    Case 100 To 107
                        iByte = iByte + 2
                        iBit = (iTemp - 100) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed and Read In Disable"
                    Case 108 To 115
                        iByte = iByte + 3
                        iBit = (iTemp - 108) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed and Read In Disable"
                    Case 116 To 123
                        iByte = iByte + 4
                        iBit = (iTemp - 116) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed and Read In Disable"
                    Case 124 To 131
                        iByte = iByte + 5
                        iBit = (iTemp - 124) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed and Read In Disable"
                    Case 200 To 207
                        iByte = iByte + 6
                        iBit = (iTemp - 200) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Read In Disable"
                    Case 208 To 215
                        iByte = iByte + 7
                        iBit = (iTemp - 208) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Read In Disable"
                    Case 216 To 223
                        iByte = iByte + 8
                        iBit = (iTemp - 216) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Read In Disable"
                    Case 224 To 231
                        iByte = iByte + 9
                        iBit = (iTemp - 224) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Read In Disable"
                    Case 300 To 307
                        iByte = iByte + 10
                        iBit = (iTemp - 300) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "NC Start Disable"
                    Case 308 To 315
                        iByte = iByte + 11
                        iBit = (iTemp - 308) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "NC Start Disable"
                    Case 1100 To 1107
                        iByte = iByte + 12
                        iBit = (iTemp - 1100) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Stop GEO Axis 1"
                    Case 1108 To 1115
                        iByte = iByte + 13
                        iBit = (iTemp - 1108) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Stop GEO Axis 1"
                    Case 1200 To 1207
                        iByte = iByte + 14
                        iBit = (iTemp - 1200) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Stop GEO Axis 2"
                    Case 1208 To 1215
                        iByte = iByte + 15
                        iBit = (iTemp - 1208) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Stop GEO Axis 2"
                    Case 1300 To 1307
                        iByte = iByte + 16
                        iBit = (iTemp - 1300) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Stop GEO Axis 3"
                    Case 1308 To 1315
                        iByte = iByte + 17
                        iBit = (iTemp - 1308) Mod 8
                        lblAlType.Caption = lblAlType.Caption & "Feed Stop GEO Axis 3"
                    Case Else
                        lblAlType.Caption = "N.A."
                        txtDB2.Text = "--//--"
                        Exit Sub
                End Select
            Else
                txtDB2.Text = "--//--"
                If iChan > 8 And iChan < 11 Then
                    lblAlType.Caption = "Channels 9 and 10 Not implemented in SW.5"
                Else
                    lblAlType.Caption = "N.A."
                End If
                Exit Sub
            End If
        Case "6"
            If sAlarm5 <> "0" Then
                lblAlType.Caption = "N.A."
                txtDB2.Text = "--//--"
                Exit Sub
            End If
            iAxis = Int(Val(sAlarm4 & sAlarm3))
            If iAxis < 1 Or iAxis > 18 Then
                txtDB2.Text = "--//--"
                lblAlType.Caption = "Axes 18 ... 31 not implemented"
                Exit Sub
            End If
            iByte = 144 + ((iAxis - 1) * 2)
            lblAlType.Caption = "Axis " & Str(iAxis) & ": Feed Stop / Spindle Stop"
            iTemp = Int(Val(sAlarm2 & sAlarm1))
            Select Case iTemp
                Case 0 To 7
                    iBit = iTemp Mod 8
                Case 8 To 15
                    iByte = iByte + 1
                    iBit = (iTemp - 8) Mod 8
                Case Else
                    txtDB2.Text = "--//--"
                    lblAlType.Caption = "N.A."
                    Exit Sub
            End Select
        Case "7"
            lblAlType.Caption = "User"
            If sAlarm5 <> 0 Then
                 txtDB2.Text = "--//--"
                 lblAlType.Caption = "N.A."
                 Exit Sub
            End If
            iChan = Int(Val(sAlarm4 & sAlarm3))
            If iChan >= 0 And iChan <= 31 Then
                iByte = 180 + (iChan * 8)
            Else
                 txtDB2.Text = "--//--"
                 lblAlType.Caption = "N.A."
                 Exit Sub
            End If
            iAxis = Int(Val(sAlarm2 & sAlarm1))
            If iAxis >= 0 And iAxis <= 63 Then
                iByte = iByte + Int(iAxis / 8)
                iBit = iAxis Mod 8
            Else
                 txtDB2.Text = "--//--"
                 lblAlType.Caption = "N.A."
                 Exit Sub
            End If
        Case Else
            lblAlType.Caption = "N.A."
            txtDB2.Text = "--//--"
            Exit Sub
    End Select
    
    txtDB2.Text = "DB2 DBX" & Str(iByte) & "." & Trim$(Str(iBit))
    Exit Sub
ErrAlarm:
    frmAlCalc.Caption = Err.Description
    
End Sub

Private Sub txtDB2_KeyPress(KeyAscii As Integer)
On Error GoTo errDB2
    If KeyAscii <> 13 Then
        Exit Sub
    End If
'    Dim iAlarma As Long
    Dim iTemp, iLung, iAlarm, iAlarm2, iByte, iBit As Integer
    Dim sTemp, sByte, sBit As String
    
    frmAlCalc.Caption = "Alarm Calculator"
    lblAlType.Caption = ""
    txtAlarm.Text = ""
    iAlarm = 0
    sTemp = Trim$(txtDB2.Text)
    iLung = Len(sTemp)
    sByte = ""
    sBit = ""
    
    If (StrComp(Left$(sTemp, 3), "DB2", vbTextCompare) = 0) Then
        iTemp = InStr(1, sTemp, "DBX", vbTextCompare)
        If iTemp > 0 Then
            iTemp = iTemp + 2
            sTemp = Trim$(Right$(sTemp, iLung - iTemp))
        End If
    End If
    
    iTemp = InStr(1, sTemp, ".", vbTextCompare)
    iLung = Len(sTemp)
    If iTemp > 0 Then
        sByte = Left$(sTemp, iTemp - 1)
        sBit = Right$(sTemp, iLung - iTemp)
    End If
    iByte = Int(Val(sByte))
    iBit = Int(Val(sBit))
    sByte = Trim$(Str$(iByte))
    sBit = Trim$(Str$(iBit))
    
    If (iByte >= 0) And (iByte <= 999) And (iBit >= 0) And (iBit <= 7) Then
        txtDB2.Text = "DB2 DBX " & sByte & "." & sBit
    Else
        txtDB2.Text = ""
        lblAlType.Caption = "Err."
        Exit Sub
    End If
    Select Case iByte
        Case 0 To 143
            iAlarm2 = (Int(iByte / 18) + 1) * 10000
            iTemp = iByte - (Int(iByte / 18) * 18)
            Select Case iTemp
                Case 0 To 1
                    iAlarm2 = iAlarm2
                    iAlarm2 = iAlarm2 + (iTemp Mod 2) * 8 + iBit
                Case 2 To 5
                    iAlarm2 = iAlarm2 + 100
                    iAlarm2 = iAlarm2 + ((iTemp - 2) Mod 4) * 8 + iBit
                Case 6 To 9
                    iAlarm2 = iAlarm2 + 200
                    iAlarm2 = iAlarm2 + ((iTemp - 6) Mod 4) * 8 + iBit
                Case 10 To 11
                    iAlarm2 = iAlarm2 + 300
                    iAlarm2 = iAlarm2 + ((iTemp - 10) Mod 2) * 8 + iBit
                Case 12 To 13
                    iAlarm2 = iAlarm2 + 1100
                    iAlarm2 = iAlarm2 + ((iTemp - 12) Mod 2) * 8 + iBit
                Case 14 To 15
                    iAlarm2 = iAlarm2 + 1200
                    iAlarm2 = iAlarm2 + ((iTemp - 14) Mod 2) * 8 + iBit
                Case 16 To 17
                    iAlarm2 = iAlarm2 + 1300
                    iAlarm2 = iAlarm2 + ((iTemp - 16) Mod 2) * 8 + iBit
                Case Else
                    iAlarm2 = 0
            End Select
            iAlarm = 500000 + iAlarm2
        Case 144 To 179
            iAlarm2 = Int((iByte - 144) / 2) + 1
            iTemp = (iByte - 144) Mod 2
            iAlarm = 600000 + iAlarm2 * 100 + iTemp * 8 + iBit
        Case 180 To 435
            iAlarm2 = Int((iByte - 180) / 8)
            iTemp = (iByte - 180) Mod 8
            iAlarm = 700000 + (iAlarm2 * 100) + (iTemp * 8 + iBit)
        Case Else
            iAlarm = 0
    End Select
    
    If iAlarm > 0 Then
        txtAlarm.Text = Trim$(Str(iAlarm))
        txtAlarm_KeyPress (13)
    Else
        txtAlarm = ""
    End If
    
    Exit Sub
errDB2:
    frmAlCalc.Caption = Err.Description
End Sub
