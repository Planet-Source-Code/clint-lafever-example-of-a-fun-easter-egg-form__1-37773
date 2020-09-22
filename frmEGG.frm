VERSION 5.00
Begin VB.Form frmEGG 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Credits (Easter Egg)"
   ClientHeight    =   4155
   ClientLeft      =   795
   ClientTop       =   480
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   27.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSOUND 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6960
      TabIndex        =   1
      Top             =   3840
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   3360
      Top             =   3000
   End
   Begin VB.Label lblLABEL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to begin."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   7695
   End
End
Attribute VB_Name = "frmEGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public QuitIt As Boolean
Private Const INDELAY = 1
Private Const OUTDELAY = 1
Private Const ColorStep = 2

'------------------------------------------------------------
' Here is the string to edit to have the credits
' you want.  Use a tilde to indicate carriage returns.
'  Seperate with semi colons.
'------------------------------------------------------------
Private Const CreditString As String = "Credits:;Lead Developer:;Clint LaFever~aka: DaVBMan;;Program Manager:;Franco Tao~aka: The ""Real "" Boss;;Product Testing:;Andrew Englen~Wade Hampton~aka: ""dunno"";;Documentation:;Clint LaFever~aka: I do it all.;;Requirements:;Keith Wheeler~aka: The Client;;Special Thanks:;Danielle LaFever~aka: My Wife!;I had to~get her name~in here somehow.;;The End...;;Fin;That's All Folks;Goodbye;Adios;See Ya;You can leave now;Anytime today now...;Click Close;Fine then.~Bye;;;You're still here;Go Home;Scram;Beat It;Leave me alone.;Stop Stalking Me!;¿#!@;;;Elvis has left~the building.;;;;;Still Watching?~You have no life.;;Get back to work!;;Tax dollars~at work huh?;;;Thank You~Armed Forces.;God Bless America!;Amen.;;;"
Private Sub DoIt()
    On Error Resume Next
    Dim h As String, v As String, OldTime As Long, x As Long, vs As String, z As Long, xo As Long
    Dim R As Long, g As Long, b As Long, v2 As String, vh As String
    h = CreditString
    While h <> "" And Me.QuitIt = False
        '------------------------------------------------------------
        ' fade in
        '------------------------------------------------------------
        v = GetToken(h, ";")
        vh = v
        vs = v
        v = Replace(v, "~", vbCrLf)
        R = 0: g = 0: b = 0
        Do While R < 255 And Me.QuitIt = False
            vh = vs
            OldTime = GetTickCount()
            Me.ForeColor = RGB(R, g, b)
            Me.CurrentY = Me.ScaleHeight / 2 - (Me.TextHeight(v) / 2)
            For x = 1 To CountTokens(vh, "~")
                v2 = GetToken(vh, "~")
                Me.CurrentX = Me.ScaleWidth / 2 - (Me.TextWidth(v2) / 2)
                Me.Print v2
            Next x
            R = R + ColorStep: g = g + ColorStep: b = b + ColorStep
            If R > 255 Then R = 255: If g > 255 Then g = 255: If b > 255 Then b = 255
            Do While OldTime + INDELAY > GetTickCount():  Loop
            LineIt
        Loop
        '------------------------------------------------------------
        ' shake
        '------------------------------------------------------------
        For z = 1 To 20
            vh = vs
            R = 255: g = 255: b = 255
            OldTime = GetTickCount()
            Me.ForeColor = RGB(R, g, b)
            xo = (CInt(Rnd * 600) + 1) - 300
            Me.CurrentY = (Me.ScaleHeight / 2 - (Me.TextHeight(v) / 2)) + (CInt(Rnd * 600) + 1) - 300
            For x = 1 To CountTokens(vh, "~")
                v2 = GetToken(vh, "~")
                Me.CurrentX = Me.ScaleWidth / 2 - (Me.TextWidth(v2) / 2) + xo
                Me.Print v2
            Next x
            Me.Print " "
            R = R + ColorStep: g = g + ColorStep: b = b + ColorStep
            If R > 255 Then R = 255: If g > 255 Then g = 255: If b > 255 Then b = 255
            Do While OldTime + 20 > GetTickCount():  Loop
            Me.Cls
            LineIt
        Next z
        '------------------------------------------------------------
        ' fade out
        '------------------------------------------------------------
        DoEvents
        R = 255: g = 255: b = 255
        Do While R > 0 And Me.QuitIt = False
            vh = vs
            OldTime = GetTickCount()
            Me.ForeColor = RGB(R, g, b)
            Me.CurrentY = Me.ScaleHeight / 2 - (Me.TextHeight(v) / 2)
            For x = 1 To CountTokens(vh, "~")
                v2 = GetToken(vh, "~")
                Me.CurrentX = Me.ScaleWidth / 2 - (Me.TextWidth(v2) / 2)
                Me.Print v2
            Next x
            R = R - ColorStep: g = g - ColorStep: b = b - ColorStep
            If R < 0 Then R = 0: If g < 0 Then g = 0: If b < 0 Then b = 0
            Do While OldTime + OUTDELAY > GetTickCount():  Loop
            LineIt
        Loop
        DoEvents
    Wend
    If Me.QuitIt = False Then Me.lblLABEL.Visible = True
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: August,09 2002 @ 11:26:19
'------------------------------------------------------------
Public Sub LineIt()
    On Error GoTo ErrorLineIt
    Static y1 As Long, y2 As Long, x As Long, Drawn As Boolean
    If Drawn = False Then
        Drawn = True
        y1 = Int(Rnd * Me.Height) + 1
        y2 = Int(Rnd * Me.Height) + 1
        x = Int(Rnd * Me.Width) + 1
        Me.Line (x, y1)-(x, y2), 13619151
    Else
        Drawn = False
        Me.Line (x, y1)-(x, y2), vbBlack
    End If
    Exit Sub
ErrorLineIt:
    MsgBox Err & ":Error in call to LineIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub chkSOUND_Click()
    On Error Resume Next
    If chkSOUND.Value = 0 Then
        PlayWaveRes sndFOLD
    Else
        PlayWaveRes sndEGG, soundASYNC Or soundLOOP
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.QuitIt = False
    PlayWaveRes sndEGG, soundASYNC Or soundLOOP
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    PlayWaveRes sndFOLD
    Me.QuitIt = True
End Sub
Private Sub lblLABEL_Click()
    On Error Resume Next
    Me.lblLABEL.Visible = False
    DoIt
End Sub
Private Sub Form_Click()
    On Error Resume Next
    If Me.lblLABEL.Visible = True Then
        Me.lblLABEL.Visible = False
        DoIt
    End If
End Sub
Private Sub Timer_Timer()
    On Error Resume Next
    Me.Timer.Enabled = False
    Me.Refresh
    DoIt
End Sub




