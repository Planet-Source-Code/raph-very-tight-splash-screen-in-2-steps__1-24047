VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "EZ Open"
   ClientHeight    =   2325
   ClientLeft      =   1710
   ClientTop       =   1725
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2325
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   960
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   495
      Left            =   3840
      Shape           =   2  'Oval
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblRaph 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "by rAph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T  i  g  h  t . S  p  l  a  s  h"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   -360
      TabIndex        =   1
      Top             =   795
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Line L2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      X1              =   480
      X2              =   840
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      X1              =   5880
      X2              =   6240
      Y1              =   240
      Y2              =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tight Splash Screen by rAph
'June 13, 2001

'Hey, the moving line code came from someone at PSC, sorry,
'I don't know your name. Thanks for influencing this splash
'screen though. To customize this splash screen  just change
'lblRaph 's caption to your name, and change the const sName
'to your app title. If you want to use this splash go right
'ahead, just give me a little credit somewhere, and Drop me
'an e-mail so I can check out your prog! Hope you enjoy it,
'       -rAph

Dim XX1 As Integer
Dim XX2 As Integer
Dim YY1 As Integer
Dim YY2 As Integer
Dim XXX1 As Integer
Dim XXX2 As Integer
Dim YYY1 As Integer
Dim YYY2 As Integer

Dim When As Integer
Dim Start As Boolean
Dim I As Integer

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Const sname = "T  i  g  h  t  . S  p  l  a  s  h" '<-------------- change this to your app title (works better if you leave 2 spaces between each character
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
YY1 = L.Y1
YY2 = L.Y2
XX1 = L.X1
XX2 = L.X2
YYY1 = L2.Y1
YYY2 = L2.Y2
XXX1 = L2.X1
XXX2 = L2.X2
Start = False
I = 1
lblAppName = ""
lblRaph.Font = 29
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = False
Timer3.Enabled = False
End
End Sub

Private Sub lblAppName_DblClick()
Unload frmSplash
End Sub

Private Sub Timer2_Timer()
YY2 = YY2 - 100: If YY2 = 600 Then YY2 = 0
YY1 = YY1 + 100: If YY1 = 600 Then YY1 = 0
XX2 = XX2 - 100: If XX2 = 0 Then XX2 = 600
XX1 = XX1 - 100: If XX1 = 0 Then XX1 = 600
L.X1 = XX1
L.X2 = XX2
L.Y1 = YY1
L.Y2 = YY2
End Sub

Private Sub Timer3_Timer()
Dim s As Integer
YYY2 = YYY2 - 100: If YY2 = 0 Then YY2 = 600
YYY1 = YYY1 + 100: If YY1 = 600 Then YY1 = 0
XXX2 = XXX2 + 100: If XX2 = 600 Then XX2 = 0
XXX1 = XXX1 + 100: If XX1 = 600 Then XX1 = 0
L2.X1 = XXX1
L2.X2 = XXX2
L2.Y1 = YYY1
L2.Y2 = YYY2

If L.X1 = 3180 Then
    lblAppName.Visible = True
    Start = True
End If

If Start = True Then
    If L2.X1 = 6480 And L2.Y1 = 6120 Then
        FinishSplash
    ElseIf I = Len(sname) + 1 Then
        Exit Sub
    Else
        a = lblAppName
        b = Mid(sname, I, 1)
        a = a & b
        lblAppName = a
        I = I + 1
    End If
End If
End Sub


Sub FinishSplash()
lblRaph.Visible = True
lblRaph.FontSize = 29.5
Wait 0.5
lblRaph.FontSize = 29
Wait 0.5
lblRaph.FontSize = 28.5
Wait 0.5
lblRaph.FontSize = 28
Wait 0.5
lblRaph.FontSize = 27.5
Wait 0.5
lblRaph.FontSize = 27
Wait 0.5
lblRaph.FontSize = 26.5
Wait 0.5
lblRaph.FontSize = 26
Wait 0.5
lblRaph.FontSize = 25.5
Wait 0.5
lblRaph.FontSize = 25
Wait 0.5
lblRaph.FontSize = 24.5
Wait 0.5
lblRaph.FontSize = 24
Wait 0.5
lblRaph.FontSize = 23.5
Wait 0.5
lblRaph.FontSize = 23
Wait 0.5
lblRaph.FontSize = 22.5
Wait 0.5
lblRaph.FontSize = 22
Wait 0.5
lblRaph.FontSize = 21.5
Wait 0.5
lblRaph.FontSize = 21
Wait 0.5
lblRaph.FontSize = 20.5
Wait 0.5
lblRaph.FontSize = 20
Wait 0.5
lblRaph.FontSize = 19.5
Wait 0.5
lblRaph.FontSize = 19
Wait 0.5
lblRaph.FontSize = 18.5
Wait 0.5
lblRaph.FontSize = 18
Wait 0.5
lblRaph.FontSize = 17.5
Wait 0.5
lblRaph.FontSize = 17
Wait 0.5
lblRaph.FontSize = 16.5
Wait 0.5
lblRaph.FontSize = 16
Wait 0.5
lblRaph.FontSize = 15.5
Wait 0.5
lblRaph.FontSize = 15
Wait 0.5
lblRaph.FontSize = 14.5
Wait 0.5
lblRaph.FontSize = 14
Wait 0.5
lblRaph.FontSize = 13.5
Wait 0.5
lblRaph.FontSize = 13
Wait 0.5
lblRaph.FontSize = 12.5
Wait 0.5
Shape1.Visible = True
lblRaph.FontSize = 12
Wait 0.5
lblRaph.FontSize = 11.5
Wait 0.5
lblRaph.FontSize = 11
Wait 0.5
lblRaph.FontSize = 10.5
Wait 1.5
End
End Sub

