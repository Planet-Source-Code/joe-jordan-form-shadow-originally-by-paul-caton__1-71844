VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Shadow Demo"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScrollTrans 
      Height          =   375
      Left            =   2220
      Max             =   255
      TabIndex        =   5
      Top             =   1500
      Value           =   2
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll 
      Height          =   375
      Left            =   2220
      Max             =   100
      Min             =   2
      TabIndex        =   2
      Top             =   660
      Value           =   2
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicShadowColor 
      Height          =   1035
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1635
   End
   Begin VB.CommandButton cmdSetShadow 
      Caption         =   "Set Shadow Color"
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   1500
      Width           =   1650
   End
   Begin VB.Label lblTrans 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4020
      TabIndex        =   7
      Top             =   1140
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow Transparency:"
      Height          =   195
      Left            =   2220
      TabIndex        =   6
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label lblShadowSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow Size:"
      Height          =   195
      Left            =   2220
      TabIndex        =   3
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shadow As clsShadow

Private Sub cmdSetShadow_Click()
    Call PicShadowColor_Click
End Sub

Private Sub Form_Load()
    ' Set Shadow to form
    Set Shadow = New clsShadow
    Call Shadow.Shadow(Me)
    
    ' Get default color
    PicShadowColor.BackColor = Shadow.Color
    
    ' Get default shadow size
    If Shadow.Depth >= 2 And Shadow.Depth < 101 Then
        HScroll.Value = Shadow.Depth
        lblShadowSize.Caption = Shadow.Depth
    End If

    ' Get default transparency size
    If Shadow.Transparency >= 0 And Shadow.Depth < 256 Then
        HScrollTrans.Value = Shadow.Transparency
        lblTrans.Caption = Shadow.Transparency
    End If
End Sub

Private Sub HScroll_Change()
    Call HScroll_Scroll
End Sub

Private Sub HScrollTrans_Change()
    Call HScrollTrans_Scroll
End Sub

Private Sub HScrollTrans_Scroll()
    Shadow.Transparency = HScrollTrans.Value
    lblTrans.Caption = Shadow.Transparency
    Call Shadow.Shadow(Me)
End Sub

Private Sub PicShadowColor_Click()
    dlg.ShowColor
    PicShadowColor.BackColor = dlg.Color
    Shadow.Color = dlg.Color
End Sub

Private Sub HScroll_Scroll()
    Shadow.Depth = HScroll.Value
    lblShadowSize.Caption = Shadow.Depth
    Call Shadow.Shadow(Me)
End Sub
