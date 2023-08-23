VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "bienvenido"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      BTYPE           =   14
      TX              =   "Mercaderia/articulos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Form1.frx":0000
      PICN            =   "Form1.frx":001C
      PICH            =   "Form1.frx":08F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      BTYPE           =   14
      TX              =   "Inventario del dia"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Form1.frx":11D0
      PICN            =   "Form1.frx":11EC
      PICH            =   "Form1.frx":1AC6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      BTYPE           =   14
      TX              =   "Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Form1.frx":23A0
      PICN            =   "Form1.frx":23BC
      PICH            =   "Form1.frx":2C96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   4800
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form1.Hide


End Sub

Private Sub Command2_Click()
Form3.Show
Form1.Hide


End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub ChameleonBtn1_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub ChameleonBtn2_Click()
Form3.Show
Form1.Hide

End Sub

Private Sub ChameleonBtn3_Click()
End
End Sub

Private Sub Form_Load()
Image2.Stretch = True
Image2.Picture = LoadPicture(App.Path & "\vbWall.jpg")
Label1.ZOrder (0)
End Sub

Private Sub Form_Resize()
With Image2
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With
End Sub

Private Sub Label1_Click()
Label1.ZOrder (0)

End Sub
