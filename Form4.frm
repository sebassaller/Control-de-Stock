VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "fecha"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":08CA
   ScaleHeight     =   3810
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "Form4.frx":14A18
      Top             =   3000
   End
   Begin VB.TextBox textofecha 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin ChamaleonButton.ChameleonBtn volver 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bauhaus 93"
         Size            =   14.25
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
      MICON           =   "Form4.frx":14C4C
      PICN            =   "Form4.frx":14C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSACAL.Calendar calendario 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _Version        =   524288
      _ExtentX        =   12515
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   12632256
      Year            =   2020
      Month           =   3
      Day             =   10
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3=con el boton volver dejaguardado la fecha para  almacenar en  base de datos"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   6945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2=capture la fecha elegida para su almacenamiento"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   4545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1=seleccionar  fecha  a elegir"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   2595
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calendario_Click()

textofecha.Text = calendario.Value
End Sub

Private Sub Command1_Click()
textofecha.Text = fecha
End Sub



Private Sub Form_Load()
Image1.Stretch = True
Image1.Picture = LoadPicture(App.Path & "\g520up.jpg")
Skin1.LoadSkin App.Path & "\Skins\Office2007.skn"
Skin1.ApplySkin Form4.hWnd

End Sub

Private Sub volver_Click()
Form2.Show
Form4.Hide
Form2.Text1.Text = Form4.textofecha.Text

End Sub
