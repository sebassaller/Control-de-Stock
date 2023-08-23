VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ayuda"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
      Height          =   1215
      Left            =   3360
      TabIndex        =   2
      Top             =   4080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2143
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form5.frx":C545
      PICN            =   "Form5.frx":C561
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
      Height          =   1575
      Left            =   6480
      TabIndex        =   1
      Top             =   3600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form5.frx":EDE2
      PICN            =   "Form5.frx":EDFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "Form5.frx":132FB
      PICN            =   "Form5.frx":13317
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "4=por cualquier consulta comunicarse cel(3454061198) facebook(sebas saller)e-mail(sebas_sallerQhotmail.com)"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   11535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3=siempre despues de actualizar y de nuevo actualizar la tabla para que los datos se graben bien"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   11535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2=el inventario funciona  bien   suma bien y resta bien solo que cuando se elimina  un registro  el numero de fila no se regenera"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1=si la base de datos genera error  cambiar el numero ide porq se repite"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   11295
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   6720
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   5520
      Shape           =   5  'Rounded Square
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   9000
      Shape           =   5  'Rounded Square
      Top             =   480
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Index           =   10
      Left            =   6480
      Shape           =   2  'Oval
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   4215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
zindex = 0
End Sub

Private Sub ChameleonBtn1_Click()
Form5.Hide
End Sub

