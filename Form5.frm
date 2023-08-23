VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ayuda"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   11400
      OleObjectBlob   =   "Form5.frx":C545
      Top             =   3960
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   615
      Left            =   8160
      TabIndex        =   0
      Top             =   6600
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
      MICON           =   "Form5.frx":C779
      PICN            =   "Form5.frx":C795
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   7215
      Left            =   0
      Picture         =   "Form5.frx":D06F
      ScaleHeight     =   7155
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   0
      Width           =   9615
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "e-mail:sebas_sallerQhotmail.com"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   6600
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "facebook/sebas saller"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " cel(3454061198)"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "3=siempre despues de actualizar y de nuevo actualizar la tabla para que los datos se graben bien"
         BeginProperty Font 
            Name            =   "Righteous"
            Size            =   14.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   0
         TabIndex        =   5
         Top             =   2160
         Width           =   8895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1=si la base de datos genera error  cambiar el numero ide porq se repite"
         BeginProperty Font 
            Name            =   "Righteous"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   8895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "por    consulta:"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "2=el inventario funciona  bien   suma bien y resta bien solo que cuando se elimina  un registro  el numero de fila no se regenera"
         BeginProperty Font 
            Name            =   "Righteous"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   8895
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   615
         Left            =   -120
         Shape           =   5  'Rounded Square
         Top             =   6600
         Width           =   735
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   735
         Left            =   480
         Shape           =   5  'Rounded Square
         Top             =   6480
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080C0FF&
         BorderStyle     =   0  'Transparent
         Height          =   615
         Left            =   960
         Shape           =   2  'Oval
         Top             =   6600
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   735
         Left            =   0
         Top             =   6000
         Width           =   975
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00808080&
         BorderColor     =   &H80000002&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   855
         Index           =   10
         Left            =   1560
         Shape           =   2  'Oval
         Top             =   6360
         Width           =   1215
      End
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

Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Skins\Office2007.skn"
Skin1.ApplySkin Form5.hWnd

End Sub
