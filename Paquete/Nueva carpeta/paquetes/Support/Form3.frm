VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de inventario"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13770
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   5280
      TabIndex        =   21
      Text            =   "Sub-Total"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Text            =   "Total"
      Top             =   6360
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   840
      OleObjectBlob   =   "Form3.frx":08CA
      Top             =   6120
   End
   Begin MSFlexGridLib.MSFlexGrid inventario 
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   200
      Cols            =   5
      BackColor       =   8421504
      ForeColor       =   16777215
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      ForeColorSel    =   14737632
      BackColorBkg    =   14737632
      GridColor       =   16777215
      GridColorFixed  =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   13
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   12
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   6015
      Left            =   10920
      TabIndex        =   6
      Top             =   120
      Width           =   2655
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
         Height          =   1095
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "editar agregado"
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1931
         BTYPE           =   14
         TX              =   "eliminar tabla"
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
         BCOL            =   16761024
         BCOLO           =   16761024
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Form3.frx":0AFE
         PICN            =   "Form3.frx":0B1A
         PICH            =   "Form3.frx":13F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   1095
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "agregar producto a la tabla"
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1931
         BTYPE           =   14
         TX              =   "Agregar"
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
         MICON           =   "Form3.frx":1CCE
         PICN            =   "Form3.frx":1CEA
         PICH            =   "Form3.frx":25C4
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
         Height          =   1095
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "eliminar registro"
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1931
         BTYPE           =   14
         TX              =   "Eliminar"
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
         BCOL            =   16761024
         BCOLO           =   16761024
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Form3.frx":2E9E
         PICN            =   "Form3.frx":2EBA
         PICH            =   "Form3.frx":3794
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn4 
         Height          =   1095
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   "volver menu"
         Top             =   3720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1931
         BTYPE           =   14
         TX              =   "volver"
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
         BCOL            =   16761024
         BCOLO           =   16761024
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Form3.frx":406E
         PICN            =   "Form3.frx":408A
         PICH            =   "Form3.frx":4964
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn5 
         Height          =   1095
         Left            =   360
         TabIndex        =   11
         ToolTipText     =   "salir"
         Top             =   4800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1931
         BTYPE           =   14
         TX              =   "salir"
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
         BCOL            =   16761024
         BCOLO           =   16761024
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Form3.frx":523E
         PICN            =   "Form3.frx":525A
         PICH            =   "Form3.frx":5B34
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "productos"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10695
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   6840
         TabIndex        =   19
         Text            =   "fecha"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   5280
         TabIndex        =   18
         Text            =   "Buscar"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   3720
         TabIndex        =   17
         Text            =   "cantidad"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox art 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Text            =   "articulo"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Text            =   "Precio"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   1200
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c As Long
Private Sub LaVolpeButton1_Click()

End Sub


Private Sub ChameleonBtn1_Click()



For c = 1 To inventario.Row + 1
inventario.TextMatrix(c, 0) = c


Next

If fila = 0 Then
fila = 1
End If

inventario.Col = 1
inventario.Row = fila
inventario.Text = Text1.Text
inventario.Col = 2
inventario.Row = fila
inventario.Text = Text2.Text
inventario.Col = 3
inventario.Row = fila
inventario.Text = Text4.Text
inventario.Col = 4
inventario.Row = fila
inventario.Text = Text2.Text

X = Val(Text2.Text)
total = total + X
Text6.Text = total

''total = total + x
''Text6.Text = total

Text7.Text = X

fila = fila + 1
Text7.Text = Text2.Text
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus

End Sub

Private Sub ChameleonBtn1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
End Sub

Private Sub ChameleonBtn2_Click()
inventario.Clear

inventario.Col = 1
inventario.Row = 0
inventario.ColWidth(1) = 5000
inventario.Text = "articulo"
inventario.Col = 2
inventario.Row = 0
inventario.Text = "precio"
inventario.Col = 3
inventario.Row = 0
inventario.Text = "fecha"

fila = 1

inventario.Col = 4
inventario.Row = 0
inventario.Text = "SUB-Total"
Text6.Text = ""
Text7.Text = ""
X = 0
total = 0
End Sub

Private Sub ChameleonBtn3_Click()
total = Text6.Text
X = 0

If fila = 0 Then
MsgBox ("no pose registros")
fila = fila + 1
Text7.Text = ""
Text6.Text = ""
total = 0
X = 0
End If

''Text6.Text = total
fila = fila - 1
''If Val(Text6.Text) < 0 Then
''Text6.Text = "0"
''End If

resta = Val(inventario.TextMatrix(inventario.Row, 2))


total = total - resta
Text6.Text = total
Text7.Text = ""





For c = 1 To inventario.Row - 1
inventario.TextMatrix(c, 0) = c


Next


''inventario.Col = 0
''inventario.Row = fila
''inventario.Text = ""
''inventario.Col = 1
''inventario.Row = fila
''inventario.Text = ""
''inventario.Col = 2
''inventario.Row = fila
''inventario.Text = ""
''inventario.Col = 3
''inventario.Row = fila
''inventario.Text = ""
''inventario.Col = 4
''inventario.Row = fila
''inventario.Text = ""

''resta = Val(Text6.Text) - Val(inventario.TextMatrix(inventario.Row, 2))
''Text6.Text = resta
''x = inventario.TextMatrix(inventario.Row, 2)
''Text7.Text = x
inventario.RemoveItem (inventario.Row)
Text1.SetFocus



End Sub

Private Sub ChameleonBtn4_Click()
Form1.Show
Form3.Hide
total = 0
X = 0
fila = 0

End Sub

Private Sub ChameleonBtn5_Click()
End
End Sub


Private Sub Form_GotFocus()
If KeyAscii = 13 Then
ChameleonBtn1.SetFocus
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ChameleonBtn1.SetFocus
End If

End Sub

Private Sub Form_Load()
Image1.Stretch = True
Image1.Picture = LoadPicture(App.Path & "\g520up.jpg")
Text4.Text = Format$(Date, "dd/mm/yy")



inventario.ColWidth(0) = 500
inventario.Col = 1
inventario.Row = 0
inventario.ColWidth(1) = 5000
inventario.Text = "articulo"
inventario.Col = 2
inventario.Row = 0
inventario.Text = "precio"
inventario.Col = 3
inventario.Row = 0
inventario.Text = "fecha"


fila = 1

inventario.Col = 4
inventario.Row = 0
inventario.Text = "SUB-Total"
Text7.Enabled = False
Text6.Enabled = False

Skin1.LoadSkin App.Path & "\Skins\Office2007.skn"
Skin1.ApplySkin Form3.hWnd



End Sub

Private Sub Form_LostFocus()
If KeyAscii = 13 Then
ChameleonBtn1.SetFocus
End If

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
Text2.SetFocus
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46) Then
If KeyAscii = 32 Then
Text3.SetFocus
End If
Else
Beep
KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46) Then


If KeyAscii = 32 Then
ChameleonBtn1.SetFocus
End If
Else
Beep
KeyAscii = 0
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 47) Then
If KeyAscii = 13 Then
End If
Else
Beep
KeyAscii = 0
End If


End Sub
