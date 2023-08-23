VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de Mercaderia"
   ClientHeight    =   8355
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":08CA
   ScaleHeight     =   8355
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text16 
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
      Left            =   7200
      TabIndex        =   30
      Text            =   "buscar="
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text15 
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
      Left            =   2520
      TabIndex        =   29
      Text            =   "vencimiento"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text14 
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
      Left            =   5160
      TabIndex        =   28
      Text            =   "seccion"
      Top             =   2160
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "Form2.frx":14A18
      Top             =   2280
   End
   Begin ChamaleonButton.ChameleonBtn anterior 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   7440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "anterior"
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
      MICON           =   "Form2.frx":14C4C
      PICN            =   "Form2.frx":14C68
      PICH            =   "Form2.frx":15542
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn siguente 
      Height          =   855
      Left            =   9720
      TabIndex        =   18
      Top             =   7440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "siguiente"
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
      MICON           =   "Form2.frx":15E1C
      PICN            =   "Form2.frx":15E38
      PICH            =   "Form2.frx":16712
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton seccion 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton articulo 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2280
      TabIndex        =   16
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "vencimiento"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8760
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":16FEC
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   4210752
      ForeColor       =   8421504
      HeadLines       =   1
      RowHeight       =   29
      RowDividerStyle =   5
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   2280
      Top             =   7440
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   4210752
      ForeColor       =   12632256
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":17001
      OLEDBString     =   $"Form2.frx":1708D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "maercaderiaRN"
      Caption         =   "      mover registros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Edicion"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5295
      Left            =   12000
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
      Begin ChamaleonButton.ChameleonBtn nuevo 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         BTYPE           =   14
         TX              =   "Nuevo"
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
         BCOL            =   16744576
         BCOLO           =   16711680
         FCOL            =   16744576
         FCOLO           =   16744576
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Form2.frx":17119
         PICN            =   "Form2.frx":17135
         PICH            =   "Form2.frx":17A0F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn editar 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         BTYPE           =   14
         TX              =   "editar"
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
         MICON           =   "Form2.frx":182E9
         PICN            =   "Form2.frx":18305
         PICH            =   "Form2.frx":18BDF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn eliminar 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Form2.frx":194B9
         PICN            =   "Form2.frx":194D5
         PICH            =   "Form2.frx":19DAF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn guardar 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         BTYPE           =   14
         TX              =   "Guardar"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form2.frx":1A689
         PICN            =   "Form2.frx":1A6A5
         PICH            =   "Form2.frx":1AF7F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn cancelar 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         BTYPE           =   14
         TX              =   "cancelar"
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
         MICON           =   "Form2.frx":1B859
         PICN            =   "Form2.frx":1B875
         PICH            =   "Form2.frx":1C14F
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
      Caption         =   "Detalle de articulos"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   0
      MousePointer    =   4  'Icon
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   120
      Width           =   15255
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
         Left            =   11520
         TabIndex        =   27
         Text            =   "ID-produc."
         Top             =   480
         Width           =   2295
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
         Left            =   8160
         TabIndex        =   26
         Text            =   "Precio sug."
         Top             =   480
         Width           =   2055
      End
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
         Left            =   10200
         TabIndex        =   25
         Text            =   "vencimiento"
         Top             =   1320
         Width           =   2295
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
         Left            =   5400
         TabIndex        =   24
         Text            =   "Fecha de Hoy"
         Top             =   1320
         Width           =   2535
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
         Left            =   5400
         TabIndex        =   23
         Text            =   "Precio"
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
         Left            =   0
         TabIndex        =   22
         Text            =   "secciom"
         Top             =   1200
         Width           =   1575
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
         TabIndex        =   21
         Text            =   "articulo"
         Top             =   480
         Width           =   1455
      End
      Begin ChamaleonButton.ChameleonBtn calendario 
         Height          =   615
         Left            =   14400
         TabIndex        =   20
         Top             =   1200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "Form2.frx":1CA29
         PICN            =   "Form2.frx":1CA45
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8040
         TabIndex        =   13
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         DataField       =   "Id"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13920
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "seccion"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         ItemData        =   "Form2.frx":1D31F
         Left            =   1680
         List            =   "Form2.frx":1D344
         TabIndex        =   5
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         DataField       =   "articulo"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         DataField       =   "precion"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataField       =   "vencimiento"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12600
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Line Line7 
         X1              =   5280
         X2              =   5280
         Y1              =   1920
         Y2              =   0
      End
      Begin VB.Line Line1 
         X1              =   -600
         X2              =   14880
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   735
      Left            =   3000
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   8400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9720
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Menu BTactualizar 
      Caption         =   "actualizar"
   End
   Begin VB.Menu BTayuda 
      Caption         =   "ayuda"
   End
   Begin VB.Menu BTcalcular 
      Caption         =   "calcular precio"
   End
   Begin VB.Menu BTvolver 
      Caption         =   "volver"
   End
   Begin VB.Menu BTsalir 
      Caption         =   "salir"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public opcion As String
Private Sub ChameleonBtn1_Click()
Adodc1.Refresh


End Sub

Private Sub ChameleonBtn2_Click()
Adodc1.Recordset.Update

End Sub

Private Sub ChameleonBtn3_Click()
Adodc1.Recordset.Delete

End Sub

Private Sub ChameleonBtn4_Click()

End Sub

Private Sub ChameleonBtn5_Click()
Form1.Show
Form2.Hide

End Sub

Private Sub ChameleonBtn6_Click()
End
End Sub

Private Sub anterior_Click()
Adodc1.Recordset.MovePrevious
If (Adodc1.Recordset.BOF = False) Then
If Adodc1.Recordset(4).Value = Text6.Text Then
    MsgBox ("mercaderia en vencimiento")
    Form2.SetFocus
End If
    
End If
If (Adodc1.Recordset.BOF = True) Then
    Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub BTactualizar_Click()

''f comprobar <> 0 Then
''Adodc1.Recordset.MoveFirst
''While Not (Adodc1.Recordset.EOF = True)
''If UCase(Adodc1.Recordset(1)) = "" And (Adodc1.Recordset(2) = "") Then
''MsgBox ("registro existente")

 '' Exit Sub
 '' End If
''Adodc1.Recordset.MoveNext
 ''Wend
'MsgBox ("registro no encontrado")
''Adodc1.Recordset.MoveFirst
'' End If











Adodc1.Refresh
End Sub

Private Sub BTayuda_Click()
Form5.Show
End Sub

Private Sub BTcalcular_Click()
Text7.Text = (Val(Text3.Text) + (Val(Text3.Text) * 0.3))
End Sub

Private Sub BTsalir_Click()
End
End Sub

Private Sub BTvolver_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub calendario_Click()
Form4.Show
End Sub

Private Sub cancelar_Click()
''Adodc1.Recordset.Cancel

''If Adodc1.Recordset(0) = "" Then
''Adodc1.Recordset.CancelUpdate
''Adodc1.Recordset.MoveFirst
''End If

''Adodc2.Recordset.CancelUpdate
''Adodc2.Recordset.MoveFirst



''Adodc1.Recordset.Cancel
''Adodc1.Recordset.MoveLast
Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.MovePrevious
Adodc1.Refresh



Text4.Enabled = False
Text3.Enabled = False
Text1.Enabled = False
Combo1.Enabled = False
guardar.Enabled = False
editar.Enabled = True
eliminar.Enabled = True
nuevo.Enabled = True






End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
''If (Combo1.ListIndex = -1) Then
''Combo1.BackColor = red
''Else
''Combo1.BackColor = Color.white
''End If
If KeyPress = 32 Then
Text1.SetFocus
End If
End Sub

Private Sub DataGrid1_Click()
DataGrid1.Columns(0).Alignment = dbgCenter
''DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Width = 500
DataGrid1.EditActive = False
DataGrid1.BorderStyle = dbgNoBorder




End Sub

Private Sub editar_Click()



Text4.Enabled = True
Text3.Enabled = True
Text1.Enabled = True
Combo1.Enabled = True
guardar.Enabled = True
editar.Enabled = True
eliminar.Enabled = False
nuevo.Enabled = False


End Sub

Private Sub eliminar_Click()


''If Adodc1.Recordset.EOF = True Then
''Adodc1.Recordset.Cancel

''End If
''If Adodc1.Recordset.BOF = True Then
''''Adodc1.Recordset.Cancel
''End If
If Adodc1.Recordset.BOF = False Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
Else
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.Cancel
End If
End If

If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.Cancel
''Else

''If Adodc1.Recordset.EOF = False Then

''Adodc1.Recordset.Delete
''Adodc1.Recordset.MovePrevious
End If


End Sub

Private Sub Form_Load()
Image1.Stretch = True
Image1.Picture = LoadPicture(App.Path & "\g520up.jpg")

Adodc1.ConnectionString = "provider=microsoft.jet.oledb.4.0;" & "data source=" & App.Path & "\Database2.mdb"
Adodc1.CursorType = adOpenDynamic
Adodc1.RecordSource = "maercaderiaRN"
Adodc1.Refresh
Text6.Text = Format$(Date, "dd/mm/yyyy")

''If comprobar <> 0 Then
''Adodc1.Recordset.MoveFirst
''While Not (Adodc1.Recordset.EOF = True)
''If UCase(Adodc1.Recordset(1)) = "" And (Adodc1.Recordset(2) = "") Then
''MsgBox ("registro existente")

 '' Exit Sub
 '' End If
''Adodc1.Recordset.MoveNext
 ''Wend
'MsgBox ("registro no encontrado")
''Adodc1.Recordset.MoveFirst
'' End If

Text4.Enabled = False
Text3.Enabled = False
Text1.Enabled = False
Combo1.Enabled = False
guardar.Enabled = False

DataGrid1.Columns(0).Alignment = dbgCenter
''DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Width = 500
DataGrid1.BorderStyle = dbgNoBorder

Skin1.LoadSkin App.Path & "\Skins\Office2007.skn"
Skin1.ApplySkin Form2.hWnd



''Dim i As Integer
''Label10 = ""
''On Error GoTo 2147467259(80004005)
''i = Rnd * 10 ^ 16
''Label10 = Label10 & "esta instruccion no se ejecuta"
'' If Err.Number = 2147467259 =80004005 goto)  then
''MsgBox ("se ha producido un error" & Err & "descripcion:" & Err.Description)
''Label10 = Label10 & "la ejecucion continua"
''End If



End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With
End Sub

Private Sub guardar_Click()
If Adodc1.Recordset(0) = Null Then
Adodc1.Recordset.CancelUpdate
End If


Text4.Enabled = False
Text3.Enabled = False
Text1.Enabled = False
Combo1.Enabled = False
guardar.Enabled = False
editar.Enabled = True
nuevo.Enabled = True
eliminar.Enabled = True
Dim comprobar As Integer
comprobar = Text5.Text



''If comprobar <> 0 Then
'' Adodc1.Recordset.MoveFirst
''While Not (Adodc1.Recordset.EOF = True)
'' If UCase(comprobar) = Adodc1.Recordset(1) Then
' MsgBox ("registro existente")
'' comprobar = Int(Rnd * 999)
'' Text5.Text = comprobar
''  Exit Sub
'' End If
''Adodc1.Recordset.MoveNext
'' Wend
 ''MsgBox ("registro no encontrado")
 ''Adodc1.Recordset.MoveFirst
'' End If









Adodc1.Recordset.Update

Form2.SetFocus


End Sub

Private Sub Label1_Click(Index As Integer)
Label1.Item = 0
End Sub

Private Sub nuevo_Click()
Dim num1 As Single


Text4.Enabled = True
Text3.Enabled = True
Text1.Enabled = True
Combo1.Enabled = True
guardar.Enabled = True
editar.Enabled = False
eliminar.Enabled = False
Text4.SetFocus

Adodc1.Recordset.AddNew
If num1 = 0 Then
num1 = Int(Rnd * 999 + 3) - 1
If num1 > 500 Then
num1 = num1 / 2 - 3 + 1
If num1 > 200 Then
num1 = num1 + (Rnd * 9) / 2 + 1
End If
End If
End If

''num1 = Int(Rnd * 2) - 1 * (Rnd * 1)


Text5.Text = num1



End Sub

Private Sub refrescar_Click()
Dim preciosugerido, pre As Double
preciosugerido = Val(Text3.Text) * 0.3
Text7.Text = preciosugerido + Val(Text3.Text)
Adodc1.Refresh
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub siguente_Click()
Adodc1.Recordset.MoveNext
If (Adodc1.Recordset.EOF = True) Then
    Adodc1.Recordset.MoveLast
End If
If (Adodc1.Recordset.EOF = False) Then
If Adodc1.Recordset(4).Value = Text6.Text Then
    MsgBox ("mercaderia en vencimiento")
    Form2.SetFocus
End If
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 47) Then
If KeyAscii = 13 Then
End If
Else
Beep
KeyAscii = 0
End If

End Sub

Private Sub Text2_Change()
If articulo.Value = False And seccion.Value = False Then
With Adodc1
opcion = "articulo"
End With
End If

If articulo.Value = True Then
With Adodc1
    opcion = "vencimiento"
End With
Else
If seccion.Value = True Then
With Adodc1
opcion = "seccion"
End With
End If
End If




With Adodc1
 If Text2.Text <> "" Then
    .Recordset.Filter = opcion & " LIKE '*" + Text2 + "*'"
 Set DataGrid1.DataSource = Adodc1.Recordset
 If Adodc1.Recordset.EOF = False Then
 End If
Else
    Adodc1.Recordset.MoveFirst
    Set DataGrid1.DataSource = Adodc1.Recordset
    Form2.SetFocus
     If Adodc1.Recordset.EOF = False Then
    End If
End If
    .Refresh
End With

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46) Then
If KeyAscii = 32 Then
Combo1.SetFocus
End If
Else
Beep
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
Text3.SetFocus
End If
End Sub

Private Sub Text7_Change()
Text7.Text = Val(Text3.Text) * 0.3
End Sub

Private Sub volver_Click()
Form1.Show
Form2.Hide

End Sub
