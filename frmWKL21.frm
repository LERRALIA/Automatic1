VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frmWKL21 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Tagesabschluss"
   ClientHeight    =   8595
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11910
   StartUpPosition =   1  'Fenstermitte
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   9840
      TabIndex        =   104
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pbr1 
      Height          =   255
      Left            =   7800
      TabIndex        =   102
      Top             =   45
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   6000
      TabIndex        =   93
      Top             =   1200
      Visible         =   0   'False
      Width           =   6720
      Begin VB.FileListBox File3 
         Height          =   285
         Left            =   0
         Pattern         =   "*.DBF"
         TabIndex        =   107
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "nur diese Kasse"
         Height          =   210
         Left            =   6720
         TabIndex        =   106
         Top             =   7440
         Width           =   1575
      End
      Begin VB.CheckBox check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ohne EK - Zahlen "
         Height          =   210
         Left            =   1440
         TabIndex        =   105
         Top             =   7440
         Width           =   1695
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   7
         Left            =   8430
         TabIndex        =   100
         Top             =   7680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "nach AGN"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   6
         Left            =   6600
         TabIndex        =   99
         Top             =   7680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "nach Lieferanten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   10
         Left            =   4890
         TabIndex        =   98
         Top             =   7680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "nicht umsatz- relevante VKs"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12515
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   13
         Left            =   3300
         TabIndex        =   101
         Top             =   7680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Artikel detailliert"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   12
         Left            =   1350
         TabIndex        =   97
         Top             =   7680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Artikel kumuliert"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   96
         Top             =   7680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Z - Bon"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Left            =   10020
         TabIndex        =   95
         Top             =   7680
         Width           =   1720
         _ExtentX        =   3043
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   0
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "Kundenliste.doc"
      PrintFileType   =   24
      PrintFileLinesPerPage=   60
      WindowShowExportBtn=   0   'False
   End
   Begin sevCommand3.Command SSCommand1 
      Height          =   495
      Index           =   5
      Left            =   8610
      TabIndex        =   127
      Top             =   7920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Kasse"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command SSCommand1 
      Height          =   495
      Index           =   1
      Left            =   10200
      TabIndex        =   128
      Top             =   7920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command SSCommand1 
      Height          =   495
      Index           =   0
      Left            =   7020
      TabIndex        =   129
      Top             =   7920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Tagesprotokoll"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command SSCommand1 
      Height          =   495
      Index           =   2
      Left            =   5430
      TabIndex        =   130
      Top             =   7920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Terminalschnitt"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   55
      Left            =   10080
      TabIndex        =   134
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kreditkarte, nicht umsr. Verkäufe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   55
      Left            =   6600
      TabIndex        =   133
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label neuLbl 
      BackColor       =   &H00C0C000&
      Caption         =   "neuerAbschluss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   132
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label altLbl 
      BackColor       =   &H00C0C000&
      Caption         =   "alterAbschluss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   131
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "daraus Fremdgutscheine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   54
      Left            =   360
      TabIndex        =   126
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   54
      Left            =   3600
      TabIndex        =   125
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Gesamtrabatt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   6600
      TabIndex        =   18
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Artikelrabatt Anz.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   53
      Left            =   6600
      TabIndex        =   124
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kreditkarte,insgesamt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   52
      Left            =   6600
      TabIndex        =   122
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   52
      Left            =   10080
      TabIndex        =   121
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   51
      Left            =   2280
      TabIndex        =   120
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Dukaten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   51
      Left            =   240
      TabIndex        =   119
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   50
      Left            =   3720
      TabIndex        =   118
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   49
      Left            =   10080
      TabIndex        =   117
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kassendifferenz, aufgelaufene"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   50
      Left            =   6600
      TabIndex        =   116
      Top             =   7320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   48
      Left            =   9600
      TabIndex        =   115
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kassendifferenz jetzt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   49
      Left            =   6600
      TabIndex        =   114
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   47
      Left            =   3600
      TabIndex        =   113
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   46
      Left            =   3600
      TabIndex        =   112
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Wechselgeld"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   48
      Left            =   600
      TabIndex        =   111
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "- Abschöpfung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   47
      Left            =   600
      TabIndex        =   110
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Gutschein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   46
      Left            =   360
      TabIndex        =   109
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   45
      Left            =   3600
      TabIndex        =   108
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   44
      Left            =   3840
      TabIndex        =   92
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "nicht umsatzrelevante Verkäufe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   45
      Left            =   120
      TabIndex        =   91
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   43
      Left            =   3720
      TabIndex        =   90
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Tilgung in Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   600
      TabIndex        =   89
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   42
      Left            =   3600
      TabIndex        =   88
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   600
      TabIndex        =   87
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   41
      Left            =   3600
      TabIndex        =   86
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "- Gutscheinauszahlungen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   600
      TabIndex        =   85
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   40
      Left            =   9480
      TabIndex        =   84
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   39
      Left            =   9480
      TabIndex        =   83
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   38
      Left            =   9480
      TabIndex        =   82
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   37
      Left            =   9480
      TabIndex        =   81
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   36
      Left            =   9480
      TabIndex        =   80
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz ohne MWSt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   6600
      TabIndex        =   79
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Betrag erm MWSt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   6600
      TabIndex        =   78
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz erm. MWSt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   6600
      TabIndex        =   77
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Betrag volle MWSt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   6600
      TabIndex        =   76
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz volle MWSt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   6600
      TabIndex        =   75
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   35
      Left            =   3600
      TabIndex        =   74
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Lastschrift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   360
      TabIndex        =   73
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   34
      Left            =   3720
      TabIndex        =   72
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Lastschrift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   240
      TabIndex        =   71
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   33
      Left            =   3600
      TabIndex        =   70
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "daraus Rest-Gutscheine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   360
      TabIndex        =   69
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   32
      Left            =   3600
      TabIndex        =   68
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "eingereichte Gutscheine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   67
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kredit-Tilgungen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   6600
      TabIndex        =   66
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   27
      Left            =   9480
      TabIndex        =   65
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Tilgung in Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   6720
      TabIndex        =   64
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Tilgung per Scheck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   6720
      TabIndex        =   63
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Tilgung per Gutschein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   6720
      TabIndex        =   62
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Tilgung per Karte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   6720
      TabIndex        =   61
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   28
      Left            =   9600
      TabIndex        =   60
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   29
      Left            =   9600
      TabIndex        =   59
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   30
      Left            =   9600
      TabIndex        =   58
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   31
      Left            =   9600
      TabIndex        =   57
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   6120
      X2              =   11760
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   $"frmWKL21.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6120
      TabIndex        =   56
      Top             =   4680
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   26
      Left            =   3600
      TabIndex        =   55
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Schecks:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   240
      TabIndex        =   54
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   25
      Left            =   3600
      TabIndex        =   53
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Bar-Verkäufe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   600
      TabIndex        =   52
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   24
      Left            =   3600
      TabIndex        =   51
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   23
      Left            =   3600
      TabIndex        =   50
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   22
      Left            =   3600
      TabIndex        =   49
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   21
      Left            =   3600
      TabIndex        =   48
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Karte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   360
      TabIndex        =   47
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Kredite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   360
      TabIndex        =   46
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Scheck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   360
      TabIndex        =   45
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Gutschein über Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   360
      TabIndex        =   44
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   20
      Left            =   3720
      TabIndex        =   43
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Umsatz aus Gutscheinen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   240
      TabIndex        =   42
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6120
      X2              =   6120
      Y1              =   720
      Y2              =   4560
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   19
      Left            =   9480
      TabIndex        =   41
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   18
      Left            =   9480
      TabIndex        =   40
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   17
      Left            =   9480
      TabIndex        =   39
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   16
      Left            =   9480
      TabIndex        =   38
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   15
      Left            =   9480
      TabIndex        =   37
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   14
      Left            =   9480
      TabIndex        =   36
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   13
      Left            =   9480
      TabIndex        =   35
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   12
      Left            =   9480
      TabIndex        =   34
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   11
      Left            =   9480
      TabIndex        =   33
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   32
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   31
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   8
      Left            =   3840
      TabIndex        =   30
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   29
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   28
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   27
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   26
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   25
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   24
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   23
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   22
      Top             =   480
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Storno Anz.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   6600
      TabIndex        =   21
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Stornosumme:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   6600
      TabIndex        =   20
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Gesamtrabatt Anz.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   6600
      TabIndex        =   19
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Artikelrabatt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   6600
      TabIndex        =   17
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Schublade geöffnet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   6600
      TabIndex        =   16
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Sonderpreise Anz.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   6600
      TabIndex        =   15
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Sonderpreise:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   6600
      TabIndex        =   14
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kundenschnitt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   6600
      TabIndex        =   13
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kundenzahl:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   6600
      TabIndex        =   12
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Bargeld:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kassensoll:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "- Auszahlungen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Einzahlungen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Gutschein-Verkäufe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Kreditkarten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Kredite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Schecks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "+ Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz gesamt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Donnerstag,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   3732
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Tagesabschluß:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "0,00 DEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   53
      Left            =   9480
      TabIndex        =   123
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "frmWKL21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "FIRMA_21A", gdBase
    loeschNEW "FIRMA_21D", gdBase
    loeschNEW "FIRMA_21F", gdBase
    loeschNEW "afcd2", gdBase
    loeschNEW "afcd3", gdBase
    loeschNEW "afcd4", gdBase
    
    loeschNEW "TAKOLLKASS", gdBase
    loeschNEW "AFCDL", gdBase
    loeschNEW "afcnu", gdBase
    
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Positionieren()
    On Error GoTo LOKAL_ERROR
    
    With Frame1 'Tagesprotokoll
        .Top = 0
        .Left = 0
        .Height = 9000
        .Width = 12000
    End With
    
    With MSFlexGrid1 'Tagesprotokoll
        .Top = 120
        .Left = 120
        .Height = 7095
        .Width = 11655
    End With
    
    With Pbr1 'Progressbar
        .Top = 45
        .Left = 7800
        .Height = 255
        .Width = 3960
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DruckeKassenEinAuszahlungAufBonDruckerWKL21a()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cDatum As String
    Dim czeit As String
    Dim cBedNr As String
    Dim cART As String
    Dim dBetrag As Double
    Dim cBetrag As String
    Dim cBezeich As String
    
    Dim dWert As Double
    Dim dSummeEin As Double
    Dim dSummeAus As Double
    
    Dim lAnz As Long
    Dim lcount As Long
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim iLenZeile As Integer
    Dim cDaten As String
    Dim lAnzZeile As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim iRet As Integer
    
    Screen.MousePointer = 11
    
   
    cSQL = "Select * from KAEINAUS where Kasnum = " & gcKasNum & " order by ART, ADATE, AZEIT"
    Set rsrs = gdBase.OpenRecordset(cSQL)

    If Not rsrs.EOF Then
        '********************************************
        '*** 1.Schritt: Umschalten auf BonDrucker ***
        '********************************************
        setzedrucker gcBonDrucker
    
        '********************************************************
        '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
        '********************************************************
    
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
        OpenDrawer aDeviceName, cEscapeSequenz
    
        iLenZeile = 32
        'Drucker ist bereits auf BonDrucker geschaltet
        aDeviceName = gcBonDrucker
    
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "K.I.S.S. Warenwirtschaft"
        Else
            cDaten = gcBonText(0)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To 1) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "Carsten Schröder"
        Else
            cDaten = gcBonText(1)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "DEMO-VERSION!"
        Else
            cDaten = gcBonText(4)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '***********************************************
        'Kopfdaten 4.Zeile an Drucker senden
        '***********************************************
    
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "DEMO - VERSION"
        Else
            cDaten = gcBonText(12)
        End If
        
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '******************************************************************
        
        cDaten = "TAGESABSCHLUSS VOM " & Format$(Now, "DD.MM.YYYY")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        cDaten = "Kasse Ein- und Auszahlungen"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '******************************************************************
    
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        cDaten = "Kasse: " & gcKasNum
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '********************************************************
        '*** 3.Schritt: Daten drucken                         ***
        '********************************************************
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ADATE) Then
                dWert = rsrs!ADATE
            Else
                dWert = 0
            End If
            dWert = Fix(dWert)
            If dWert <> 0 Then
                cDatum = Format$(dWert, "DD.MM.YYYY")
            Else
                cDatum = Space$(10)
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                czeit = rsrs!AZEIT
            Else
                czeit = "00:00:00"
            End If
            
            If Not IsNull(rsrs!BEDNU) Then
                cBedNr = rsrs!BEDNU
            Else
                cBedNr = "000"
            End If
            cBedNr = Space$(3 - Len(cBedNr)) & cBedNr
            
            If Not IsNull(rsrs!Betrag) Then
                dBetrag = rsrs!Betrag
            Else
                dBetrag = 0
            End If
            cBetrag = Format$(dBetrag, "#####0.00")
            cBetrag = Space$(9 - Len(cBetrag)) & cBetrag
            cBetrag = cBetrag & " " & gcWaehrung
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            
            If Not IsNull(rsrs!art) Then
                cART = rsrs!art
            Else
                cART = ""
            End If
            
            If InStr(UCase$(cART), "EIN") > 0 Then
                dSummeEin = dSummeEin + dBetrag
            ElseIf InStr(UCase$(cART), "AUS") > 0 Then
                dSummeAus = dSummeAus + dBetrag
            End If
            
            '***** erste Zeile *****
            
            cDaten = cDatum & " " & czeit & " " & "Bed: " & cBedNr
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** zweite Zeile *****
            
            cDaten = cART & "  " & cBetrag
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** dritte Zeile *****
            
            cDaten = cBezeich
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** vierte Zeile *****
            
            cDaten = String$(32, "-")
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            rsrs.MoveNext
        Loop
        
        '***** Summenblock *****
        '***** erste Zeile *****
        
        cDaten = String$(32, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***** zweite Zeile *****
        cBetrag = Format$(dSummeAus, "#####0.00")
        cBetrag = Space$(9 - Len(cBetrag)) & cBetrag & " " & gcWaehrung
        cDaten = "Summe Ausz.: " & cBetrag
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***** Dritte Zeile *****
        cBetrag = Format$(dSummeEin, "#####0.00")
        cBetrag = Space$(9 - Len(cBetrag)) & cBetrag & " " & gcWaehrung
        cDaten = "Summe Einz.: " & cBetrag
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***** vierte Zeile *****
        
        cDaten = String$(32, "=")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '//Leerzeilen
        For lcount = 1 To 9
            If lcount = 9 Then
                cEscapeSequenz = "." & vbCrLf
            Else
                cEscapeSequenz = " " & vbCrLf
            End If
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
        '//end
        
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
        
        'Bon-Daten sichern
        
        Erase cDruckZeile
        
BON_SCHNEIDEN:
    
        'Kassenbon abschneiden
        If gbAPI Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = Chr$(27) + Chr$(105)
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
        
ENDE:

    '...und tschüß!

    '*******************************************************
    '*** Letzter Schritt: Umschalten auf ListenDrucker   ***
    '*******************************************************
    setzedrucker gcListenDrucker
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenEinAuszahlungAufBonDruckerWKL21a"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DruckeKassenAgnAuswertungaufBondrucker()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cAgn As String
    Dim cMenge As String
    Dim cAGNBEZEICH As String
    Dim dAnteil As Double
    Dim cAnteil As String
    Dim cPreis As String
    
    Dim dWert As Double
    Dim dPreis As Double
    Dim dSummeEin As Double
    Dim dSummeAus As Double
    
    
    Dim lAnz As Long
    Dim lcount As Long
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim iLenZeile As Integer
    Dim cDaten As String
    Dim lAnzZeile As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim iRet As Integer
    Dim dSum As Double
    
    Screen.MousePointer = 11
    
    loeschNEW "ABAGN", gdBase
    CreateTable "ABAGN", gdBase
    
    cSQL = "insert into ABAGN Select aartnr as artnr, apreis as preis, amenge as menge, 0 as agn from Afcbuch where Kasnum = " & gcKasNum
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ABAGN inner join Artikel on ABAGN.ARTNR = ARTIKEL.ARTNR set ABAGN.AGN = ARTIKEL.AGN"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "ABAGNS", gdBase
    CreateTable "ABAGNS", gdBase
    
    cSQL = "insert into ABAGNS Select agn,sum(preis) as apreis,sum(menge) as amenge from ABAGN group by agn "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ABAGNS inner join AGNDBF on ABAGNS.AGN = AGNDBF.AGN set ABAGNS.AGNBEZEICH = AGNDBF.AGTEXT"
    gdBase.Execute cSQL, dbFailOnError
    
    dSum = 0
    cSQL = "Select sum(apreis) as maxi from ABAGNS  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSum = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from ABAGNS order by AGN "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        '********************************************
        '*** 1.Schritt: Umschalten auf BonDrucker ***
        '********************************************
        setzedrucker gcBonDrucker
        
        '********************************************************
        '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
        '********************************************************
    
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
        OpenDrawer aDeviceName, cEscapeSequenz
    
    
        iLenZeile = 32
        'Drucker ist bereits auf BonDrucker geschaltet
        aDeviceName = gcBonDrucker
    
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "K.I.S.S. Warenwirtschaft"
        Else
            cDaten = gcBonText(0)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To 1) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "Carsten Schröder"
        Else
            cDaten = gcBonText(1)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "DEMO-VERSION!"
        Else
            cDaten = gcBonText(4)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '***********************************************
        'Kopfdaten 4.Zeile an Drucker senden
        '***********************************************
    
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "DEMO - VERSION"
        Else
            cDaten = gcBonText(12)
        End If
        
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '******************************************************************
        
        cDaten = "Datum " & Format$(Now, "DD.MM.YYYY") & " " & Format$(Now, "HH:MM:SS")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        cDaten = "Artikelgruppenauswertung"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '******************************************************************
    
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        cDaten = "Kasse: " & gcKasNum
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '********************************************************
        '*** 3.Schritt: Daten drucken                         ***
        '********************************************************

        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            Else
                cAgn = "0"
            End If
            
            If Not IsNull(rsrs!aMenge) Then
                cMenge = rsrs!aMenge
            Else
                cMenge = "0"
            End If
            
            cMenge = cMenge & " Stück"
            
            If Not IsNull(rsrs!AGNBEZEICH) Then
                cAGNBEZEICH = rsrs!AGNBEZEICH
            Else
                cAGNBEZEICH = ""
            End If
            
            


            dPreis = 0
        
            If Not IsNull(rsrs!APREIS) Then
                cPreis = rsrs!APREIS
                dPreis = rsrs!APREIS
            Else
                cPreis = "0"
            End If
            
            cPreis = Format$(cPreis, "#####0.00")
            
            dAnteil = Format((dPreis * 100) / dSum, "##0.00")
            cAnteil = CStr(dAnteil)
            cAnteil = cAnteil & " %"
            '***** erste Zeile *****
            
            cDaten = cAgn & Space$(6 - Len(cAgn)) & cAnteil & Space$(10 - Len(cAnteil)) & Space$(16 - Len(cMenge)) & cMenge
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** zweite Zeile *****
            
            cDaten = Left(cAGNBEZEICH, 22) & Space$(22 - Len(Left(cAGNBEZEICH, 22))) & Space$(10 - Len(cPreis)) & cPreis
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** dritte Zeile *****
            cDaten = String$(32, "-")
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            rsrs.MoveNext
        Loop
        '//Leerzeilen
        For lcount = 1 To 9
            If lcount = 9 Then
                cEscapeSequenz = "." & vbCrLf
            Else
                cEscapeSequenz = " " & vbCrLf
            End If
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
        '//end
        
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
        
        'Bon-Daten sichern
        SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False
        
        '//potong
        '//Aenderung
        'Kassenbon abschneiden
        If gbAPI = True Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcSchneiden
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
        '//End Aenderung
        
        '*******************************************************
        '*** Letzter Schritt: Umschalten auf ListenDrucker   ***
        '*** und ggf. Löschen der gedruckten Daten           ***
        '*******************************************************
        setzedrucker gcListenDrucker
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenAgnAuswertungaufBondrucker"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function DruckeKassenEinAuszahlungWKL21() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsZ As Recordset
    
    Dim cDatum As String
    Dim czeit As String
    Dim cBedNr As String
    Dim cBetrag As String
    Dim cBezeich As String
    Dim cART As String
    
    Dim cZielSatz As String
    Dim dWert As Double
    Dim dBetrag As Double
    Dim dSummeEin As Double
    Dim dSummeAus As Double
    
    Dim lCountEin As Long
    Dim lCountAus As Long
    
    lCountEin = 0
    lCountAus = 0
    
    DruckeKassenEinAuszahlungWKL21 = False
    
    loeschNEW "DRU_TEXT", gdBase
    
    cSQL = "Create Table DRU_TEXT"
    cSQL = cSQL & "(ZEILE Text(90)"
    cSQL = cSQL & ", Kasnum Text(2))"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from DRU_TEXT"
    Set rsZ = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select * from KAEINAUS where Kasnum = " & gcKasNum & " order by ART, ADATE, AZEIT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
   
    If Not rsrs.EOF Then
        DruckeKassenEinAuszahlungWKL21 = True
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ADATE) Then
                dWert = rsrs!ADATE
            Else
                dWert = 0
            End If
            dWert = Fix(dWert)
            If dWert <> 0 Then
                cDatum = Format$(dWert, "DD.MM.YYYY")
            Else
                cDatum = Space$(10)
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                czeit = rsrs!AZEIT
            Else
                czeit = "00:00:00"
            End If
            
            If Not IsNull(rsrs!BEDNU) Then
                cBedNr = rsrs!BEDNU
            Else
                cBedNr = "000"
            End If
            cBedNr = Space$(3 - Len(cBedNr)) & cBedNr
            
            If Not IsNull(rsrs!Betrag) Then
                dBetrag = rsrs!Betrag
            Else
                dBetrag = 0
            End If
            
            cBetrag = Format$(dBetrag, "#####0.00")
            cBetrag = Space$(9 - Len(cBetrag)) & cBetrag
            cBetrag = cBetrag & " " & gcWaehrung
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = Space$(32)
            End If
            
            If Not IsNull(rsrs!art) Then
                cART = rsrs!art
            Else
                cART = Space$(12)
            End If
            
            If InStr(UCase$(cART), "EIN") > 0 Then
                dSummeEin = dSummeEin + dBetrag
                lCountEin = lCountEin + 1
                
            ElseIf InStr(UCase$(cART), "AUS") > 0 Then
                dSummeAus = dSummeAus + dBetrag
                lCountAus = lCountAus + 1
            End If
            
            cZielSatz = cDatum & " " & czeit & " " & cBedNr & " " & cART & " " & cBetrag & " " & cBezeich
            
            rsZ.AddNew
            rsZ!ZEILE = cZielSatz
                rsZ!kasnum = gcKasNum
            rsZ.Update
            
            rsrs.MoveNext
        Loop
        
        cZielSatz = String$(82, "=")
        rsZ.AddNew
        rsZ!ZEILE = cZielSatz
        rsZ.Update
        
        cBetrag = Format$(dSummeAus, "#####0.00")
        cBetrag = Space$(9 - Len(cBetrag)) & cBetrag
        cBetrag = cBetrag & " " & gcWaehrung
                
        cZielSatz = "Summe Auszahlungen:               " & cBetrag & " (" & lCountAus & ")"
        rsZ.AddNew
        rsZ!ZEILE = cZielSatz
        rsZ.Update
        
        
        
        cBetrag = Format$(dSummeEin, "#####0.00")
        cBetrag = Space$(9 - Len(cBetrag)) & cBetrag
        cBetrag = cBetrag & " " & gcWaehrung
                
        cZielSatz = "Summe Einzahlungen:               " & cBetrag & " (" & lCountEin & ")"
        rsZ.AddNew
        rsZ!ZEILE = cZielSatz
        rsZ.Update
        
        
        
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsZ.Close: Set rsZ = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenEinAuszahlungWKL21"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub DruckeTagesAbschlussAufBonDruckerWKL21a()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount              As Long
    Dim lAnz                As Long
    Dim lAktSatz            As Long
    Dim lAnzSatz            As Long
    Dim dSumme              As Double
    Dim aDeviceName         As String
    Dim cEscapeSequenz      As String
    Dim cDaten              As String
    Dim iLenZeile           As Integer
    Dim lAnzZeile           As Long
    Dim cMeld               As String
    Dim ctmp                As String
    Dim cSQL                As String
    Dim rsrs                As Recordset
    ReDim cFeldName(0 To 0) As String
    ReDim iFeldPos(0 To 0) As Integer
    ReDim cFeldDruck(0 To 0) As String
    
    ReDim cDruckZeile(1 To 1) As String
    
    Dim cAlterAbschluß      As String
    Dim cNeuerAbschluß      As String
    
    Screen.MousePointer = 11

    cNeuerAbschluß = neuLbl.Caption
    cAlterAbschluß = altLbl.Caption
    
    

    
    '***********************************************************
    '*** Lese Konfiguration des Tagesprotokolls (LAYOUT.DBF) ***
    '***********************************************************
    Tabcheck "ZBON"
    
    cSQL = "Select * from ZBONLay where ANZEIGE = 'J' and Tabname = 'ZBON'"
    cSQL = cSQL & " order by Reihenf "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        lAnzSatz = rsrs.RecordCount
        
        ReDim cFeldName(1 To lAnzSatz) As String
        ReDim iFeldPos(1 To lAnzSatz) As Integer
        ReDim cFeldDruck(1 To lAnzSatz) As String
        
        rsrs.MoveFirst
        lAktSatz = 0
        Do While Not rsrs.EOF
            lAktSatz = lAktSatz + 1
            If Not IsNull(rsrs!Spaltenbez) Then
                cFeldName(lAktSatz) = rsrs!Spaltenbez
            Else
                cFeldName(lAktSatz) = ""
            End If
            If Not IsNull(rsrs!Reihenf) Then
                iFeldPos(lAktSatz) = rsrs!Reihenf
            Else
                iFeldPos(lAktSatz) = -1
            End If
            If Not IsNull(rsrs!anzeige) Then
                cFeldDruck(lAktSatz) = rsrs!anzeige
            Else
                cFeldDruck(lAktSatz) = "N"
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '********************************************
    '*** 1.Schritt: Umschalten auf BonDrucker ***
    '********************************************
    
    setzedrucker gcBonDrucker


    '********************************************************
    '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
    '********************************************************

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
    OpenDrawer aDeviceName, cEscapeSequenz


    iLenZeile = 32
    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker

    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To 1) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(1)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO-VERSION!"
    Else
        cDaten = gcBonText(4)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(12)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '*****************************************************************
    
    If cNeuerAbschluß = "neuerAbschluss" Then
        'hier geht jetzt alles anders
        
        
        cDaten = "Zwischenablesung: " & Format$(Now, "DD.MM.YY") & " " & Format$(Now, "HH:MM")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = "Kasse: " & gcKasNum
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '******************************************************************

    
        cDaten = "Kein gültiger Tagesabschluß"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        
        
    Else
    
        cDaten = "TAGESABSCHLUSS VOM " & Format$(Now, "DD.MM.YYYY")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = "Kasse: " & gcKasNum
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '******************************************************************

    
        cDaten = "jetziger Abschluß:"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        cDaten = Mid(cNeuerAbschluß, 22, 26)
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = " " & vbCrLf
        
        '******************************************************************
        
        cDaten = "vorheriger Abschluß:"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        cDaten = Mid(cAlterAbschluß, 22, 26)
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    End If
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    '******************************************************************************
    '***** An Bondrucker in gewünschter Reihenfolge und ohne ausgeblendete Zeilen *
    '******************************************************************************
    
    For lAktSatz = 1 To lAnzSatz
        If cFeldDruck(lAktSatz) = "J" Then
            Select Case cFeldName(lAktSatz)
                Case Is = "Umsatz gesamt:"
                    '******************************************************************
                    'Umsatz gesamt
                    '******************************************************************
                    cDaten = Trim$(Label3(0).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Umsatz gesamt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz gesamt:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Bar"
                    '******************************************************************
                    '+ Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(1).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Bar' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Bar:             " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Schecks"
                    '******************************************************************
                    '+ Schecks
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(2).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Schecks' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Schecks:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                Case Is = "+ Kredite"
                    '******************************************************************
                    '+ Kredite
                    
                    cDaten = Trim$(Label3(3).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Kredite' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Kredite:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Kreditkarte"
                    '******************************************************************
                    '+ Kreditkarten
                    
                    cDaten = Trim$(Label3(4).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Kreditkarte' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Kreditkarten:    " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Umsatz aus Gutscheinen:"
                    '******************************************************************
                    '+ Gutscheine
                    
                    cDaten = Trim$(Label3(20).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Umsatz aus Gutscheinen' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Umsatz Gutsch:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Lastschrift"
                    '******************************************************************
                    '+ Lastschrift
                    
                    cDaten = Trim$(Label3(34).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Lastschrift' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Lastschrift:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Dukaten"
                    '******************************************************************
                    '+ Dukaten
                    
                    cDaten = Trim$(Label3(50).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Dukaten' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Dukaten:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Kassensoll:"
                    '******************************************************************
                    'Kassensoll
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(8).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Kassensoll' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kassensoll:        " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Bargeld:"
                    '******************************************************************
                    '+ Bargeld
                    '******************************************************************
                    cDaten = Trim$(Label3(9).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Bargeld' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Bargeld:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Barverkäufe"
                    '******************************************************************
                    '+ Barverkäufe
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(25).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Barverkäufe' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Barverkäufe:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                Case Is = "+ Einzahlungen:"
                    '******************************************************************
                    '+ Einzahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(6).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Einzahlungen' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Einzahlungen:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ GutschVk Bar"
                    '******************************************************************
                    '+ Gutscheinverkäufe gegen Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(42).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ GutschVk Bar' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + GutschVk Bar:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Tilgung Bar"
                    '******************************************************************
                    '+ Tilgung gegen Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(43).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Tilgung Bar' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Tilgung Bar:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Wechselgeld:"
                    '******************************************************************
                    '+ Tilgung gegen Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(46).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Wechselgeld' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Wechselgeld:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Differenz summiert"
                    '******************************************************************
                    'Differenz summiert
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(49).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Differenz summiert' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    
                    cDaten = "Differenz summiert " & cDaten
                    
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Differenz jetzt:"
                    '******************************************************************
                    'Differenz jetzt:
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(48).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Differenz jetzt:' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    
                    cDaten = "Differenz jetzt:   " & cDaten
                
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    
                Case Is = "- Abschöpfung:"
                    '******************************************************************
                    '- Auszahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(47).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '- Abschöpfung' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  - Abschöpfung:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                    
                Case Is = "- Auszahlungen:"
                    '******************************************************************
                    '- Auszahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(7).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '- Auszahlungen' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  - Auszahlungen:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                Case Is = "- Gutscheinauszahlungen:"
                    '******************************************************************
                    '- Gutscheinauszahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(41).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '- Gutscheinauszahlungen' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  - GutschAuszahl: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                Case Is = "+ Schecks:"
                    '******************************************************************
                    '+ Schecks
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(26).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Schecks' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Schecks:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Kredit-Tilgungen"
                    '******************************************************************
                    'Tilgung
                    '******************************************************************
                    cDaten = Trim$(Label3(27).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Kredittilgungen' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kredittilgung:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Tilgung in Bar"
                    '******************************************************************
                    '+ Tilgung in Bar
                    '******************************************************************
                    cDaten = Trim$(Label3(28).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Tilgung in Bar' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ in Bar:          " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Tilgung per Scheck"
                    '******************************************************************
                    '+ Tilgung per Scheck
                    '******************************************************************
                    cDaten = Trim$(Label3(29).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Tilgung per Scheck' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ per Scheck:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Tilgung per Gutschein"
                    '******************************************************************
                    '+ Tilgung per Gutschein
                    '******************************************************************
                    cDaten = Trim$(Label3(30).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Tilgung per Gutschein' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ per Gutschein:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Tilgung per Karte"
                    '******************************************************************
                    '+ Tilgung per Kreditkarte
                    '******************************************************************
                    cDaten = Trim$(Label3(31).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Tilgung per Karte' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ per Kreditkarte: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "eingereichte Gutscheine:"
                    '******************************************************************
                    'eingelöste Gutsch
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(32).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'eingereichte Gutscheine' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "eingelöste Gutsch. " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "daraus Fremdgutscheine:"
                    '******************************************************************
                    'Restgutscheine
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(54).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'daraus Fremdgutscheine' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "daraus Fremdgutsch." & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "daraus Rest-Gutscheine:"
                    '******************************************************************
                    'Restgutscheine
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(33).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'daraus Rest-Gutscheine' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "daraus RestGutsch. " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Kundenzahl:"
                    '******************************************************************
                    'Kundenzahl
                    '******************************************************************
                    cDaten = Trim$(Label3(10).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Kundenzahl' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kundenzahl:        " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Kundenschnitt:"
                    '******************************************************************
                    'Kundenschitt
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(11).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Kundenschnitt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kundenschnitt:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Sonderpreise:"
                    '******************************************************************
                    'Sonderpreise
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(12).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Sonderpreise' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Sonderpreise:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Sonderpreise Anz.:"
                    '******************************************************************
                    'Sonderpreise Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(13).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Sonderpreise Anz.' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Sonderpreise Anz.: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Schublade geöffnet:"
                    '******************************************************************
                    'Schublade offen
                    '******************************************************************
                    cDaten = Trim$(Label3(14).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Schublade geöffnet' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Schublade offen:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Artikelrabatt:"
                    '******************************************************************
                    'Artikelrabatt
                    '******************************************************************
                    cDaten = Trim$(Label3(15).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Artikelrabatt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Artikelrabatt:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Artikelrabatt Anz.:"
                    '******************************************************************
                    'Gesamtrabatt Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(53).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Artikelrabatt Anz.' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Artikelrabatt Anz.:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Gesamtrabatt:"
                    '******************************************************************
                    'Gesamtrabatt
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(16).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Gesamtrabatt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gesamtrabatt:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Gesamtrabatt Anz.:"
                    '******************************************************************
                    'Gesamtrabatt Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(17).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Gesamtrabatt Anz.' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gesamtrabatt Anz.: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Stornosumme:"
                    '******************************************************************
                    'Stornosumme
                    '******************************************************************
                    cDaten = Trim$(Label3(18).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Stornosumme' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Stornosumme:       " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Storno Anz.:"
                    '******************************************************************
                    'Storno Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(19).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Storno Anz.' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Storno Anz.:       " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                
                
                Case Is = "Gutschein-Verkäufe:"
                    '******************************************************************
                    'Gutschein-Verkauf Gesamt
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(5).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Gutschein-Verkäufe' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein-Verkauf: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Bar"
                    '******************************************************************
                    'Gutschein-Verkauf BAR
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(21).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Gutscheine über Bar' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Bar:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Scheck"
                    '******************************************************************
                    'Gutschein-Verkauf SCHECK
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(22).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Gutscheine über Scheck' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Scheck:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Kredite"
                    '******************************************************************
                    'Gutschein-Verkauf KREDIT
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(23).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Gutscheine über Kredit' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Kredit:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Karte"
                    '******************************************************************
                    'Gutschein-Verkauf KARTE
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(24).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Gutscheine über Karte' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Karte:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Lastschrift"
                    '******************************************************************
                    'Gutschein-Verkauf LASTSCHRIFT
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(35).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei '+ Gutscheine über Lastschrift' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Lastschr:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Gutscheine über Gutscheine"
                    '******************************************************************
                    'Gutschein-Verkauf Gutscheine
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(45).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Gutschei:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Umsatz volle MWSt:"
                    '******************************************************************
                    'Umsatz volle MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(36).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Umsatz volle MWSt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz volle MWSt: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Betrag volle MWSt:"
                    '******************************************************************
                    'Betrag volle MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(37).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Betrag volle MWSt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Betrag volle MWSt: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Umsatz erm. MWSt:"
                    '******************************************************************
                    'Umsatz erm. MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(38).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Umsatz erm. MWSt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz erm. MWSt:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Betrag erm. MWSt:"
                    '******************************************************************
                    'Betrag erm. MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(39).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Betrag erm. MWSt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Betrag erm. MWSt:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Umsatz ohne MWSt:"
                    '******************************************************************
                    'Umsatz ohne MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(40).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Umsatz ohne MWSt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz ohne MWSt:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Gesamtumsatz:"
                    
                    '******************************************************************
                    'Gesamtumsatz(E,V,O)
                    '******************************************************************
                    
                    dSumme = CDbl(Left(Label3(40).Caption, Len(Label3(40).Caption) - Len(gcWaehrung) - 1))
                    dSumme = dSumme + CDbl(Left(Label3(38).Caption, Len(Label3(38).Caption) - Len(gcWaehrung) - 1))
                    dSumme = dSumme + CDbl(Left(Label3(36).Caption, Len(Label3(36).Caption) - Len(gcWaehrung) - 1))
                    
                    cDaten = Format$(dSumme, "####,##0.00") & " " & gcWaehrung
                    cDaten = Space$(19 - Len(cDaten)) & cDaten
                    
                    cDaten = "Gesamtumsatz:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                
                Case Is = "Kreditkarte gesamt"
                    '******************************************************************
                    'Kreditkarte gesamt
                    '******************************************************************
                    cDaten = Trim$(Label3(52).Caption)
                    If Len(cDaten) > 13 Then
                        cMeld = "Zu großer Wert bei 'Kreditkarte gesamt' (mehr als 13 Stellen)!" & vbCrLf
                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
                        MsgBox cMeld, vbCritical, "STOP!"
                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kreditkarte gesamt " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Else
                    MsgBox cFeldName(lAktSatz)
            End Select
        End If
    Next lAktSatz
        
    '******************************************************************
    'nicht umsatzrelevante Verkäufe
    '******************************************************************
    cDaten = Trim$(Label3(44).Caption)
    If Len(cDaten) > 13 Then
        cMeld = "Zu großer Wert bei 'nicht umsatzrelevante Verkäufe' (mehr als 13 Stellen)!" & vbCrLf
        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
        MsgBox cMeld, vbCritical, "STOP!"
    End If
    cDaten = Space$(13 - Len(cDaten)) & cDaten
    cDaten = "nicht umsatzrelv.: " & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    
    '******************************************************************
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = Format$(Now, "DD.MM.YYYY") & "                 " & Format$(Now, "HH:MM")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If

    
    '******************************************************************
    
    For lcount = 1 To 9
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    
    'Bon-Daten sichern
    
    Erase cDruckZeile
    
BON_SCHNEIDEN:

    'Kassenbon abschneiden
    If gbAPI Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(105)
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
ENDE:

    '...und tschüß!

    '*******************************************************
    '*** Letzter Schritt: Umschalten auf ListenDrucker   ***
    '*******************************************************
    setzedrucker gcListenDrucker
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeTagesAbschlussAufBonDrucker"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeTagesabschlussNeuWKL21a(iAuswahl As Integer, gbzweit As Boolean)
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim lcount As Long
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim cTmp2 As String
    Dim cDaten As String
    Dim dKey As Double
    Dim cPfad As String
    Dim dbMdb As Database
    Dim cWaeCode As String

    Dim cAlterAbschluß As String
    Dim cNeuerAbschluß As String


    cNeuerAbschluß = neuLbl.Caption
    cAlterAbschluß = altLbl.Caption
    
    If cNeuerAbschluß = "neuerAbschluss" Then
        cNeuerAbschluß = "Dies ist kein gültiger Tagesabschluss"
        cAlterAbschluß = "Zwischenablesung " & Format$(Fix(Now), "DD.MM.YY") & " " & Format$(Now, "HH:MM:SS")
    End If



    '//new
    If gcWaehrung = "EUR" Then
        cWaeCode = "alle Preise in EURO"
    Else
        cWaeCode = "alle Preise in " & gcWaehrung
    End If


    '****************************************
    '* dieser Teil wird immer durchlaufen!  *
    '****************************************
    
    

    loeschNEW "TAGKOPF", gdBase

    cSQL = "Create Table TAGKOPF "
    cSQL = cSQL & "(SCHLUESSEL double"
    cSQL = cSQL & ", WAE_CODE TEXT(20)"
    cSQL = cSQL & ", DATEN1 TEXT(50)"
    cSQL = cSQL & ", DATEN2 TEXT(50)"
    cSQL = cSQL & ", DATEN3 TEXT(50)"
    cSQL = cSQL & ", DATEN4 TEXT(50)"
    cSQL = cSQL & ", DATEN5 TEXT(50)"
    cSQL = cSQL & ", DATEN6 TEXT(50)"
    cSQL = cSQL & ", DATEN7 TEXT(50)"
    cSQL = cSQL & ", DATEN8 TEXT(50)"
    cSQL = cSQL & ", DATEN9 TEXT(50)"
    cSQL = cSQL & ", DATEN10 TEXT(50)"
    cSQL = cSQL & ", DATEN11 TEXT(50)"
    cSQL = cSQL & ", DATEN12 TEXT(50)"
    cSQL = cSQL & ", DATEN13 TEXT(50)"
    cSQL = cSQL & ", DATEN14 TEXT(50)"
    cSQL = cSQL & ", DATEN15 TEXT(50)"
    cSQL = cSQL & ", DATEN16 TEXT(50)"
    cSQL = cSQL & ", DATEN17 TEXT(50)"
    cSQL = cSQL & ", DATEN18 TEXT(50)"
    cSQL = cSQL & ", DATEN19 TEXT(50)"
    cSQL = cSQL & ", DATEN20 TEXT(50)"
    cSQL = cSQL & ", DATEN21 TEXT(50)"
    cSQL = cSQL & ", DATEN22 TEXT(50)"
    cSQL = cSQL & ", DATEN23 TEXT(50)"
    cSQL = cSQL & ", DATEN24 TEXT(50)"
    cSQL = cSQL & ", DATEN25 TEXT(50)"
    cSQL = cSQL & ", DATEN26 TEXT(50)"
    cSQL = cSQL & ", DATEN27 TEXT(50)"
    cSQL = cSQL & ", DATEN28 TEXT(50)"
    cSQL = cSQL & ", DATEN29 TEXT(50)"
    cSQL = cSQL & ", DATEN30 TEXT(50)"
    cSQL = cSQL & ", DATEN31 TEXT(50)"
    cSQL = cSQL & ", DATEN32 TEXT(50)"
    cSQL = cSQL & ", DATEN33 TEXT(50)"
    cSQL = cSQL & ", DATEN34 TEXT(50)"
    cSQL = cSQL & ", DATEN35 TEXT(50)"
    cSQL = cSQL & ", DATEN36 TEXT(50)"
    cSQL = cSQL & ", DATEN37 TEXT(50)"
    cSQL = cSQL & ", DATEN38 TEXT(50)"
    cSQL = cSQL & ", DATEN39 TEXT(50)"
    cSQL = cSQL & ", DATEN40 TEXT(50)"
    cSQL = cSQL & ", DATEN41 TEXT(50)"
    cSQL = cSQL & ", DATEN42 TEXT(50)"
    cSQL = cSQL & ", DATEN43 TEXT(50)"
    cSQL = cSQL & ", DATEN44 TEXT(50)"
    cSQL = cSQL & ", DATEN45 TEXT(50)"
    cSQL = cSQL & ", DATEN46 TEXT(50)"
    cSQL = cSQL & ", DATEN47 TEXT(50)"
    cSQL = cSQL & ", DATEN48 TEXT(50)"
    cSQL = cSQL & ", DATEN49 TEXT(50)"
    cSQL = cSQL & ", DATEN50 TEXT(50)"
    cSQL = cSQL & ", DATEN51 TEXT(50)"
    cSQL = cSQL & ", DATEN52 TEXT(50)"
    cSQL = cSQL & ", DATEN53 TEXT(50)"
    cSQL = cSQL & ", DATEN54 TEXT(50)"
    cSQL = cSQL & ", DATEN TEXT(150)"
    cSQL = cSQL & ", ALTERABSCH TEXT(50)"
    cSQL = cSQL & ", NEUERABSCH TEXT(50)"
    cSQL = cSQL & ", DATEN55 TEXT(50)"
    cSQL = cSQL & ")"

    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Drop Index SCHLUESSEL on TAGKOPF"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Create Index SCHLUESSEL on TAGKOPF (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError

    dKey = Now
    dKey = Fix(dKey * 1000)

    '****************************************
    '* dieser Teil produziert Kopfdaten!    *
    '****************************************

    If iAuswahl = 1 Or iAuswahl = 3 Then

        cSQL = "Select * from TAGKOPF"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If rsrs.EOF Then
            rsrs.AddNew
            rsrs!schluessel = dKey

            rsrs!WAE_CODE = cWaeCode

            '**************************************
            ' linke Seite der Kopfdaten
            '**************************************

            'Umsatz gesamt
            cTmp2 = Label3(0).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN1 = cTmp2

            '+ Bar
            cTmp2 = Label3(1).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN2 = cTmp2

            '+ Schecks
            cTmp2 = Label3(2).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN3 = cTmp2

            '+ Kredite
            cTmp2 = Label3(3).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN4 = cTmp2

            '+ Kreditkarten
            cTmp2 = Label3(4).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN5 = cTmp2

            '+ Gutscheine
            cTmp2 = Label3(20).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN6 = cTmp2

            '+ Lastschrift
            cTmp2 = Label3(34).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN7 = cTmp2
            
            '+ Dukaten
            cTmp2 = Label3(50).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN52 = cTmp2

            'Kassensoll
            
            
            If Label3(51).Caption <> "" Then
            
                rsrs!DATEN8 = Trim$(Label3(51).Caption) & " "
            
            End If
            
            cTmp2 = Label3(8).Caption
            cTmp2 = Trim$(cTmp2)
            
            rsrs!DATEN8 = rsrs!DATEN8 & cTmp2

            '+ Bargeld
            cTmp2 = Label3(9).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN9 = cTmp2

            '+ Barverkäufe
            cTmp2 = Label3(25).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN10 = cTmp2

            '+ Einzahlungen
            cTmp2 = Label3(6).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN11 = cTmp2
            
            '+ Wechselgeld
            cTmp2 = Label3(46).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN49 = cTmp2

            '- Abschöfpung
            cTmp2 = Label3(47).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN48 = cTmp2


            '- Auszahlungen
            cTmp2 = Label3(7).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN12 = cTmp2

            '- Gutschein Auszahlungen
            cTmp2 = Label3(41).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN13 = cTmp2

            '+ Schecks
            cTmp2 = Label3(26).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN14 = cTmp2

            'Kredittilgungen
            cTmp2 = Label3(27).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN15 = cTmp2

            '+ Tilgung Bar
            cTmp2 = Label3(28).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN16 = cTmp2

            '+ Tilgung Scheck
            cTmp2 = Label3(29).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN17 = cTmp2

            '+ Tilgung Gutschein
            cTmp2 = Label3(30).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN18 = cTmp2

            '+ Tilgung Karte
            cTmp2 = Label3(31).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN19 = cTmp2

            'Eingelöste Gutscheine
            cTmp2 = Label3(32).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN20 = cTmp2

            'Daraus Restgutscheine
            cTmp2 = Label3(33).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN21 = cTmp2

            'nicht umsatzrelevante Verkäufe
            cTmp2 = Label3(44).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN45 = cTmp2

            '**************************************
            ' rechte Seite der Kopfdaten
            '**************************************

            'Kundenzahl
            cTmp2 = Label3(10).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN22 = cTmp2

            'Kundenschnitt
            cTmp2 = Label3(11).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN23 = cTmp2

            'Sonderpreise
            cTmp2 = Label3(12).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN24 = cTmp2

            'Sonderpreise Anzahl
            cTmp2 = Label3(13).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN25 = cTmp2

            'Schublade geöffnet
            cTmp2 = Label3(14).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN26 = cTmp2

            'Artikelrabatt

            cTmp2 = Label3(15).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN27 = cTmp2
            
            'artikelrabatt Anz
            cTmp2 = Label3(53).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN54 = cTmp2

            'Gesamtrabatt
            cTmp2 = Label3(16).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN28 = cTmp2

            'Gesamtrabatt Anz
            cTmp2 = Label3(17).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN29 = cTmp2

            'Stornosumme
            cTmp2 = Label3(18).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN30 = cTmp2

            'Stornosumme Anz
            cTmp2 = Label3(19).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN31 = cTmp2

            'Gutschein Verkauf
            cTmp2 = Label3(5).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN32 = cTmp2

            'Gutschein in Bar

            cTmp2 = Label3(21).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN33 = cTmp2

            'Gutschein per Scheck
            cTmp2 = Label3(22).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN34 = cTmp2

            'Gutschein per Kredit
            cTmp2 = Label3(23).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN35 = cTmp2

            'Gutschein per Karte
            cTmp2 = Label3(24).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN36 = cTmp2

            'Gutschein per Lastschrift
            cTmp2 = Label3(35).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN37 = cTmp2
            
            'Gutschein per Gutschein
            cTmp2 = Label3(45).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN47 = cTmp2

            'Umsatz volle MWSt
            cTmp2 = Label3(36).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN38 = cTmp2
            'Betrag volle MWSt

            cTmp2 = Label3(37).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN39 = cTmp2

            'Umsatz erm. MWSt
            cTmp2 = Label3(38).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN40 = cTmp2

            'Betrag erm. MWSt
            cTmp2 = Label3(39).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN41 = cTmp2

            'Umsatz ohne MWSt
            cTmp2 = Label3(40).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN42 = cTmp2

            '+ GutschVk Bar
            cTmp2 = Label3(42).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN43 = cTmp2

            '+ Tilgung Bar
            cTmp2 = Label3(43).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN44 = cTmp2
            
            'differenz jetzt
            cTmp2 = Label3(48).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN50 = cTmp2
            
            'differenz kumu
            cTmp2 = Label3(49).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN51 = cTmp2
            
            'karte insgesamt
            cTmp2 = Label3(52).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN53 = cTmp2
            
            'Kreditkarte,nichtumsrele vk
            cTmp2 = Label3(55).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN55 = cTmp2

            'Kassennummer
            rsrs!DATEN46 = gcKasNum

            'Abschlußdaten
            rsrs!ALTERABSCH = cAlterAbschluß
            rsrs!NEUERABSCH = cNeuerAbschluß

            rsrs.Update
        End If
        rsrs.Close: Set rsrs = Nothing

    End If
    



Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 3372 Or err.Number = 3051 Or err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeTagesabschlussNeuWKL21a"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."

        Fehlermeldung1
    End If
End Sub
Public Sub GDPdU_ZBON_sichern_Teil1(cAlterZbon As String, cNeuerZBon As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim i       As Integer
    Dim rsrs    As DAO.Recordset
    Dim dKey    As Double
    Dim cTmp2   As String
    
    dKey = Now
    dKey = Fix(dKey * 1000)

    loeschNEW "TAGKOPF_" & srechnertab, gdBase
    
    cSQL = "Create Table TAGKOPF_" & srechnertab & " "
    cSQL = cSQL & "(SCHLUESSEL double"
    cSQL = cSQL & ", WAE_CODE TEXT(20)"
    For i = 1 To 66
        cSQL = cSQL & ", DATEN" & i & " TEXT(54)"
    Next i
    cSQL = cSQL & ", DATEN TEXT(150)"
    cSQL = cSQL & ", ALTERABSCH TEXT(50)"
    cSQL = cSQL & ", NEUERABSCH TEXT(50)"
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from TAGKOPF_" & srechnertab & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.AddNew
        rsrs!schluessel = dKey

        rsrs!WAE_CODE = "alle Preise in EURO"

        '**************************************
        ' linke Seite der Kopfdaten
        '**************************************

        'Umsatz gesamt
        cTmp2 = Label3(0).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN1 = cTmp2

        '+ Bar
        cTmp2 = Label3(1).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN2 = cTmp2

        '+ Schecks
        cTmp2 = Label3(2).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN3 = cTmp2

        '+ Kredite
        cTmp2 = Label3(3).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN4 = cTmp2

        '+ Kreditkarten
        cTmp2 = Label3(4).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN5 = cTmp2

        '+ Gutscheine
        cTmp2 = Label3(20).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN6 = cTmp2

        '+ Lastschrift
        cTmp2 = Label3(34).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN7 = cTmp2
        
        '+ Dukaten
        cTmp2 = Label3(50).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN52 = cTmp2

        'Kassensoll
        cTmp2 = Label3(8).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN8 = cTmp2

        '+ Bargeld
        cTmp2 = Label3(9).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN9 = cTmp2

        '+ Barverkäufe
        cTmp2 = Label3(25).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN10 = cTmp2

        '+ Einzahlungen
        cTmp2 = Label3(6).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN11 = cTmp2
        
        '+ Wechselgeld
        cTmp2 = Label3(46).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN49 = cTmp2

        '- Abschöfpung
        cTmp2 = Label3(47).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN48 = cTmp2


        '- Auszahlungen
        cTmp2 = Label3(7).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN12 = cTmp2

        '- Gutschein Auszahlungen
        cTmp2 = Label3(41).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN13 = cTmp2

        '+ Schecks
        cTmp2 = Label3(26).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN14 = cTmp2

        'Kredittilgungen
        cTmp2 = Label3(27).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN15 = cTmp2

        '+ Tilgung Bar
        cTmp2 = Label3(28).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN16 = cTmp2

        '+ Tilgung Scheck
        cTmp2 = Label3(29).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN17 = cTmp2

        '+ Tilgung Gutschein
        cTmp2 = Label3(30).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN18 = cTmp2

        '+ Tilgung Karte
        cTmp2 = Label3(31).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN19 = cTmp2

        'Eingelöste Gutscheine
        cTmp2 = Label3(32).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN20 = cTmp2

        'Daraus Restgutscheine
        cTmp2 = Label3(33).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN21 = cTmp2

        'nicht umsatzrelevante Verkäufe
        cTmp2 = Label3(44).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN45 = cTmp2

        '**************************************
        ' rechte Seite der Kopfdaten
        '**************************************

        'Kundenzahl
        cTmp2 = Label3(10).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN22 = cTmp2

        'Kundenschnitt
        cTmp2 = Label3(11).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN23 = cTmp2

        'Sonderpreise
        cTmp2 = Label3(12).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN24 = cTmp2

        'Sonderpreise Anzahl
        cTmp2 = Label3(13).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN25 = cTmp2

        'Schublade geöffnet
        cTmp2 = Label3(14).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN26 = cTmp2

        'Artikelrabatt

        cTmp2 = Label3(15).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN27 = cTmp2
        
        'artikelrabatt Anz
        cTmp2 = Label3(53).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN54 = cTmp2

        'Gesamtrabatt
        cTmp2 = Label3(16).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN28 = cTmp2

        'Gesamtrabatt Anz
        cTmp2 = Label3(17).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN29 = cTmp2

        'Stornosumme
        cTmp2 = Label3(18).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN30 = cTmp2

        'Stornosumme Anz
        cTmp2 = Label3(19).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN31 = cTmp2

        'Gutschein Verkauf
        cTmp2 = Label3(5).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN32 = cTmp2

        'Gutschein in Bar
        cTmp2 = Label3(21).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN33 = cTmp2

        'Gutschein per Scheck
        cTmp2 = Label3(22).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN34 = cTmp2

        'Gutschein per Kredit
        cTmp2 = Label3(23).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN35 = cTmp2

        'Gutschein per Karte
        cTmp2 = Label3(24).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN36 = cTmp2

        'Gutschein per Lastschrift
        cTmp2 = Label3(35).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN37 = cTmp2
        
        'Gutschein per Gutschein
        cTmp2 = Label3(45).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN47 = cTmp2

        'Umsatz volle MWSt
        cTmp2 = Label3(36).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN38 = cTmp2
        'Betrag volle MWSt

        cTmp2 = Label3(37).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN39 = cTmp2

        'Umsatz erm. MWSt
        cTmp2 = Label3(38).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN40 = cTmp2

        'Betrag erm. MWSt
        cTmp2 = Label3(39).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN41 = cTmp2

        'Umsatz ohne MWSt
        cTmp2 = Label3(40).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN42 = cTmp2

        '+ GutschVk Bar
        cTmp2 = Label3(42).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN43 = cTmp2

        '+ Tilgung Bar
        cTmp2 = Label3(43).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN44 = cTmp2
        
        'differenz jetzt
        cTmp2 = Label3(48).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN50 = cTmp2
        
        'differenz kumu
        cTmp2 = Label3(49).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN51 = cTmp2
        
        'karte insgesamt
        cTmp2 = Label3(52).Caption
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN53 = cTmp2

        'Kassennummer
        rsrs!DATEN46 = gcKasNum

        'Abschlußdaten
        rsrs!ALTERABSCH = cAlterZbon
        rsrs!NEUERABSCH = cNeuerZBon

        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GDPdU_ZBON_sichern_Teil1"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FormatiereMSFlexGrid1DetailWKL21()
    On Error GoTo LOKAL_ERROR
    
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Cols = 13
    
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = "Datum"
    MSFlexGrid1.ColWidth(0) = 1000
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = "Uhrzeit"
    MSFlexGrid1.ColWidth(1) = 800
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Text = "ArtNr."
    MSFlexGrid1.ColWidth(2) = 800
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Text = "Artikelbezeichnung"
    MSFlexGrid1.ColWidth(3) = 3500
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Text = "VK-Menge"
    MSFlexGrid1.ColWidth(4) = 1000
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Text = "Preis"
    MSFlexGrid1.ColWidth(5) = 800
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = "VK-Wert"
    MSFlexGrid1.ColWidth(6) = 1000
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Text = "Listenpreis"
    MSFlexGrid1.ColWidth(7) = 1000
    
    MSFlexGrid1.Col = 8
    MSFlexGrid1.Text = "Bestand"
    MSFlexGrid1.ColWidth(8) = 1000
    
    MSFlexGrid1.Col = 9
    MSFlexGrid1.Text = "KK"  '** soll KK sein **
    MSFlexGrid1.ColWidth(9) = 500
    
    MSFlexGrid1.Col = 10
    MSFlexGrid1.Text = "Bed"
    MSFlexGrid1.ColWidth(10) = 500
    
    MSFlexGrid1.Col = 11
    MSFlexGrid1.Text = "Bon"
    MSFlexGrid1.ColWidth(11) = 800
    
    MSFlexGrid1.Col = 12
    MSFlexGrid1.Text = "ZhlgGutsch"
    MSFlexGrid1.ColWidth(12) = 1000

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereMSFlexGrid1DetailWKL21"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseDatenWKL21()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs3 As Recordset
    Dim dWert As Double
    Dim dSumme As Double
    Dim dUmsatz As Double
    Dim cMWSK As String
    Dim dEinzahlung As Double
    Dim dAuszahlung As Double
    Dim dAuszGutsch As Double
    Dim dBar As Double
    Dim dKunden As Double
    Dim dScheck As Double
    Dim dKredit As Double
    Dim dKarte As Double
    Dim dLast As Double
    Dim dUmsBar As Double
    Dim dUmsScheck As Double
    
    Dim dZhlgGutsch As Double
    
    Dim dKasse As Double
    Dim dKassenBargeld As Double
    Dim dKassenSchecks As Double
    Dim dSchVerkauf As Double
    Dim dBarVerkauf As Double
    
    Dim dGutschein As Double
    Dim dGutschBar As Double
    Dim dGutschSch As Double
    Dim dGutschKre As Double
    Dim dGutschKar As Double
    Dim dGutschLast As Double
    Dim dGutschGUTSCH As Double
    Dim dABSCHOPF As Double
    Dim dKDIFF As Double
    Dim dTDIFF As Double
    Dim dDUKA As Double
    Dim dWECHSEL As Double
    Dim dEinrGutsch As Double
    
    Dim dTilgung As Double
    Dim dTilgBar As Double
    Dim dTilgSch As Double
    Dim dTilgGut As Double
    Dim dTilgKar As Double
    
    Dim dNichtUmsReleKar As Double
    
    'check
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "GUTSCHGUTSCH", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "GUTSCHGUTSCH", "double", gdBase
    
        cSQL = "Update AFCSTAT set GUTSCHGUTSCH = 0 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "ABSCHOPF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "ABSCHOPF", "double", gdBase
    
        cSQL = "Update AFCSTAT set ABSCHOPF = 0 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "WECHSEL", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "WECHSEL", "double", gdBase
    
        cSQL = "Update AFCSTAT set WECHSEL = 0 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "KDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "KDIFF", "double", gdBase
    
        cSQL = "Update AFCSTAT set KDIFF = 0 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "TDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "TDIFF", "double", gdBase
    
        cSQL = "Update AFCSTAT set TDIFF = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "DUKA", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "DUKA", "double", gdBase
    
        cSQL = "Update AFCSTAT set DUKA = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "NUMSKARTE", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "NUMSKARTE", "double", gdBase
    
        cSQL = "Update AFCSTAT set NUMSKARTE = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "KDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "KDIFF", "double", gdBase
    
        cSQL = "Update AFCSTATP set KDIFF = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "TDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "TDIFF", "double", gdBase
    
        cSQL = "Update AFCSTATP set TDIFF = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "DUKA", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "DUKA", "double", gdBase
    
        cSQL = "Update AFCSTATP set DUKA = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "NUMSKARTE", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "NUMSKARTE", "double", gdBase
    
        cSQL = "Update AFCSTATP set NUMSKARTE = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    
    
    'check ende
    
    
'    If NewTableSuchenDBKombi("AFCSTATP", gdBase) Then
'        cSQL = "Select KASNUM"
'        cSQL = cSQL & ", SUM(KDIFF) as SKDIFF, max(adate) "
'        cSQL = cSQL & " from AFCSTATP where KASNUM = " & gcKasNum & " "
'        cSQL = cSQL & " group by KASNUM "
'        Set rsRs3 = gdBase.OpenRecordset(cSQL)
'
'        If Not rsRs3.EOF Then
'
'            rsRs3.MoveFirst
'
'            If Not IsNull(rsRs3!SKDIFF) Then
'                dWert = rsRs3!SKDIFF
'            Else
'                dWert = 0
'            End If
'            dKDIFF = dWert
'
'        End If
'        rsRs3.Close: Set rsRs3 = Nothing
'    Else
        dKDIFF = 0
'    End If
    
    cSQL = "Select KASNUM"
    cSQL = cSQL & ", SUM(UMS_BAR) as SUMS_BAR"
    cSQL = cSQL & ", SUM(UMS_KRED) as SUMS_KRED"
    cSQL = cSQL & ", SUM(UMS_KARTE) as SUMS_KARTE"
    cSQL = cSQL & ", SUM(UMS_SCHECK) as SUMS_SCHEC"
    cSQL = cSQL & ", SUM(UMS_LAST) as SUMS_LAST"
    
    cSQL = cSQL & ", SUM(TILGBAR) as STILGBAR"
    cSQL = cSQL & ", SUM(TILGSCH) as STILGSCH"
    cSQL = cSQL & ", SUM(TILGGUT) as STILGGUT"
    cSQL = cSQL & ", SUM(TILGKAR) as STILGKAR"
    
    cSQL = cSQL & ", SUM(GUTSCHBAR) as SGUTSCHBAR"
    cSQL = cSQL & ", SUM(GUTSCHSCH) as SGUTSCHSCH"
    cSQL = cSQL & ", SUM(GUTSCHKRE) as SGUTSCHKRE"
    cSQL = cSQL & ", SUM(GUTSCHKAR) as SGUTSCHKAR"
    cSQL = cSQL & ", SUM(GUTSCHLAST) as SGUTSCHLAS"
    cSQL = cSQL & ", SUM(GUTSCHGUTSCH) as SGUTSCHGUTSCH"
    cSQL = cSQL & ", SUM(ABSCHOPF) as SABSCHOPF"
    cSQL = cSQL & ", SUM(KDIFF) as SKDIFF"
    cSQL = cSQL & ", SUM(TDIFF) as STDIFF"
    cSQL = cSQL & ", SUM(DUKA) as SDUKA"
    cSQL = cSQL & ", SUM(WECHSEL) as SWECHSEL"
    
    cSQL = cSQL & ", SUM(BARVERKAUF) as SBARVERKAU"
    cSQL = cSQL & ", SUM(SCHVERKAUF) as SSCHVERKAU"
    
    cSQL = cSQL & ", SUM(AUSZAHLUNG) as SAUSZAHLUN"
    cSQL = cSQL & ", SUM(EINZAHLUNG) as SEINZAHLUN"
    cSQL = cSQL & ", SUM(AUSZGUTSCH) as SAUSZGUTSC"
    
    cSQL = cSQL & ", SUM(SPREIS_GES) as SSPREIS_GE"
    cSQL = cSQL & ", SUM(SPREIS_ANZ) as SSPREIS_AN"
    cSQL = cSQL & ", SUM(GESRAB_GES) as SGESRAB_GE"
    cSQL = cSQL & ", SUM(GESRAB_ANZ) as SGESRAB_AN"
    cSQL = cSQL & ", SUM(ARTRAB_GES) as SARTRAB_GE"
    cSQL = cSQL & ", SUM(ARTRAB_ANZ) as SARTRAB_AN"
    cSQL = cSQL & ", SUM(STORNO_GES) as SSTORNO_GE"
    cSQL = cSQL & ", SUM(STORNO_ANZ) as SSTORNO_AN"
    
    cSQL = cSQL & ", SUM(ZHLGGUTSCH) as SZHLGGUTSC"
    cSQL = cSQL & ", SUM(KUNDENZAHL) as SKUNDENZAH"
    cSQL = cSQL & ", SUM(GELDFACH) as SGELDFACH"
    
    cSQL = cSQL & ", SUM(EINRGUTSCH) as SEINRGUTSC"
    cSQL = cSQL & ", SUM(RESTGUTSCH) as SRESTGUTSC"
    cSQL = cSQL & ", SUM(GUTSCHEIN) as SGUTSCH"
    cSQL = cSQL & ", SUM(NUMSKARTE) as SNUMSKARTE"
    cSQL = cSQL & " from AFCSTAT where KASNUM = " & gcKasNum & " group by KASNUM "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        dWert = rsrs.RecordCount
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SUMS_BAR) Then
            dWert = rsrs!SUMS_BAR
        Else
            dWert = 0
        End If
        dUmsBar = dWert
        
        If Not IsNull(rsrs!SUMS_SCHEC) Then
            dWert = rsrs!SUMS_SCHEC
        Else
            dWert = 0
        End If
        dUmsScheck = dWert
        
        If Not IsNull(rsrs!SUMS_KARTE) Then
            dWert = rsrs!SUMS_KARTE
        Else
            dWert = 0
        End If
        dKarte = dWert
                
        If Not IsNull(rsrs!SUMS_KRED) Then
            dWert = rsrs!SUMS_KRED
        Else
            dWert = 0
        End If
        dKredit = dWert
        
        
        If Not IsNull(rsrs!SUMS_LAST) Then
            dWert = rsrs!SUMS_LAST
        Else
            dWert = 0
        End If
        dLast = dWert
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        
        If Not IsNull(rsrs!STILGBAR) Then
            dWert = rsrs!STILGBAR
        Else
            dWert = 0
        End If
        dTilgBar = dWert
        
        If Not IsNull(rsrs!STILGSCH) Then
            dWert = rsrs!STILGSCH
        Else
            dWert = 0
        End If
        dTilgSch = dWert
        
        If Not IsNull(rsrs!STILGGUT) Then
            dWert = rsrs!STILGGUT
        Else
            dWert = 0
        End If
        dTilgGut = dWert
        
        If Not IsNull(rsrs!STILGKAR) Then
            dWert = rsrs!STILGKAR
        Else
            dWert = 0
        End If
        dTilgKar = dWert
        
        dTilgung = dTilgBar + dTilgSch + dTilgGut + dTilgKar
        
        If Not IsNull(rsrs!SGUTSCHBAR) Then
            dWert = rsrs!SGUTSCHBAR
        Else
            dWert = 0
        End If
        dGutschBar = dWert
        
        If Not IsNull(rsrs!SGUTSCHSCH) Then
            dWert = rsrs!SGUTSCHSCH
        Else
            dWert = 0
        End If
        dGutschSch = dWert
        
        If Not IsNull(rsrs!SGUTSCHKRE) Then
            dWert = rsrs!SGUTSCHKRE
        Else
            dWert = 0
        End If
        dGutschKre = dWert
        
        If Not IsNull(rsrs!SGUTSCHKAR) Then
            dWert = rsrs!SGUTSCHKAR
        Else
            dWert = 0
        End If
        dGutschKar = dWert
        
        
        If Not IsNull(rsrs!SNUMSKARTE) Then
            dWert = rsrs!SNUMSKARTE
        Else
            dWert = 0
        End If
        dNichtUmsReleKar = dWert
        
        
        
        
        
        If Not IsNull(rsrs!SGUTSCHLAS) Then
            dWert = rsrs!SGUTSCHLAS
        Else
            dWert = 0
        End If
        dGutschLast = dWert
        
        If Not IsNull(rsrs!SGUTSCHGUTSCH) Then
            dWert = rsrs!SGUTSCHGUTSCH
        Else
            dWert = 0
        End If
        dGutschGUTSCH = dWert
        
        If Not IsNull(rsrs!SABSCHOPF) Then
            dWert = rsrs!SABSCHOPF
        Else
            dWert = 0
        End If
        dABSCHOPF = dWert
        

        dTDIFF = 0
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        If Not IsNull(rsrs!SWECHSEL) Then
            dWert = rsrs!SWECHSEL
        Else
            dWert = 0
        End If
        dWECHSEL = dWert
        
        If Not IsNull(rsrs!sGutsch) Then
            dWert = rsrs!sGutsch
        Else
            dWert = 0
        End If
        dGutschein = dWert
        
        If Not IsNull(rsrs!SSCHVERKAU) Then
            dWert = rsrs!SSCHVERKAU
        Else
            dWert = 0
        End If
        dSchVerkauf = dWert
        
'        dKassenSchecks = dSchVerkauf + dGutschSch + dTilgSch
        dKassenSchecks = dSchVerkauf + dTilgSch
        
        If Not IsNull(rsrs!SAUSZAHLUN) Then
            dWert = rsrs!SAUSZAHLUN
        Else
            dWert = 0
        End If
        dAuszahlung = dWert
        
        If Not IsNull(rsrs!SEINZAHLUN) Then
            dWert = rsrs!SEINZAHLUN
        Else
            dWert = 0
        End If
        dEinzahlung = dWert
        
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        dAuszGutsch = dWert
        
        If Not IsNull(rsrs!SBARVERKAU) Then
            dWert = rsrs!SBARVERKAU
        Else
            dWert = 0
        End If
        dBarVerkauf = dWert
        
         dKassenBargeld = dBarVerkauf + dGutschBar + dTilgBar + dEinzahlung - dAuszahlung - dAuszGutsch - dABSCHOPF + dWECHSEL
          
        'Odayy Änderung
        'dKassenBargeld = dBarVerkauf + dTilgBar + dEinzahlung - dAuszahlung - dAuszGutsch - dABSCHOPF + dWECHSEL
        'Odayy Änderung
        
         
        
        If gbBargeldEingabe = True Then
            dTDIFF = gdKassenGeldGezählt - dKassenBargeld
            gdKassenGeldGezählt = 0
            
            dKDIFF = dKDIFF + dTDIFF
        End If
        
        dKasse = dKassenBargeld + dKassenSchecks
        
        If Not IsNull(rsrs!SZHLGGUTSC) Then
            dWert = rsrs!SZHLGGUTSC
        Else
            dWert = 0
        End If
        dZhlgGutsch = dWert
                
        dScheck = dKassenSchecks - dGutschSch - dTilgSch
        dBar = dBarVerkauf
            
        '//gefunden
        dUmsatz = dZhlgGutsch + dKarte + dKredit + dUmsScheck + dUmsBar + dLast + dDUKA
        
        
           
        ctmp = Format$(dUmsatz, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(0).Caption = ctmp
           
        ctmp = Format$(dUmsBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(1).Caption = ctmp
        
        ctmp = Format$(dUmsScheck, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(2).Caption = ctmp
        
        ctmp = Format$(dKredit, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(3).Caption = ctmp
        
        ctmp = Format$(dKarte, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(4).Caption = ctmp
        
        ctmp = Format$(dZhlgGutsch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(20).Caption = ctmp
        
        ctmp = Format$(dKasse, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(8).Caption = ctmp
        
        ctmp = Format$(dKassenBargeld, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(9).Caption = ctmp
        
        ctmp = Format$(dBarVerkauf, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(25).Caption = ctmp
        
        ctmp = Format$(dEinzahlung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(6).Caption = ctmp
        
        ctmp = Format$(dAuszahlung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(7).Caption = ctmp
        
        ctmp = Format$(dABSCHOPF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(47).Caption = ctmp
        
        If dABSCHOPF > 0 Then
            ctmp = Format$(dABSCHOPF + dKasse, "###,###,##0.00")
            Label3(51).Caption = "(" & ctmp & ")"
        Else
            Label3(51).Caption = ""
        End If
        
        
        ctmp = Format$(dKDIFF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(49).Caption = ctmp
        
        
        
        ctmp = Format$(dTDIFF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(48).Caption = ctmp
        
        ctmp = Format$(dDUKA, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(50).Caption = ctmp
        
        ctmp = Format$(dWECHSEL, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(46).Caption = ctmp
        
        ctmp = Format$(dKassenSchecks, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(26).Caption = ctmp
        
        ctmp = Format$(dGutschein, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(5).Caption = ctmp
        
        ctmp = Format$(dGutschBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(21).Caption = ctmp
        Label3(42).Caption = ctmp

        
        ctmp = Format$(dGutschSch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(22).Caption = ctmp
        
        ctmp = Format$(dGutschKre, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(23).Caption = ctmp
        
        ctmp = Format$(dGutschKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(24).Caption = ctmp
        
        ctmp = Format$(dBarVerkauf, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(25).Caption = ctmp
        
        ctmp = Format$(dKassenSchecks, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(26).Caption = ctmp
        
        ctmp = Format$(dTilgung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(27).Caption = ctmp
        
        ctmp = Format$(dTilgBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(28).Caption = ctmp
        Label3(43).Caption = ctmp
        
        ctmp = Format$(dTilgSch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(29).Caption = ctmp
        
        ctmp = Format$(dTilgGut, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(30).Caption = ctmp
        
        ctmp = Format$(dTilgKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(31).Caption = ctmp
        
        ctmp = Format$(dLast, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(34).Caption = ctmp
        
        ctmp = Format$(dGutschLast, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(35).Caption = ctmp
        
        ctmp = Format$(dGutschGUTSCH, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(45).Caption = ctmp
        
        If Not IsNull(rsrs!SKUNDENZAH) Then
            dWert = rsrs!SKUNDENZAH
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(10).Caption = ctmp
        
        If dKunden = 0 Then
            dKunden = 1
        End If
        
        dWert = dUmsatz / dKunden
        ctmp = Format$(dWert, "###,###,##0.00")
        Label3(11).Caption = ctmp & " " & gcWaehrung
        
        If Not IsNull(rsrs!SSPREIS_GE) Then
            dWert = rsrs!SSPREIS_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(12).Caption = ctmp
    
        If Not IsNull(rsrs!SSPREIS_AN) Then
            dWert = rsrs!SSPREIS_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(13).Caption = ctmp
    
        If Not IsNull(rsrs!SGELDFACH) Then
            dWert = rsrs!SGELDFACH
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(14).Caption = ctmp
    
        If Not IsNull(rsrs!SARTRAB_GE) Then
            dWert = rsrs!SARTRAB_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(15).Caption = ctmp
    
        If Not IsNull(rsrs!SGESRAB_GE) Then
            dWert = rsrs!SGESRAB_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(16).Caption = ctmp
    
        If Not IsNull(rsrs!SGESRAB_AN) Then
            dWert = rsrs!SGESRAB_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(17).Caption = ctmp
        
        If Not IsNull(rsrs!SARTRAB_AN) Then
            dWert = rsrs!SARTRAB_AN
        Else
            dWert = 0
        End If
        
        ctmp = Format$(dWert, "###,###,##0")
        Label3(53).Caption = ctmp
    
        If Not IsNull(rsrs!SSTORNO_GE) Then
            dWert = rsrs!SSTORNO_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(18).Caption = ctmp
    
        If Not IsNull(rsrs!SSTORNO_AN) Then
            dWert = rsrs!SSTORNO_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(19).Caption = ctmp
        
        If Not IsNull(rsrs!SEINRGUTSC) Then
            dWert = rsrs!SEINRGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(32).Caption = ctmp
    
        If Not IsNull(rsrs!SRESTGUTSC) Then
            dWert = rsrs!SRESTGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(33).Caption = ctmp
    
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(41).Caption = ctmp
        
        
        
        ctmp = Format$(dNichtUmsReleKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(55).Caption = ctmp
        
       
        
        If gbGutscheinBeiVKversteuern = True Then
            ctmp = Format$(dKarte + dTilgKar + dNichtUmsReleKar, "###,###,##0.00")
        Else
            ctmp = Format$(dKarte + dGutschKar + dTilgKar + dNichtUmsReleKar, "###,###,##0.00")
        End If
        
        
        ctmp = ctmp & " " & gcWaehrung
        Label3(52).Caption = ctmp
        
        ctmp = Format$(ermSumAlterg, "###,###,##0.00 ") & gcWaehrung
        Label3(54).Caption = ctmp
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gbBargeldEingabe = True Then
        updateafcstat "TDIFF", dTDIFF, gcKasNum
        updateafcstat "KDIFF", dTDIFF, gcKasNum
    End If
    
    Label3(36).Caption = "0,00 " & gcWaehrung
    Label3(37).Caption = "0,00 " & gcWaehrung
    Label3(38).Caption = "0,00 " & gcWaehrung
    Label3(39).Caption = "0,00 " & gcWaehrung
    Label3(40).Caption = "0,00 " & gcWaehrung
    
    
    
    
    Dim dNichtUmsGutschbetrag As Double
    dNichtUmsGutschbetrag = 0
    
    If gbGutscheinBeiVKversteuern = True Then
    
        
        cSQL = "Select SUM(Wert) as UMSATZ from Gemischte_Z where kasnum = " & gcKasNum
        cSQL = cSQL & " and Thema = 'nicht ums GUTSCHBETRAG'"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!UMSATZ) Then
                dNichtUmsGutschbetrag = rsrs!UMSATZ
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
    
        cSQL = "Select KASNUM, AMWSK, SUM(APREIS) as UMSATZ from AFCBUCH "
        cSQL = cSQL & "where KASNUM = " & gcKasNum & " and UMS_OK <> 'N' group by KASNUM, AMWSK "
    Else
    
        cSQL = "Select KASNUM, AMWSK, SUM(APREIS) as UMSATZ from AFCBUCH "
        cSQL = cSQL & "where KASNUM = " & gcKasNum & " and AARTNR <> 666666 and UMS_OK <> 'N' group by KASNUM, AMWSK "
    
    End If
    
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!AMWSK) Then
                cMWSK = rsrs!AMWSK
            Else
                    
            End If
            
            
            
            If Not IsNull(rsrs!UMSATZ) Then
                dUmsatz = rsrs!UMSATZ
            Else
                dUmsatz = 0
            End If
            
            Select Case cMWSK
                Case Is = "V"
                    
                    dUmsatz = dUmsatz - dNichtUmsGutschbetrag
                    
                    Label3(36).Caption = Format$(dUmsatz, "######0.00") & " " & gcWaehrung
                    Label3(37).Caption = Format$((dUmsatz / (gdMWStV + 100)) * gdMWStV, "######0.00") & " " & gcWaehrung
                    
                Case Is = "E"
                    Label3(38).Caption = Format$(dUmsatz, "######0.00") & " " & gcWaehrung
                    Label3(39).Caption = Format$((dUmsatz / (gdMWStE + 100)) * gdMWStE, "######0.00") & " " & gcWaehrung
                Case Is = "O"
                    Label3(40).Caption = Format$(dUmsatz, "######0.00") & " " & gcWaehrung
                Case Else
                
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    
    
    
    
    
    
    loeschNEW "AfcTempo", gdBase
    
    
    If gbGutscheinBeiVKversteuern = True Then
        cSQL = "Select * into AFCTempo from AFCBUCH where UMS_OK = 'N' "
        gdBase.Execute cSQL, dbFailOnError
    Else
        cSQL = "Select * into AFCTempo from AFCBUCH where UMS_OK = 'N' or AARTNR = 666666"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    dWert = 0
    cSQL = "Select SUM(APREIS) as UMSATZ from AFCTempo where kasnum = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!UMSATZ) Then
            dWert = rsrs!UMSATZ
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "AfcTempo", gdBase
    
    
    dWert = dWert + dNichtUmsGutschbetrag
    
    
    
    Label3(44).Caption = Format$(dWert, "######0.00") & " " & gcWaehrung
    
Exit Sub
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leseDatenWKL21"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub leseDatenWKL21Lokal()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim dWert As Double
    Dim dSumme As Double
    Dim dUmsatz As Double
    Dim cMWSK As String
    Dim dEinzahlung As Double
    Dim dAuszahlung As Double
    Dim dAuszGutsch As Double
    Dim dBar As Double
    Dim dKunden As Double
    Dim dScheck As Double
    Dim dKredit As Double
    Dim dKarte As Double
    Dim dLast As Double
    Dim dUmsBar As Double
    Dim dUmsScheck As Double
    
    Dim dZhlgGutsch As Double
    
    Dim dKasse As Double
    Dim dKassenBargeld As Double
    Dim dKassenSchecks As Double
    Dim dSchVerkauf As Double
    Dim dBarVerkauf As Double
    
    Dim dGutschein As Double
    Dim dGutschBar As Double
    Dim dGutschSch As Double
    Dim dGutschKre As Double
    Dim dGutschKar As Double
    Dim dGutschLast As Double
    Dim dGutschGUTSCH As Double
    Dim dABSCHOPF As Double
    Dim dDUKA As Double
    Dim dKDIFF As Double
    Dim dTDIFF As Double
    Dim dWECHSEL As Double
    Dim dEinrGutsch As Double
    
    Dim dTilgung As Double
    Dim dTilgBar As Double
    Dim dTilgSch As Double
    Dim dTilgGut As Double
    Dim dTilgKar As Double
    
    Dim dNichtUmsReleKar As Double
    
    'check
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "GUTSCHGUTSCH", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "GUTSCHGUTSCH", "double", gdBase
    
        cSQL = "Update AFCSTAT set GUTSCHGUTSCH = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "ABSCHOPF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "ABSCHOPF", "double", gdBase
    
        cSQL = "Update AFCSTAT set ABSCHOPF = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "WECHSEL", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "WECHSEL", "double", gdBase
    
        cSQL = "Update AFCSTAT set WECHSEL = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "KDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "KDIFF", "double", gdBase
    
        cSQL = "Update AFCSTAT set KDIFF = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "TDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "TDIFF", "double", gdBase
    
        cSQL = "Update AFCSTAT set TDIFF = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "DUKA", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "DUKA", "double", gdBase
    
        cSQL = "Update AFCSTAT set DUKA = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "NUMSKARTE", gdBase) Then
        SpalteAnfuegenNEW "AFCSTAT", "NUMSKARTE", "double", gdBase
    
        cSQL = "Update AFCSTAT set NUMSKARTE = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "KDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "KDIFF", "double", gdBase
    
        cSQL = "Update AFCSTATP set KDIFF = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "TDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "TDIFF", "double", gdBase
    
        cSQL = "Update AFCSTATP set TDIFF = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "DUKA", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "DUKA", "double", gdBase
    
        cSQL = "Update AFCSTATP set DUKA = 0 "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "NUMSKARTE", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "NUMSKARTE", "double", gdBase
    
        cSQL = "Update AFCSTATP set NUMSKARTE = 0 "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    'check ende
    
    cSQL = "Select KASNUM"
    cSQL = cSQL & ", SUM(UMS_BAR) as SUMS_BAR"
    cSQL = cSQL & ", SUM(UMS_KRED) as SUMS_KRED"
    cSQL = cSQL & ", SUM(UMS_KARTE) as SUMS_KARTE"
    cSQL = cSQL & ", SUM(UMS_SCHECK) as SUMS_SCHEC"
    cSQL = cSQL & ", SUM(UMS_LAST) as SUMS_LAST"
    
    cSQL = cSQL & ", SUM(TILGBAR) as STILGBAR"
    cSQL = cSQL & ", SUM(TILGSCH) as STILGSCH"
    cSQL = cSQL & ", SUM(TILGGUT) as STILGGUT"
    cSQL = cSQL & ", SUM(TILGKAR) as STILGKAR"
    
    cSQL = cSQL & ", SUM(GUTSCHBAR) as SGUTSCHBAR"
    cSQL = cSQL & ", SUM(GUTSCHSCH) as SGUTSCHSCH"
    cSQL = cSQL & ", SUM(GUTSCHKRE) as SGUTSCHKRE"
    cSQL = cSQL & ", SUM(GUTSCHKAR) as SGUTSCHKAR"
    cSQL = cSQL & ", SUM(GUTSCHLAST) as SGUTSCHLAS"
    cSQL = cSQL & ", SUM(GUTSCHGUTSCH) as SGUTSCHGUTSCH"
    
    cSQL = cSQL & ", SUM(ABSCHOPF) as SABSCHOPF"
    cSQL = cSQL & ", SUM(KDIFF) as SKDIFF"
    cSQL = cSQL & ", SUM(TDIFF) as STDIFF"
    cSQL = cSQL & ", SUM(DUKA) as SDUKA"
    cSQL = cSQL & ", SUM(WECHSEL) as SWECHSEL"
    
    cSQL = cSQL & ", SUM(BARVERKAUF) as SBARVERKAU"
    cSQL = cSQL & ", SUM(SCHVERKAUF) as SSCHVERKAU"
    
    cSQL = cSQL & ", SUM(AUSZAHLUNG) as SAUSZAHLUN"
    cSQL = cSQL & ", SUM(EINZAHLUNG) as SEINZAHLUN"
    cSQL = cSQL & ", SUM(AUSZGUTSCH) as SAUSZGUTSC"
    
    cSQL = cSQL & ", SUM(SPREIS_GES) as SSPREIS_GE"
    cSQL = cSQL & ", SUM(SPREIS_ANZ) as SSPREIS_AN"
    cSQL = cSQL & ", SUM(GESRAB_GES) as SGESRAB_GE"
    cSQL = cSQL & ", SUM(GESRAB_ANZ) as SGESRAB_AN"
    cSQL = cSQL & ", SUM(ARTRAB_GES) as SARTRAB_GE"
    cSQL = cSQL & ", SUM(ARTRAB_ANZ) as SARTRAB_AN"
    cSQL = cSQL & ", SUM(STORNO_GES) as SSTORNO_GE"
    cSQL = cSQL & ", SUM(STORNO_ANZ) as SSTORNO_AN"
    
    cSQL = cSQL & ", SUM(ZHLGGUTSCH) as SZHLGGUTSC"
    cSQL = cSQL & ", SUM(KUNDENZAHL) as SKUNDENZAH"
    cSQL = cSQL & ", SUM(GELDFACH) as SGELDFACH"
    
    cSQL = cSQL & ", SUM(EINRGUTSCH) as SEINRGUTSC"
    cSQL = cSQL & ", SUM(RESTGUTSCH) as SRESTGUTSC"
    cSQL = cSQL & ", SUM(GUTSCHEIN) as SGUTSCH"
    cSQL = cSQL & ", SUM(NUMSKARTE) as SNUMSKARTE"
    
    cSQL = cSQL & " from AFCSTAT where KASNUM = " & gcKasNum & " group by KASNUM "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        dWert = rsrs.RecordCount
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SUMS_BAR) Then
            dWert = rsrs!SUMS_BAR
        Else
            dWert = 0
        End If
        dUmsBar = dWert
        
        If Not IsNull(rsrs!SUMS_SCHEC) Then
            dWert = rsrs!SUMS_SCHEC
        Else
            dWert = 0
        End If
        dUmsScheck = dWert
        
        If Not IsNull(rsrs!SUMS_KARTE) Then
            dWert = rsrs!SUMS_KARTE
        Else
            dWert = 0
        End If
        dKarte = dWert
                
        If Not IsNull(rsrs!SUMS_KRED) Then
            dWert = rsrs!SUMS_KRED
        Else
            dWert = 0
        End If
        dKredit = dWert
        
        
        If Not IsNull(rsrs!SUMS_LAST) Then
            dWert = rsrs!SUMS_LAST
        Else
            dWert = 0
        End If
        dLast = dWert
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        
        If Not IsNull(rsrs!STILGBAR) Then
            dWert = rsrs!STILGBAR
        Else
            dWert = 0
        End If
        dTilgBar = dWert
        
        If Not IsNull(rsrs!STILGSCH) Then
            dWert = rsrs!STILGSCH
        Else
            dWert = 0
        End If
        dTilgSch = dWert
        
        If Not IsNull(rsrs!STILGGUT) Then
            dWert = rsrs!STILGGUT
        Else
            dWert = 0
        End If
        dTilgGut = dWert
        
        If Not IsNull(rsrs!STILGKAR) Then
            dWert = rsrs!STILGKAR
        Else
            dWert = 0
        End If
        dTilgKar = dWert
        
        dTilgung = dTilgBar + dTilgSch + dTilgGut + dTilgKar
        
        If Not IsNull(rsrs!SGUTSCHBAR) Then
            dWert = rsrs!SGUTSCHBAR
        Else
            dWert = 0
        End If
        dGutschBar = dWert
        
        If Not IsNull(rsrs!SGUTSCHSCH) Then
            dWert = rsrs!SGUTSCHSCH
        Else
            dWert = 0
        End If
        dGutschSch = dWert
        
        If Not IsNull(rsrs!SGUTSCHKRE) Then
            dWert = rsrs!SGUTSCHKRE
        Else
            dWert = 0
        End If
        dGutschKre = dWert
        
        If Not IsNull(rsrs!SGUTSCHKAR) Then
            dWert = rsrs!SGUTSCHKAR
        Else
            dWert = 0
        End If
        dGutschKar = dWert
        
        If Not IsNull(rsrs!SNUMSKARTE) Then
            dWert = rsrs!SNUMSKARTE
        Else
            dWert = 0
        End If
        dNichtUmsReleKar = dWert
        
        If Not IsNull(rsrs!SGUTSCHLAS) Then
            dWert = rsrs!SGUTSCHLAS
        Else
            dWert = 0
        End If
        dGutschLast = dWert
        
        If Not IsNull(rsrs!SGUTSCHGUTSCH) Then
            dWert = rsrs!SGUTSCHGUTSCH
        Else
            dWert = 0
        End If
        dGutschGUTSCH = dWert
        
        If Not IsNull(rsrs!sGutsch) Then
            dWert = rsrs!sGutsch
        Else
            dWert = 0
        End If
        dGutschein = dWert
        
        If Not IsNull(rsrs!SSCHVERKAU) Then
            dWert = rsrs!SSCHVERKAU
        Else
            dWert = 0
        End If
        dSchVerkauf = dWert
        
'        dKassenSchecks = dSchVerkauf + dGutschSch + dTilgSch
        dKassenSchecks = dSchVerkauf + dTilgSch
        
        If Not IsNull(rsrs!SAUSZAHLUN) Then
            dWert = rsrs!SAUSZAHLUN
        Else
            dWert = 0
        End If
        dAuszahlung = dWert
        
        If Not IsNull(rsrs!SEINZAHLUN) Then
            dWert = rsrs!SEINZAHLUN
        Else
            dWert = 0
        End If
        dEinzahlung = dWert
        
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        dAuszGutsch = dWert
        
        If Not IsNull(rsrs!SBARVERKAU) Then
            dWert = rsrs!SBARVERKAU
        Else
            dWert = 0
        End If
        dBarVerkauf = dWert
        
        If Not IsNull(rsrs!SABSCHOPF) Then
            dWert = rsrs!SABSCHOPF
        Else
            dWert = 0
        End If
        dABSCHOPF = dWert
        

        dTDIFF = 0
        
        If Not IsNull(rsrs!SKDIFF) Then
            dWert = rsrs!SKDIFF
        Else
            dWert = 0
        End If
        dKDIFF = dWert
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        If Not IsNull(rsrs!SWECHSEL) Then
            dWert = rsrs!SWECHSEL
        Else
            dWert = 0
        End If
        dWECHSEL = dWert
        
        'Odayy Änderung
        'dKassenBargeld = dBarVerkauf + dTilgBar + dEinzahlung - dAuszahlung - dAuszGutsch - dABSCHOPF + dWECHSEL
        'Odayy Änderung
        
        dKassenBargeld = dBarVerkauf + dGutschBar + dTilgBar + dEinzahlung - dAuszahlung - dAuszGutsch - dABSCHOPF + dWECHSEL
        
        If gbBargeldEingabe = True Then
            dTDIFF = gdKassenGeldGezählt - dKassenBargeld
            gdKassenGeldGezählt = 0
        End If
        
        
        dKasse = dKassenBargeld + dKassenSchecks
        
        If Not IsNull(rsrs!SZHLGGUTSC) Then
            dWert = rsrs!SZHLGGUTSC
        Else
            dWert = 0
        End If
        dZhlgGutsch = dWert
                
        dScheck = dKassenSchecks - dGutschSch - dTilgSch
        dBar = dBarVerkauf
            
        '//gefunden
        dUmsatz = dZhlgGutsch + dKarte + dKredit + dUmsScheck + dUmsBar + dLast + dDUKA
        
           
        ctmp = Format$(dUmsatz, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(0).Caption = ctmp
           
        ctmp = Format$(dUmsBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(1).Caption = ctmp
        
        ctmp = Format$(dUmsScheck, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(2).Caption = ctmp
        
        ctmp = Format$(dKredit, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(3).Caption = ctmp
        
        ctmp = Format$(dKarte, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(4).Caption = ctmp
        
        ctmp = Format$(dZhlgGutsch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(20).Caption = ctmp
        
        ctmp = Format$(dKasse, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(8).Caption = ctmp
        
        ctmp = Format$(dKassenBargeld, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(9).Caption = ctmp
        
        ctmp = Format$(dBarVerkauf, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(25).Caption = ctmp
        
        ctmp = Format$(dEinzahlung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(6).Caption = ctmp
        
        ctmp = Format$(dAuszahlung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(7).Caption = ctmp
        
        ctmp = Format$(dABSCHOPF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(47).Caption = ctmp
        
        ctmp = Format$(dKDIFF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(49).Caption = ctmp
        
        ctmp = Format$(dDUKA, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(50).Caption = ctmp
        
        ctmp = Format$(dTDIFF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(48).Caption = ctmp
        
        ctmp = Format$(dWECHSEL, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(46).Caption = ctmp
        
        ctmp = Format$(dKassenSchecks, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(26).Caption = ctmp
        
        ctmp = Format$(dGutschein, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(5).Caption = ctmp
        
        ctmp = Format$(dGutschBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(21).Caption = ctmp
        Label3(42).Caption = ctmp

        ctmp = Format$(dGutschSch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(22).Caption = ctmp
        
        ctmp = Format$(dGutschKre, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(23).Caption = ctmp
        
        ctmp = Format$(dGutschKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(24).Caption = ctmp
        
        ctmp = Format$(dBarVerkauf, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(25).Caption = ctmp
        
        ctmp = Format$(dKassenSchecks, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(26).Caption = ctmp
        
        ctmp = Format$(dTilgung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(27).Caption = ctmp
        
        ctmp = Format$(dTilgBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(28).Caption = ctmp
        Label3(43).Caption = ctmp
        
        ctmp = Format$(dTilgSch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(29).Caption = ctmp
        
        ctmp = Format$(dTilgGut, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(30).Caption = ctmp
        
        ctmp = Format$(dTilgKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(31).Caption = ctmp
        
        ctmp = Format$(dLast, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(34).Caption = ctmp
        
        ctmp = Format$(dGutschLast, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(35).Caption = ctmp
        
        ctmp = Format$(dGutschGUTSCH, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(45).Caption = ctmp
        
        If Not IsNull(rsrs!SKUNDENZAH) Then
            dWert = rsrs!SKUNDENZAH
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(10).Caption = ctmp
        
        If dKunden = 0 Then
            dKunden = 1
        End If
        
        dWert = dUmsatz / dKunden
        ctmp = Format$(dWert, "###,###,##0.00")
        Label3(11).Caption = ctmp & " " & gcWaehrung
        
        If Not IsNull(rsrs!SSPREIS_GE) Then
            dWert = rsrs!SSPREIS_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(12).Caption = ctmp
    
        If Not IsNull(rsrs!SSPREIS_AN) Then
            dWert = rsrs!SSPREIS_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(13).Caption = ctmp
    
        If Not IsNull(rsrs!SGELDFACH) Then
            dWert = rsrs!SGELDFACH
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(14).Caption = ctmp
    
        If Not IsNull(rsrs!SARTRAB_GE) Then
            dWert = rsrs!SARTRAB_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(15).Caption = ctmp
    
        If Not IsNull(rsrs!SGESRAB_GE) Then
            dWert = rsrs!SGESRAB_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(16).Caption = ctmp
    
        If Not IsNull(rsrs!SGESRAB_AN) Then
            dWert = rsrs!SGESRAB_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(17).Caption = ctmp
        
        If Not IsNull(rsrs!SARTRAB_AN) Then
            dWert = rsrs!SARTRAB_AN
        Else
            dWert = 0
        End If
        
        ctmp = Format$(dWert, "###,###,##0")
        Label3(53).Caption = ctmp
    
        If Not IsNull(rsrs!SSTORNO_GE) Then
            dWert = rsrs!SSTORNO_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(18).Caption = ctmp
    
        If Not IsNull(rsrs!SSTORNO_AN) Then
            dWert = rsrs!SSTORNO_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(19).Caption = ctmp
        
        If Not IsNull(rsrs!SEINRGUTSC) Then
            dWert = rsrs!SEINRGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(32).Caption = ctmp
    
        If Not IsNull(rsrs!SRESTGUTSC) Then
            dWert = rsrs!SRESTGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(33).Caption = ctmp
    
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(41).Caption = ctmp
        
        
        ctmp = Format$(dNichtUmsReleKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(55).Caption = ctmp
        
        
        If gbGutscheinBeiVKversteuern = True Then
            ctmp = Format$(dKarte + dTilgKar + dNichtUmsReleKar, "###,###,##0.00")
        Else
            ctmp = Format$(dKarte + dGutschKar + dTilgKar + dNichtUmsReleKar, "###,###,##0.00")
        End If
        
        ctmp = ctmp & " " & gcWaehrung
        Label3(52).Caption = ctmp
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gbBargeldEingabe = True Then
        updateafcstat "TDIFF", dTDIFF, gcKasNum
        updateafcstat "KDIFF", dTDIFF, gcKasNum
    End If
    
    Label3(36).Caption = "0,00 " & gcWaehrung
    Label3(37).Caption = "0,00 " & gcWaehrung
    Label3(38).Caption = "0,00 " & gcWaehrung
    Label3(39).Caption = "0,00 " & gcWaehrung
    Label3(40).Caption = "0,00 " & gcWaehrung
    
    
    Dim dNichtUmsGutschbetrag As Double
    dNichtUmsGutschbetrag = 0
    
    If gbGutscheinBeiVKversteuern = True Then

        cSQL = "Select SUM(Wert) as UMSATZ from Gemischte_Z where kasnum = " & gcKasNum
        cSQL = cSQL & " and Thema = 'nicht ums GUTSCHBETRAG'"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!UMSATZ) Then
                dNichtUmsGutschbetrag = rsrs!UMSATZ
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        cSQL = "Select KASNUM, AMWSK, SUM(APREIS) as UMSATZ from AFCBUCH "
        cSQL = cSQL & "where KASNUM = " & gcKasNum & " and UMS_OK <> 'N' group by KASNUM, AMWSK "
    Else
    
        cSQL = "Select KASNUM, AMWSK, SUM(APREIS) as UMSATZ from AFCBUCH "
        cSQL = cSQL & "where KASNUM = " & gcKasNum & " and AARTNR <> 666666 and UMS_OK <> 'N' group by KASNUM, AMWSK "
    
    End If
    
    
    
    
    
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!AMWSK) Then
                cMWSK = rsrs!AMWSK
            Else
                    
            End If
            
            If Not IsNull(rsrs!UMSATZ) Then
                dUmsatz = rsrs!UMSATZ
            Else
                dUmsatz = 0
            End If
            
            Select Case cMWSK
                Case Is = "V"
                
                    dUmsatz = dUmsatz - dNichtUmsGutschbetrag
                
                    Label3(36).Caption = Format$(dUmsatz, "######0.00") & " " & gcWaehrung
                    Label3(37).Caption = Format$((dUmsatz / (gdMWStV + 100)) * gdMWStV, "######0.00") & " " & gcWaehrung
                Case Is = "E"
                    Label3(38).Caption = Format$(dUmsatz, "######0.00") & " " & gcWaehrung
                    Label3(39).Caption = Format$((dUmsatz / (gdMWStE + 100)) * gdMWStE, "######0.00") & " " & gcWaehrung
                Case Is = "O"
                    Label3(40).Caption = Format$(dUmsatz, "######0.00") & " " & gcWaehrung
                Case Else
                
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "AfcTempo", gdBase
    cSQL = "Select * into AFCTempo from AFCBUCH where UMS_OK = 'N' or AARTNR = 666666"
    gdBase.Execute cSQL, dbFailOnError
    
    dWert = 0
    cSQL = "Select SUM(APREIS) as UMSATZ from AFCTempo where kasnum = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!UMSATZ) Then
            dWert = rsrs!UMSATZ
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "AfcTempo", gdBase
    
    dWert = dWert + dNichtUmsGutschbetrag
    
    Label3(44).Caption = Format$(dWert, "######0.00") & " " & gcWaehrung
    
Exit Sub
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leseDatenWKL21Lokal"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseVerkaufsZahlenWKL21()
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    Dim cArtNr As String
    Dim cFeld As String
    Dim cLBSatz As String
    
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    
    Dim cKKart As String
    Dim cBedNr As String
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim dWert As Double
    Dim iFormel As Integer
    Dim dSumMenge As Double
    Dim dSumVKWert As Double
    Dim dAnzBons As Double
    Dim cBonMerker As String
    
 
    '//new
    If gcWaehrung = "EUR" Then
        iFormel = 2
        Frame1.Caption = "Verkäufe (alle Preise in " & gcWaehrung & ")"
        Frame1.ForeColor = vbCyan
    Else
        iFormel = 1
        Frame1.Caption = "Verkäufe (alle Preise in EURO)"
        Frame1.ForeColor = vbYellow
    End If
        
    FormatiereMSFlexGrid1DetailWKL21
    
    loeschNEW "AFCBUCH", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "AFCBUCH"
    
    cSQL = "Select * from AFCBUCH  order by ADATE, AZEIT, AARTNR, BESTAND desc "
    
    Set rsrs = gdApp.OpenRecordset(cSQL)
    
    dSumMenge = 0
    dSumVKWert = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzRecords = rsrs.RecordCount
        MSFlexGrid1.Rows = lAnzRecords + 2
        rsrs.MoveFirst
        lAktRecord = 1
        Do While Not rsrs.EOF
            lAktRecord = lAktRecord + 1
            MSFlexGrid1.Row = lAktRecord
            If Not IsNull(rsrs!ADATE) Then
                dWert = rsrs!ADATE
            Else
                dWert = 0
            End If
            
            If dWert > 0 Then
                cFeld = Format$(dWert, "DD.MM.YYYY")
            Else
                cFeld = ""
            End If
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = cFeld
            
            If Not IsNull(rsrs!AZEIT) Then
                cFeld = rsrs!AZEIT
            Else
                cFeld = ""
            End If
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = cFeld
            
            
            If Not IsNull(rsrs!aartnr) Then
                dWert = rsrs!aartnr
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "#####0")
            cArtNr = Trim$(cFeld)
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = cFeld
            
            
            If Not IsNull(rsrs!ABEZEICH) Then
                cFeld = rsrs!ABEZEICH
            Else
                cFeld = ""
            End If
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = cFeld
            
            If Not IsNull(rsrs!aMenge) Then
                dWert = rsrs!aMenge
            Else
                dWert = 0
            End If
            dSumMenge = dSumMenge + dWert
            cFeld = Format$(dWert, "###,##0")
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = cFeld
            
            
            If Not IsNull(rsrs!AALTPREIS) Then
                dWert = rsrs!AALTPREIS
            Else
                dWert = 0
            End If
            
            cFeld = Format$(dWert, "##,##0.00")
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Text = cFeld

            
            
            If Not IsNull(rsrs!APREIS) Then
                dWert = rsrs!APREIS
            Else
                dWert = 0
            End If

            dSumVKWert = dSumVKWert + dWert
            cFeld = Format$(dWert, "##,##0.00")
            MSFlexGrid1.Col = 6
            MSFlexGrid1.Text = cFeld
                
            If Not IsNull(rsrs!AVKPR) Then
                dWert = rsrs!AVKPR
            Else
                dWert = 0
            End If

            cFeld = Format$(dWert, "##,##0.00")
            MSFlexGrid1.Col = 7
            MSFlexGrid1.Text = cFeld
                
            If Not IsNull(rsrs!BESTAND) Then
                dWert = rsrs!BESTAND
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "###,##0")
            MSFlexGrid1.Col = 8
            MSFlexGrid1.Text = cFeld
                
            If Not IsNull(rsrs!kk_art) Then
                cKKart = rsrs!kk_art
            Else
                cKKart = ""
            End If
            MSFlexGrid1.Col = 9
            MSFlexGrid1.Text = cKKart
            
            If Not IsNull(rsrs!abednu) Then
                cBedNr = rsrs!abednu
            Else
                cBedNr = ""
            End If
            MSFlexGrid1.Col = 10
            MSFlexGrid1.Text = cBedNr
            
            If Not IsNull(rsrs!BELEGNR) Then
                cFeld = rsrs!BELEGNR
            Else
                cFeld = ""
            End If
            If cFeld <> cBonMerker Then
                dAnzBons = dAnzBons + 1
                cBonMerker = cFeld
            End If
            MSFlexGrid1.Col = 11
            MSFlexGrid1.Text = cFeld
            
            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                dWert = rsrs!ZHLGGUTSCH
            Else
                dWert = 0
            End If

            cFeld = Format$(dWert, "##,##0.00")
            MSFlexGrid1.Col = 12
            MSFlexGrid1.Text = cFeld
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '*************************************
    '* Nachlauf Summenzeile
    '*************************************
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 2
    MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = "Summe:"
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Text = "Zeilen Umsatz + nicht relv. Umsatz"
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Text = Format$(dSumMenge, "######0")
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = Format$(dSumVKWert, "######0.00")
    
    MSFlexGrid1.Col = 11
    MSFlexGrid1.Text = Format$(dAnzBons, "######0")
    
    MSFlexGrid1.RowHeight(1) = 0
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.15 * gdTabfak
    
    MSFlexGrid1.Visible = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseVerkaufsZahlenWKL21"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    With gridx
    
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet        As Integer
    Dim iFileNr     As Integer
    Dim ctmp        As String
    Dim dWert       As Double
    Dim cAltesDatum As String
    Dim sSQL        As String
    Dim rsAFCD      As Recordset
    Dim rsArt       As Recordset
    Dim rsrs        As Recordset
    Dim lMindat     As Long

    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 6 'nach Lieferanten
        
            loeschNEW "TAKOLLKASS", gdBase
            CreateTableT2 "TAKOLLKASS", gdBase
        
            loeschNEW "AFCDL", gdBase
            CreateTableT2 "AFCDL", gdBase
            
            If Check2.value = vbChecked Then
                Sortierung CByte(gcKasNum)
            Else
                Sortierung 99
            End If
            
            sSQL = "Insert into TAKOLLKASS Select k.LINR "
            sSQL = sSQL & ", k.AArtnr"
            sSQL = sSQL & ", k.ABEZEICH"
            sSQL = sSQL & ", sum(k.AMENGE)as Menge"
            sSQL = sSQL & ", aaltpreis as Apreis "
            sSQL = sSQL & " from afcbuch k "
            
            If Check2.value = vbChecked Then
                sSQL = sSQL & " where k.kasnum = " & gcKasNum
            End If
            
            sSQL = sSQL & " group by AArtnr"
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", ABEZEICH "
            sSQL = sSQL & ", Aaltpreis "
            gdBase.Execute sSQL, dbFailOnError
            
            lMindat = 0
            sSQL = "Select min(adate) as miniDat from AFCBUCH  "
            If Check2.value = vbChecked Then
                sSQL = sSQL & " where kasnum = " & gcKasNum
            End If
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!miniDat) Then
                    lMindat = CLng(rsrs!miniDat)
                End If
                
            End If
            rsrs.Close: Set rsrs = Nothing
           
            
            
            
            'hier kollverk
            
            If lMindat > 0 Then
            
                iRet = MsgBox("Möchten Sie zusätzlich auch die Kollegenverkäufe seit dem letzten Tagesabschluss aufgelistet haben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbYes Then
                    sSQL = "Insert into TAKOLLKASS Select k.LINR "
                    sSQL = sSQL & ", k.Artnr as aartnr"
                    sSQL = sSQL & ", k.BEZEICH as abezeich"
                    sSQL = sSQL & ", sum(k.MENGE)as Menge"
                    sSQL = sSQL & ", k.preis as Apreis "
                    sSQL = sSQL & " from Kollverk k"
                    sSQL = sSQL & " where k.adate >= " & lMindat 'datevalue(now) "
                    If Check2.value = vbChecked Then
                        sSQL = sSQL & " and k.kasnum = " & gcKasNum
                    End If
                    
                    sSQL = sSQL & " group by k.Artnr"
                    sSQL = sSQL & ", k.LINR "
                    sSQL = sSQL & ", k.BEZEICH"
                    sSQL = sSQL & ", k.preis "
                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
            
            'hier Kollvk ende
            
            sSQL = "Insert into AFCDL Select t.LINR "
            sSQL = sSQL & ", l.LIEFBEZ "
            sSQL = sSQL & ", a.EAN "
            sSQL = sSQL & ", AArtnr "
            sSQL = sSQL & ", ABEZEICH "
            sSQL = sSQL & ", sum(t.MENGE)as AMenge "
            sSQL = sSQL & ", Apreis "
            sSQL = sSQL & ", a.bestand "
            sSQL = sSQL & " from TAKOLLKASS T, lisrt l, artikel a "
            sSQL = sSQL & " WHERE t.AARTNr = a.artnr and t.linr = l.linr "
            sSQL = sSQL & " group by t.AArtnr "
            sSQL = sSQL & ", t.LINR "
            sSQL = sSQL & ", l.LIEFBEZ "
            sSQL = sSQL & ", a.EAN "
            sSQL = sSQL & ", t.ABEZEICH "
            sSQL = sSQL & ", a.bestand "
            sSQL = sSQL & ", t.apreis "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update AFCDL Set Firmaort =  '" & gFirma.Ort & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            reportbildschirm "WKL010n", "aWKL21"
            
        Case Is = 7 'nach AGN
            loeschNEW "afcd2", gdBase
            CreateTable "AFCD2", gdBase
            
            sSQL = "Insert into AFCD2 Select artikel.EAN"
            sSQL = sSQL & ", artikel.AGN"
            sSQL = sSQL & ", afcbuch.AArtnr"
            sSQL = sSQL & ", agndbf.agtext"
            sSQL = sSQL & ", afcbuch.ABEZEICH"
            sSQL = sSQL & ", sum(afcbuch.AMENGE)as AMenge"
            sSQL = sSQL & ", afcbuch.aaltpreis as Avkpr "
            sSQL = sSQL & ", artikel.bestand "
            sSQL = sSQL & " from afcbuch, agndbf, artikel"
            sSQL = sSQL & " WHERE afcbuch.AARTNr = artikel.artnr and artikel.agn = agndbf.agn "
            sSQL = sSQL & " group by afcbuch.AArtnr"
            sSQL = sSQL & ", artikel.AGN"
            sSQL = sSQL & ", artikel.EAN"
            sSQL = sSQL & ", afcbuch.ABEZEICH"
            sSQL = sSQL & ", agndbf.agtext"
            sSQL = sSQL & ", artikel.bestand "
            sSQL = sSQL & ", afcbuch.aaltpreis "
            gdBase.Execute sSQL, dbFailOnError
            
            print_firma "FIRMA_21A"
            
            reportbildschirm "WKL010", "aWKL21a"

        Case Is = 8 'Z-Bon drucken
            '********************
            '* nur Kopfdaten    *
            '********************
            If gsZBon = "" Then
                iRet = MsgBox("Wollen Sie die Daten auf dem Bondrucker ausdrucken?", vbQuestion + vbYesNo, "DRUCKER WÄHLEN")
                If iRet <> vbYes Then   'Listendrucker gewählt
                    
                    speicherZbon "Listendrucker"
                    gsZBon = "Listendrucker"
                Else 'Bondrucker gewählt
                    speicherZbon "Bondrucker"
                    gsZBon = "Bondrucker"
                End If
            End If
            
            Pbr1.Visible = True
            Pbr1.Max = 100
            lblStat.Visible = True
            lblStat.Caption = "bitte warten..."
            lblStat.Refresh
            
            If gsZBon = "Listendrucker" Then
                Pbr1.value = 10
                
                lblStat.Caption = "warten..."
                lblStat.Refresh
                DruckeTagesabschlussNeuWKL21a 1, False
                Pbr1.value = 20
                
                lblStat.Caption = "bitte warten..."
                lblStat.Refresh
                If Modul6.FindFile(gcDBPfad, "aWKL21z.rpt") Then
                    If gbQZBON = True Then
                        reportbildschirmToPrinter "aWKL21z"
                    Else
                        reportbildschirm "", "aWKL21z"
                    End If
                Else
                    If gbQZBON = True Then
                        If gbZBONDINA4HOCH = True Then
                            reportbildschirmToPrinter "aWKL21bh"
                        Else
                            reportbildschirmToPrinter "aWKL21b"
                        End If
                    Else
                        If gbZBONDINA4HOCH = True Then
                            reportbildschirm "", "aWKL21bh"
                        Else
                            reportbildschirm "", "aWKL21b"
                        End If
                        
                    End If
                End If
                Pbr1.value = 30
                
                lblStat.Caption = "warten..."
                lblStat.Refresh
                
    
                
                    
            ElseIf gsZBon = "Bondrucker" Then
                Pbr1.value = 10
                
                lblStat.Caption = "warten..."
                lblStat.Refresh
                DruckeTagesAbschlussAufBonDruckerWKL21a
                Pbr1.value = 20
                lblStat.Caption = "bitte warten..."
                lblStat.Refresh
                DruckeKassenEinAuszahlungAufBonDruckerWKL21a
                Pbr1.value = 30
                
                If gbAGNAusw = True Then
                    DruckeKassenAgnAuswertungaufBondrucker
                    Pbr1.value = 40
                End If
            End If
    
            Dim cPfad1 As String
            Dim cName As String
            Dim lWert As Long
            Dim sTime As String
            sTime = TimeValue(Now)
            sTime = Right(sTime, 8)
            sTime = Left(sTime, 5)

            lWert = DateValue(Now)
            ctmp = Format$(lWert, "MM.DD")
           
            ctmp = ctmp & sTime
            ctmp = SwapStr(ctmp, ".", "")
            ctmp = SwapStr(ctmp, ":", "")
    
            cName = ctmp
            
            cPfad1 = gcDBPfad
            If Right$(cPfad1, 1) <> "\" Then
                cPfad1 = cPfad1 & "\"
            End If
            
            Pbr1.value = 50
            lblStat.Caption = "warten......"
            lblStat.Refresh
            DruckeTagesabschlussNeuWKL21a 1, True
            Pbr1.value = 55
            

            
            Pbr1.value = 60
            lblStat.Caption = "bitte warten..."
            lblStat.Refresh
            
            cPfad1 = gcDBPfad
            If Right$(cPfad1, 1) <> "\" Then
                cPfad1 = cPfad1 & "\"
            End If
            
            If gbMitExport Then 'Mit Export - abschalten wenn es Probleme gibt wie bei der Münze
            
            
                If gbZBONDINA4HOCH = True Then
                    reportbildschirmtoText "aWKL21bh", cPfad1 & "ABPRO\" & cName & "_" & gcKasNum & ".txt"
                Else
                    reportbildschirmtoText "aWKL21b", cPfad1 & "ABPRO\" & cName & "_" & gcKasNum & ".txt"
                End If
            
            
                
                
                If FileExists(cPfad1 & "ABPRO\" & cName & "_" & gcKasNum & ".txt") Then
                    schreibeProtoAbschluss "Sicherung erfolgt " & gcKasNum & ""
                End If
            End If

            Pbr1.value = 75
            
            cPfad1 = gcDBPfad
            If Right$(cPfad1, 1) <> "\" Then
                cPfad1 = cPfad1 & "\"
            End If
            
            If gbMitExport Then 'Mit Export - abschalten wenn es Probleme gibt wie bei der Münze
                Kill cPfad1 & "Export\*.txt"
                
                
                If gbZBONDINA4HOCH = True Then
                    reportbildschirmtoText "aWKL21bh", cPfad1 & "Export\" & cName & "_" & gcKasNum & ".txt"
                Else
                    reportbildschirmtoText "aWKL21b", cPfad1 & "Export\" & cName & "_" & gcKasNum & ".txt"
                End If
                
                
                
                
                If gsKaMail <> "" Then
                    Dim sAttachment As String
                    sAttachment = cPfad1 & "Export\" & cName & "_" & gcKasNum & ".txt"
                    Dim sMess As String
    
                    sMess = "Der Tagesabschluss wurde soeben erstellt."


                    schickeMailimHintergrundSSL ermFirmenBez, gsKaMail, gsKaMail, gsKaMail _
                    , gsKaMail, gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, "Tagesabschluss", sMess, sAttachment
                    

                End If
            End If
            
            Pbr1.Visible = False
            lblStat.Visible = False

        Case Is = 10        'nicht umsatzrelevante VKs
        
            loeschNEW "afcnu", gdBase
            CreateTable "AFCNU", gdBase
            
            sSQL = "Insert into AFCNU Select afcbuch.LINR "
            sSQL = sSQL & ", afcbuch.AArtnr"
            sSQL = sSQL & ", afcbuch.ABEZEICH"
            sSQL = sSQL & ", afcbuch.AMENGE "
            sSQL = sSQL & ", afcbuch.ALEKPR "
            sSQL = sSQL & ", artikel.bestand "
            sSQL = sSQL & ", afcbuch.ABEDNU"
            sSQL = sSQL & ", afcbuch.aPreis "
            sSQL = sSQL & ", afcbuch.azeit "
            sSQL = sSQL & ", afcbuch.belegnr "
            sSQL = sSQL & ", afcbuch.kasnum "
            sSQL = sSQL & ", afcbuch.kk_art "
            sSQL = sSQL & "  from AFCBUCH inner join artikel on "
            sSQL = sSQL & " afcbuch.aartnr = artikel.artnr "
            sSQL = sSQL & " where AFCBUCH.UMS_OK = 'N' "
            sSQL = sSQL & " order by AFCBUCH.ADATE, AFCBUCH.AZEIT, AFCBUCH.AARTNR, artikel.BESTAND desc "
            gdBase.Execute sSQL, dbFailOnError
            
            print_firma "FIRMA_21D"
            
            reportbildschirm "WKL003", "aWKL21d"
            
        Case Is = 12        'Kumulierte Artikel
            Command1(12).Enabled = False
            artikelkumulierttoprinter
            
            If check1.value = vbChecked Then
                Sortierung 1
            Else
                Sortierung 2
            End If
            
            If gbKUMSUM = True Then
                ' anzeigen
                KUMSUM 1
            Else
                'nicht anzeigen
                KUMSUM 2
            End If
            
            Screen.MousePointer = 11
            reportbildschirm "WKL003dn", "aWKL21e"
            Screen.MousePointer = 0
            Command1(12).Enabled = True
        Case Is = 13  '//Detaillieren Artikel
            loeschNEW "AFCD4", gdBase
            CreateTable "AFCD4", gdBase
            
            sSQL = "Insert into AFCD4 Select afcbuch.LINR "
            sSQL = sSQL & ", afcbuch.AArtnr"
            sSQL = sSQL & ", afcbuch.ABEZEICH"
            sSQL = sSQL & ", afcbuch.AMENGE "
            sSQL = sSQL & ", afcbuch.aaltpreis as APreis "
            sSQL = sSQL & ", afcbuch.ALEKPR "
            sSQL = sSQL & ", artikel.bestand "
            sSQL = sSQL & ", afcbuch.ABEDNU "
            sSQL = sSQL & ", afcbuch.adate "
            sSQL = sSQL & ", afcbuch.azeit "
            sSQL = sSQL & ", afcbuch.belegnr "
            sSQL = sSQL & ", afcbuch.kasnum "
            sSQL = sSQL & ", afcbuch.kk_art "
            sSQL = sSQL & ", afcbuch.zhlggutsch "
            sSQL = sSQL & " "
            sSQL = sSQL & " from afcbuch, artikel"
            sSQL = sSQL & " WHERE afcbuch.AARTNr = artikel.artnr "

            print_firma "FIRMA_21F"
            
            gdBase.Execute sSQL, dbFailOnError
            reportbildschirm "WKL003cn", "aWKL21f"
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 70 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR
    
    Dim cART As String
    Dim iFilnr As Integer
    iFilnr = CInt(gcFilNr)
    
    If iFilnr > 0 Then
        
        MSFlexGrid1.Col = 2
         
        If MSFlexGrid1.Row > 0 Then
        
            cART = Trim(MSFlexGrid1.Text)
            If IsNumeric(cART) Then
                gcArtNrFiliale = cART
                frmWKLae.Show 1
            End If
         
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub SSCommand1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet            As Integer
    Dim iFileNr         As Integer
    Dim ctmp            As String
    Dim dWert           As Double
    Dim cAltesDatum     As String
    Dim cPfad1          As String
    Dim cNeuerAbschluß  As String
    Dim cAlterAbschluß  As String
    Dim lcount          As Long
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    Select Case index
        Case Is = 0
            
            Screen.MousePointer = 11
            LeseVerkaufsZahlenWKL21
            Frame1.Visible = True
            
            Command1(12).Enabled = ReportVorhanden("aWKL21e")
            Command1(13).Enabled = ReportVorhanden("aWKL21f")
            Command1(10).Enabled = ReportVorhanden("aWKL21d")
            Command1(6).Enabled = ReportVorhanden("aWKL21")
            Command1(7).Enabled = ReportVorhanden("aWKL21a")
            
            SSCommand1(0).Visible = False
            SSCommand1(1).Visible = False
            SSCommand1(5).Visible = False
            Command1(8).SetFocus
            Screen.MousePointer = 0
        Case 2
            Select Case gsEPartner
                Case Is = "ELP"
                    lese_ELPAY_opt
                    setzedrucker gcBonDrucker
                    Kassenschnitt_elPAY
                Case Is = "ZVT"
                    lese_ZVT_opt
                    setzedrucker gcBonDrucker
                    Kassenschnitt_ZVT
                Case Is = "ZV2"
                    lese_ZVT_opt2
                    
                    Kassenschnitt_ZVT2 False
            End Select
            
        Case Is = 1
            SSCommand1(1).Enabled = False
            If gbLokalModus Then
                If frmWKL00.Command3(1).Caption = "Z Bon" Then
                    Me.Refresh      'X bon aus autolokalmodus
                    Command1_Click 8
                    Me.Refresh
                Else
                    MsgBox "Sie befinden sich noch im 'Offline - Betrieb'. Synchronisieren Sie Ihre Datenbank und führen dann den Kassenabschluss durch!", vbInformation, "Winkiss Hinweis:"
                    Unload frmWKL21
                    Exit Sub
                End If
            End If
            
            If frmWKL00.Command3(1).Caption = "Z Bon" Then
                Me.Refresh
                Command1_Click 8
                Me.Refresh
            Else
                If gbBargeldEingabe = True Then
                    iRet = vbYes
                Else
                    iRet = MsgBox("Wollen Sie jetzt einen Kassenabschluß durchführen?", vbYesNo + vbQuestion + vbDefaultButton2, "KASSENABSCHLUSS?")
                End If
                
                
    
                If iRet = vbYes Then
                
                    'YES Abschluss
                    
                    cNeuerAbschluß = ""
                    cAlterAbschluß = ""
    
                    cNeuerAbschluß = Format$(Fix(Now), "DD.MM.YYYY") & " " & Format$(Now, "HH:MM:SS")
                    
            
                    iFileNr = FreeFile
                    Open gcDBPfad & "\ABSCHLUS.TXT" For Binary As #iFileNr
                    If LOF(iFileNr) > 0 Then
                        cAlterAbschluß = Space$(LOF(iFileNr))
                        Get #iFileNr, 1, cAlterAbschluß
                        Close iFileNr
                    Else
                        Close iFileNr
                        cAlterAbschluß = "00.00.0000 00:00:00 000000"
                    End If
            
                    Do While Right(cAlterAbschluß, 1) = vbCr Or Right(cAlterAbschluß, 1) = vbLf
                        cAlterAbschluß = Left(cAlterAbschluß, Len(cAlterAbschluß) - 1)
                    Loop
                    ctmp = Right(cAlterAbschluß, 6)
                    lcount = Val(ctmp)
                    lcount = lcount + 1
                    ctmp = Format$(lcount, "000000")
                    cNeuerAbschluß = cNeuerAbschluß & " " & ctmp
                    gsNeuerAbschluß = cNeuerAbschluß
            

        
                    cAlterAbschluß = "vorheriger Abschluß: " & cAlterAbschluß
                    cNeuerAbschluß = "Jetziger Abschluß:   " & cNeuerAbschluß
                    
                    altLbl.Caption = cAlterAbschluß
                    neuLbl.Caption = cNeuerAbschluß
                    
                    GDPdU_ZBON_sichern_Teil1 cAlterAbschluß, cNeuerAbschluß
                    
                    'zbon jetzt immer
                    Command1_Click 8
                
                    If gbQZBON = True Then
                        Me.Refresh
                        
                        
                        
                        lblStat.Visible = True
                        Pbr1.Visible = True
                        Pbr1.Max = 100
                        Pbr1.value = 10
                        
                        If gbARTKUM Then
                        
                            lblStat.Caption = "bitte warten..."
                            lblStat.Refresh
                            
                            artikelkumulierttoprinter
                            
                            Pause (2)
                            
                            lblStat.Caption = "warten..."
                            lblStat.Refresh
                            
                            Pbr1.value = 20
                            Sortierung 1 'ohne ekpr
                            Screen.MousePointer = 11
                            
                            If gbKUMSUM = True Then
                                ' anzeigen
                                KUMSUM 1
                            Else
                                'nicht anzeigen
                                KUMSUM 2
                            End If
                            reportbildschirmToPrinter "aWKL21e"
                            
                            lblStat.Caption = "bitte warten..."
                            lblStat.Refresh
                            Pbr1.value = 25
                            
                            
                            Screen.MousePointer = 0
                        End If
                        
                        If gbTAGFILT Then
                        
                            lblStat.Caption = "bitte warten..."
                            lblStat.Refresh
                            
                            Filtautoprinter
                            
                            lblStat.Caption = "warten..."
                            lblStat.Refresh
                            
                            If Datendrin("AFCD5", gdBase) Then
                                Pbr1.value = 30
                    
                                Screen.MousePointer = 11
                                
                                reportbildschirmToPrinter "aWKL21g"
                                reportbildschirmToPrinter "aWKL21h"
                                
                                Pause (2)
                                
                                lblStat.Caption = "bitte warten..."
                                lblStat.Refresh
                                Pbr1.value = 35
                                
                                If gbMitExport Then 'Mit Export - abschalten wenn es Probleme gibt wie bei der Münze
                                
                                    cPfad1 = gcDBPfad
                                    If Right$(cPfad1, 1) <> "\" Then
                                        cPfad1 = cPfad1 & "\"
                                    End If
                                    reportbildschirmtoText "aWKL21g", cPfad1 & "ABPRO\Filtausch.txt"
                                End If
                            End If
                            Screen.MousePointer = 0
                        End If
                        
                        If gbKK Then
                            Pbr1.value = 40
                            
                            lblStat.Caption = "warten..."
                            lblStat.Refresh
                            
                            If Datendrin("KKZAHLTE", gdBase) Then

                                Pbr1.value = 45
                                
                                lblStat.Caption = "bitte warten..."
                                lblStat.Refresh
                                
                                'Neuen Datenhintergrund erstellen
                                
                                Print_KKZAHLTE_Vorbereitung

                                'Ende Neuen Datenhintergrund erstellen
                                
                                If gcListenDrucker <> gcBonDrucker Then
                                    reportbildschirmToPrinter "aWKLkk"
                                End If
                                Pbr1.value = 50

                                lblStat.Caption = "warten..."
                                lblStat.Refresh

                                If gbMitExport Then 'Mit Export - abschalten wenn es Probleme gibt wie bei der Münze
                                
                                    cPfad1 = gcDBPfad
                                    If Right$(cPfad1, 1) <> "\" Then
                                        cPfad1 = cPfad1 & "\"
                                    End If
                                    reportbildschirmtoText "aWKLkk", cPfad1 & "ABPRO\KKzahlungen_" & gcKasNum & ".txt"
                                End If
                            End If
                            
                            Pause (2)  'plassmann
                            
                            KKtoprinter
                            
                            If Datendrin("LASTZAHLTE", gdBase) Then

                                Pbr1.value = 40
                                
                                lblStat.Caption = "bitte warten..."
                                lblStat.Refresh
                                
                                'Neuen Datenhintergrund erstellen
                            
                                Print_LASTZAHLTE_Vorbereitung
                                'Ende Neuen Datenhintergrund erstellen

                                If gcListenDrucker <> gcBonDrucker Then
                                    reportbildschirmToPrinter "aWKLLS"
                                End If
                                Pbr1.value = 50

                                lblStat.Caption = "warten..."
                                lblStat.Refresh

                                If gbMitExport Then 'Mit Export - abschalten wenn es Probleme gibt wie bei der Münze
                                
                                    cPfad1 = gcDBPfad
                                    If Right$(cPfad1, 1) <> "\" Then
                                        cPfad1 = cPfad1 & "\"
                                    End If
                                    
                                    reportbildschirmtoText "aWKLLS", cPfad1 & "ABPRO\ECLAST_" & gcKasNum & ".txt"
                                End If
                            End If
                            
                            Lasttoprinter
                        End If
                        
'                        KKZAHLTEtoKKZAHL
'                        LastZahltetoLastzahl
                        
                        
                        
                        If gbEA Then
                            Pbr1.value = 60
                            lblStat.Caption = "bitte warten..."
                            lblStat.Refresh
                            If DruckeKassenEinAuszahlungWKL21 = True Then
                            
                                lblStat.Caption = "warten..."
                                lblStat.Refresh
                                Pbr1.value = 70
                                
                                reportbildschirmToPrinter "aWKL21c"
                                
                                Pbr1.value = 80
                                
                                lblStat.Caption = "bitte warten..."
                                lblStat.Refresh
                                
                                If gbMitExport Then 'Mit Export - abschalten wenn es Probleme gibt wie bei der Münze
                                
                                    cPfad1 = gcDBPfad
                                    If Right$(cPfad1, 1) <> "\" Then
                                        cPfad1 = cPfad1 & "\"
                                    End If
                                    
                                    reportbildschirmtoText "aWKL21c", cPfad1 & "ABPRO\EinAuszahlungen" & gcKasNum & ".txt"
                                End If
                                
                            End If
                        End If
                        Pbr1.Visible = False
                        
                        Me.Refresh
                        
                    End If
                    
                    KKZAHLTEtoKKZAHL
                    LastZahltetoLastzahl
                        
                    lblStat.Visible = False

                    
                    
                    gbDate = False
                    
                    '********Thomas Statistiken Ende
                    
                    
    
                    Screen.MousePointer = 0
                    frmWK21d.Show 1 'Tagesabschluss
                Else
                    If gbBargeldEingabe = True Then
                        schreibeProtoAbschluss "Kassenabschluss wurde abgebrochen " & gcKasNum & "--------"
                    End If
                End If
            
            End If
            Unload frmWKL21
            
            If ermaktUmsatz(False) > 0 Then
                frmWKL00.Command3(1).BackColor = vbRed
            Else
                frmWKL00.Command3(1).BackColor = &H8000000F
            End If
             
            If gbLocalSec Then
                If gbAutoLokalModus Then
                
                    If gbLokalModus = False Then
                        anzeige "normal", "Offline - Betrieb", frmWKL00.Label2
                        frmWKL00.ChkLM.value = vbChecked
                    End If
                End If
            End If
        Case Is = 5
            frmWK21b.Show 1
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SSCommand1_Click"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten. " & index
        
        Fehlermeldung1
    End If
End Sub
Private Sub artikelkumulierttoprinter()
On Error GoTo LOKAL_ERROR

Dim sSQL        As String
Dim rsrs        As Recordset
Dim cArtNr      As String
Dim lGBestand   As Long
Dim i           As Integer

loeschNEW "afcd03", gdBase
CreateTable "AFCD03", gdBase

Screen.MousePointer = 11

sSQL = "Insert into AFCD03 Select a.LINR "
sSQL = sSQL & ", a.AArtnr"
sSQL = sSQL & ", a.ABEZEICH"
sSQL = sSQL & ", a.AMenge as ameng"
sSQL = sSQL & ", a.APreis as aprei"
sSQL = sSQL & ", a.ALEKPR "
sSQL = sSQL & ", a.AMWSK as AMWST"
sSQL = sSQL & ", 0 as Bestand "
sSQL = sSQL & " "
sSQL = sSQL & " from afcbuch a  "
gdBase.Execute sSQL, dbFailOnError


sSQL = "Update AFCD03 "
sSQL = sSQL & " set ABEZEICH = 'Gutschein' where aartnr = 666666"
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update AFCD03 "
sSQL = sSQL & " set ne = (aPrei * 100) /(100 + " & gdMWStV & ") - (ALEKPR * AMeng)"
sSQL = sSQL & " where amwst ='V' "
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update AFCD03 "
sSQL = sSQL & " set ne = (aPrei * 100) /(100 + " & gdMWStO & ") - (ALEKPR * AMeng)"
sSQL = sSQL & " where amwst ='O' "
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update AFCD03 "
sSQL = sSQL & " set ne = (aPrei * 100) /(100 + " & gdMWStE & ") - (ALEKPR * AMeng)"
sSQL = sSQL & " where amwst ='E' "
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update AFCD03 "
sSQL = sSQL & " set nettopreis = (aPrei * 100) /(100 + " & gdMWStV & ") "
sSQL = sSQL & " where amwst ='V' "
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update AFCD03 "
sSQL = sSQL & " set nettopreis = (aPrei * 100) /(100 + " & gdMWStO & ") "
sSQL = sSQL & " where amwst ='O' "
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update AFCD03 "
sSQL = sSQL & " set nettopreis = (aPrei * 100) /(100 + " & gdMWStE & ") "
sSQL = sSQL & " where amwst ='E' "
gdBase.Execute sSQL, dbFailOnError

loeschNEW "afcd3", gdBase
CreateTable "AFCD3", gdBase

sSQL = "Insert into AFCD3 Select a.LINR "
sSQL = sSQL & ", a.AArtnr"
sSQL = sSQL & ", a.ABEZEICH"
sSQL = sSQL & ", sum(a.AMENG)as AMenge"
sSQL = sSQL & ", sum(a.APREI)as APreis"
sSQL = sSQL & ", sum(a.ne)as nse"
sSQL = sSQL & ", sum(a.nettopreis)as nettopr"
sSQL = sSQL & ", a.ALEKPR "
sSQL = sSQL & ", 0 as WGN "
sSQL = sSQL & ", 0 as GBestand "
sSQL = sSQL & " from AFCD03 a "
sSQL = sSQL & " group by a.AArtnr"
sSQL = sSQL & ", a.LINR "
sSQL = sSQL & ", a.ABEZEICH"
sSQL = sSQL & ", a.ALEKPR "
gdBase.Execute sSQL, dbFailOnError


sSQL = "Update AFCD3 "
sSQL = sSQL & " set nettopr = 0 , nse = 0 where aartnr = 666666"
gdBase.Execute sSQL, dbFailOnError



sSQL = " Update AFCD3 inner join WARENGRU on AFCD3.AARTNR = WARENGRU.Artnr "
sSQL = sSQL & " set AFCD3.WGN = WARENGRU.WGNR "
gdBase.Execute sSQL, dbFailOnError

sSQL = " Update AFCD3 set WGN = 200 where wgn = 0 "
gdBase.Execute sSQL, dbFailOnError

sSQL = " Update AFCD3 inner join Artikel on AFCD3.AARTNR = Artikel.Artnr "
sSQL = sSQL & " set AFCD3.BESTAND = ARTIKEL.BESTAND "
sSQL = sSQL & " , AFCD3.FARBNR = val(ARTIKEL.awm) "
gdBase.Execute sSQL, dbFailOnError

'Gesamtbestand füllen
If gcFilNr <> "0" Then
    Set rsrs = gdBase.OpenRecordset("AFCD3", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cArtNr = ""
            If Not IsNull(rsrs!aartnr) Then
                cArtNr = rsrs!aartnr
            End If
            
            lGBestand = 0
            
            If cArtNr <> "" Then
                For i = 1 To giAnzFil
                    If i = CInt(gcFilNr) Then
                    
                    Else
                        lGBestand = lGBestand + ermBestandfromZbestand(cArtNr, i)
                    End If
                Next i
            End If
            
            rsrs.Edit
            rsrs!Gbestand = lGBestand
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
End If

BringFarbeInsSpiel "AFCD3", gdBase


If gbARTKUM_ohneWGN = True Then
    sSQL = "Delete * from AFCD3 where wgn <> 200 "
    gdBase.Execute sSQL, dbFailOnError
End If





If gbARTKUMWGN = True Then

    loeschNEW "afctemp", gdBase
    
    sSQL = "Select * into afctemp from AFCD3 "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "afcd3", gdBase
    CreateTable "AFCD3", gdBase
    
    sSQL = "Insert into AFCD3 Select * from afctemp order by WGN "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "afctemp", gdBase
Else

    loeschNEW "afctemp", gdBase
    
    sSQL = "Select * into afctemp from AFCD3 "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "afcd3", gdBase
    CreateTable "AFCD3", gdBase
    
    sSQL = "Insert into AFCD3 Select * from afctemp order by Linr,AArtnr "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "afctemp", gdBase
End If



Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikelkumulierttoprinter"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Filtautoprinter()
On Error GoTo LOKAL_ERROR

Dim sSQL As String

loeschNEW "AFCD5", gdBase
CreateTable "AFCD5", gdBase

Screen.MousePointer = 11

sSQL = "Insert into AFCD5 Select ARTNR "
sSQL = sSQL & ", BEZEICH "
sSQL = sSQL & ", Menge "
sSQL = sSQL & ", LINR "
sSQL = sSQL & ", LPZ "
sSQL = sSQL & ", FIL_AN "
sSQL = sSQL & ", KASNUM "
sSQL = sSQL & ", ADATE "
sSQL = sSQL & ", AZEIT "
sSQL = sSQL & ", KVKPR1 "
sSQL = sSQL & ", vkpr "
sSQL = sSQL & ", lekpr "
sSQL = sSQL & ", ekpr "
sSQL = sSQL & ", BEDIENER "
sSQL = sSQL & " from TAUSCH "
sSQL = sSQL & " where sendok = false "
gdBase.Execute sSQL, dbFailOnError

sSQL = " Update AFCD5 inner join Artikel on AFCD5.ARTNR = Artikel.Artnr "
sSQL = sSQL & " set AFCD5.BESTAND = ARTIKEL.BESTAND "
gdBase.Execute sSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Filtautoprinter"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Print_KKZAHLTE_Vorbereitung()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

loeschNEW "KKZAHLTE_PRINT", gdBase
CreateTableT2 "KKZAHLTE_PRINT", gdBase

cSQL = "Insert into KKZAHLTE_PRINT Select * from KKZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Print_KKZAHLTE_Vorbereitung"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Print_LASTZAHLTE_Vorbereitung()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

loeschNEW "LASTZAHLTE_PRINT", gdBase
CreateTableT2 "LASTZAHLTE_PRINT", gdBase

cSQL = "Insert into LASTZAHLTE_PRINT Select * from LASTZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Print_LASTZAHLTE_Vorbereitung"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KKtoprinter()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

cSQL = "Insert into KKZAHL Select * from KKZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from KKZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KKtoprinter"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Lasttoprinter()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

cSQL = "Insert into LastZAHL Select * from LastZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from LastZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Lasttoprinter"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KKZAHLTEtoKKZAHL()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

cSQL = "Insert into KKZAHL Select * from KKZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from KKZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KKZAHLTEtoKKZAHL"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LastZahltetoLastzahl()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

cSQL = "Insert into LastZAHL Select * from LastZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from LastZAHLTE where kasnum = " & gcKasNum
gdBase.Execute cSQL, dbFailOnError

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LastZahltetoLastzahl"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
    SSCommand1(0).Visible = True
    SSCommand1(1).Visible = True
    SSCommand1(5).Visible = True
'    SSCommand1(11).Visible = True
    Frame1.Visible = False
    SSCommand1(1).SetFocus
    
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnz    As Long
    Dim lcount  As Long
    Dim cDatum  As String
    Dim iRet    As Integer
    
    Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    SSCommand1(1).Enabled = True
    
    Screen.MousePointer = 11
    
    cDatum = Format$(Now, "DD.MM.YYYY")
    Label2.Caption = gcWochentag(Weekday(Fix(Now), 2)) & ", den " & cDatum
    
    For lAnz = 0 To 55
        Label3(lAnz).Caption = "0,00 " & gcWaehrung
    Next lAnz
    
    'wie immer
    Label4.Caption = "Der Verkauf von Gutscheinen sowie die Kredittilgung sind nicht umsatzrelevant und werden daher NICHT "
    Label4.Caption = Label4.Caption & "in den Umsatz-Summen berücksichtigt! Umsatzrelevant sind Kreditverkäufe und eingelöste Gutscheine."
    
    
    If gbGutscheinBeiVKversteuern = True Then
    
        
        
        Dim dateStichtag As Date
        dateStichtag = ermStichtag
        
        If dateStichtag <= DateValue(Now) Then
            Label4.Caption = "Die Einlösung von Gutscheinen sowie die Kredittilgung sind nicht umsatzrelevant und werden daher NICHT "
            Label4.Caption = Label4.Caption & "in den Umsatz-Summen berücksichtigt! Umsatzrelevant sind Kreditverkäufe und "
            Label4.Caption = Label4.Caption & "der Verkauf von Gutscheinen."
        End If
        
        
    End If
    
    
    

    

'        If Datendrin("BONPAUSE", gdBase) Then
'            MsgBox "Es sind noch unterbrochene Kassiervorgänge vorhanden.", vbOKOnly + vbInformation, "Winkiss Hinweis:"
'        End If

    
    If gsEPartner = "ELP" Or gsEPartner = "ZVT" Or gsEPartner = "ZV2" Then
        SSCommand1(2).Enabled = True
    Else
        SSCommand1(2).Enabled = False
    End If
    
    If gbLokalModus = True Then
        SSCommand1(5).Enabled = False
    Else
        If gbBargeldEingabe = True Then
            SSCommand1(5).Enabled = False
        Else
            SSCommand1(5).Enabled = True
        End If
    End If
    
    If frmWKL00.Command3(1).Caption = "Z Bon" Then
        SSCommand1(0).Enabled = False
    Else
        SSCommand1(0).Enabled = True
    End If
                
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
