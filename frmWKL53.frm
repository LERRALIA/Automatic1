VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL53 
   Caption         =   "Winkiss Programmeinstellungen"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL53.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   11565
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.Frame fraWebshop 
      Caption         =   "Webshop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   464
      Tag             =   "1"
      Top             =   8280
      Width           =   1215
      Begin VB.Frame Frame20 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   487
         Tag             =   "1"
         Top             =   3360
         Width           =   2175
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   120
            MaxLength       =   4
            TabIndex        =   491
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   120
            MaxLength       =   6
            TabIndex        =   488
            Top             =   480
            Width           =   1215
         End
         Begin sevCommand3.Command Command39 
            Height          =   255
            Left            =   1080
            TabIndex        =   489
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
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
               Weight          =   400
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
            Caption         =   "Test"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "neuer Bestand"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   492
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Artikelnummer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   490
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   120
         MaxLength       =   250
         TabIndex        =   485
         Top             =   840
         Width           =   8775
      End
      Begin VB.CheckBox Check75 
         Caption         =   "Shop-Artikel anlegen"
         Enabled         =   0   'False
         Height          =   240
         Left            =   3960
         TabIndex        =   474
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   120
         MaxLength       =   20
         TabIndex        =   470
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   120
         MaxLength       =   20
         TabIndex        =   469
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   120
         MaxLength       =   50
         TabIndex        =   468
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox Check71 
         Caption         =   "Live-Bestandsführung"
         Height          =   240
         Left            =   120
         TabIndex        =   465
         Top             =   1320
         Width           =   3015
      End
      Begin sevCommand3.Command Command40 
         Height          =   255
         Left            =   9120
         TabIndex        =   493
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
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
            Weight          =   400
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
         Caption         =   "Script erstellen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Pfad zu den PHP-Scripten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   486
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Name der Bestandsspalte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   473
         Top             =   2760
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Name der Artikelnummernspalte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   472
         Top             =   2160
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Tabellenname der Bestandstabelle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   471
         Top             =   1560
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lbl6 
         Caption         =   "MySQL Verbindungseinstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   94
         Left            =   120
         TabIndex        =   467
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lbl6 
         Caption         =   $"frmWKL53.frx":0442
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   93
         Left            =   120
         TabIndex        =   466
         Top             =   4920
         Width           =   2775
      End
   End
   Begin VB.Frame fraKisslive 
      Caption         =   "KissLive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   434
      Tag             =   "1"
      Top             =   5880
      Width           =   1815
      Begin VB.CheckBox Check86 
         Caption         =   "Live-Nachrichten"
         Height          =   240
         Left            =   3360
         TabIndex        =   524
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CheckBox Check38 
         Caption         =   "Live-'nicht geführt'=gesperrt"
         Height          =   360
         Left            =   7320
         TabIndex        =   517
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   514
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox Check46 
         Caption         =   "Live-Farbänderung"
         Height          =   240
         Left            =   7320
         TabIndex        =   513
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Live-Gutschein"
         Height          =   240
         Left            =   5520
         TabIndex        =   509
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CheckBox Check48 
         Caption         =   "bei Livebestand nur Diffmenge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   499
         Top             =   3720
         Width           =   2775
      End
      Begin VB.CheckBox Check26 
         Caption         =   "Live-KVK Preisänderung"
         Height          =   240
         Left            =   2760
         TabIndex        =   497
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   120
         MaxLength       =   20
         TabIndex        =   445
         Top             =   1680
         Width           =   5175
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   442
         Top             =   3120
         Width           =   5175
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         MaxLength       =   50
         TabIndex        =   440
         Top             =   2400
         Width           =   5175
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   436
         Top             =   960
         Width           =   3735
      End
      Begin VB.CheckBox Check29 
         Caption         =   "Live-Bestandsführung"
         Height          =   240
         Left            =   120
         TabIndex        =   435
         Top             =   3480
         Width           =   2535
      End
      Begin sevCommand3.Command Command37 
         Height          =   255
         Left            =   5400
         TabIndex        =   444
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
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
            Weight          =   400
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
         Caption         =   "Verbindungstest"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "DSN (optional)"
         Height          =   255
         Index           =   16
         Left            =   3960
         TabIndex        =   515
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Datenbank"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   446
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Passwort"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   443
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Benutzer"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   441
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Adresse "
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   439
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lbl6 
         Caption         =   "Erläuterung"
         Height          =   495
         Index           =   90
         Left            =   120
         TabIndex        =   438
         Top             =   4080
         Width           =   5175
      End
      Begin VB.Label lbl6 
         Caption         =   "SQL-Server Verbindungseinstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   89
         Left            =   120
         TabIndex        =   437
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame fraUnter 
      Caption         =   "Voreinstellungen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   44
      Tag             =   "1"
      Top             =   9600
      Width           =   4815
      Begin VB.Frame Frame7 
         Height          =   855
         Left            =   9480
         TabIndex        =   288
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   293
            Top             =   5160
            Width           =   855
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   292
            ToolTipText     =   "Beispiel ausprobieren"
            Top             =   5160
            Width           =   1455
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "proberunden"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   6
            Left            =   10440
            TabIndex        =   290
            Top             =   240
            Width           =   255
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "x"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label lbl6 
            Caption         =   "Probieren Sie es einfach aus. Geben Sie unten die ungerundete Zahl ein."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   240
            TabIndex        =   294
            Top             =   4800
            Width           =   6135
         End
         Begin VB.Label lbl6 
            Caption         =   "Variante"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Index           =   70
            Left            =   120
            TabIndex        =   291
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lbl6 
            Caption         =   "Variante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   289
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   281
         Top             =   5400
         Width           =   375
      End
      Begin VB.CheckBox Check82 
         Caption         =   "alte Stammdaten automatisch löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   280
         ToolTipText     =   "nach 7 Wochen"
         Top             =   5640
         Width           =   3375
      End
      Begin VB.Frame Frame5 
         Height          =   960
         Left            =   120
         TabIndex        =   82
         Tag             =   "2"
         Top             =   220
         Width           =   3470
         Begin VB.CheckBox Check28 
            Caption         =   "automatisches Synchronisieren"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   388
            Top             =   480
            Width           =   3255
         End
         Begin VB.CheckBox Check5 
            Caption         =   "auto LM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2160
            TabIndex        =   101
            Top             =   180
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   """local Security"" "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   83
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1170
         Left            =   120
         TabIndex        =   155
         Top             =   1080
         Width           =   3470
         Begin VB.CheckBox Check11 
            Caption         =   "an allen Computern"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   156
            Top             =   440
            Width           =   2655
         End
         Begin VB.CheckBox Check23 
            Caption         =   "schnelle Passwortanmeldung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   228
            Top             =   680
            Width           =   3255
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Anmeldung mit Bedienerkarte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   157
            Top             =   200
            Width           =   3015
         End
         Begin VB.CheckBox Check73 
            Caption         =   "mit Begrüßungsbon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   274
            Top             =   920
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   3120
         Width           =   3470
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'Kein
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Index           =   6
            Left            =   1320
            TabIndex        =   460
            Top             =   720
            Visible         =   0   'False
            Width           =   2025
            Begin VB.OptionButton opt1 
               Caption         =   "kleinsten"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   23
               Left            =   960
               TabIndex        =   462
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton opt1 
               Caption         =   "größten"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   22
               Left            =   0
               TabIndex        =   461
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.Label lbl2 
               Caption         =   "bei mehreren ausgehend vom"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   200
               Index           =   31
               Left            =   120
               TabIndex        =   463
               Top             =   80
               Width           =   1695
            End
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Listen- EK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Schnitt-EK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lbl6 
            Caption         =   "Nettospannen/ KVK Preise Berechnungsgrundlage:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   2
         Left            =   8280
         TabIndex        =   49
         ToolTipText     =   "Beispiele anzeigen"
         Top             =   240
         Width           =   2655
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   9
            Left            =   1320
            TabIndex        =   343
            Top             =   1800
            Width           =   255
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "B"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Variante 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   120
            TabIndex        =   344
            Top             =   1800
            Width           =   1215
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   287
            Top             =   1560
            Width           =   255
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "B"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   286
            Top             =   1320
            Width           =   255
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "B"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Variante 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   285
            Top             =   1560
            Width           =   1455
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Variante 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   120
            TabIndex        =   284
            Top             =   1320
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox Check80 
            Alignment       =   1  'Rechts ausgerichtet
            Caption         =   "spezielle Rundungsregeln"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   279
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   55
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   53
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   50
            Top             =   240
            Width           =   375
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   510
            Top             =   2040
            Width           =   255
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "B"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Variante 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   511
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Rechts
            Caption         =   "Rundungskriterium:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   56
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Rechts
            Caption         =   "Abrunden auf:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   54
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Rechts
            Caption         =   "Aufrunden auf:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl6 
            Caption         =   "Runden:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.CheckBox Check76 
         Caption         =   "GROßBUCHSTABEN (Artikelbez)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   276
         Top             =   5160
         Width           =   3375
      End
      Begin VB.CheckBox Check64 
         Caption         =   "alte Arbeitszeitauswertung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   272
         Top             =   4440
         Width           =   2775
      End
      Begin VB.CheckBox Check32 
         Caption         =   "mit Bildschirmtastatur"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   235
         Top             =   4680
         Width           =   2775
      End
      Begin sevCommand3.Command Command1 
         Height          =   255
         Index           =   7
         Left            =   9840
         TabIndex        =   310
         Top             =   3840
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Index           =   8
         Left            =   3600
         TabIndex        =   127
         Top             =   240
         Width           =   4660
         Begin VB.CheckBox Check53 
            Caption         =   "neue Nr vorschlagen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2640
            TabIndex        =   505
            Top             =   680
            Width           =   1935
         End
         Begin VB.CheckBox Check66 
            Caption         =   "eindeutige Suche -> Einzelmaske"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   406
            Top             =   900
            Width           =   2775
         End
         Begin VB.CheckBox Check65 
            Caption         =   "neue Artikel anlegen: erlaubt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   384
            Top             =   680
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   19
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   329
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   9
            Left            =   120
            MaxLength       =   9
            TabIndex        =   128
            ToolTipText     =   "Fil=0 im Bereich 500000 und 599999"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lbl6 
            Caption         =   "Standard MWST"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   78
            Left            =   1680
            TabIndex        =   330
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lbl6 
            Caption         =   "neue Artnr ab:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   129
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Index           =   4
         Left            =   3600
         TabIndex        =   78
         Top             =   1320
         Width           =   4660
         Begin VB.OptionButton opt1 
            Caption         =   "CipherLab"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   26
            Left            =   1680
            TabIndex        =   494
            Top             =   840
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Casio-Mde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   1680
            TabIndex        =   459
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Bela-Mde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   458
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Rewe-Mde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   391
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   264
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cbocom 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmWKL53.frx":04CA
            Left            =   2160
            List            =   "frmWKL53.frx":04CC
            TabIndex        =   158
            Text            =   "cbocom"
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Scanpal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   80
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Forcom"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbl2 
            Caption         =   "Pause"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1440
            TabIndex        =   266
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lbl2 
            Caption         =   "Com P"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   2160
            TabIndex        =   265
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lbl6 
            Caption         =   "MDE-Gerät"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   81
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Index           =   13
         Left            =   3600
         TabIndex        =   311
         Top             =   2655
         Width           =   4660
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   8
            Left            =   2880
            TabIndex        =   321
            Top             =   2160
            Width           =   1695
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "LUG's löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   318
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   317
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton opt1 
            Caption         =   "alle Artikel, die geführt werden(geführt='J')"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   20
            Left            =   120
            TabIndex        =   316
            Top             =   360
            Width           =   4095
         End
         Begin VB.OptionButton opt1 
            Caption         =   "alle Artikel, die seit "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   19
            Left            =   120
            TabIndex        =   314
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton opt1 
            Caption         =   "alle Artikel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   120
            TabIndex        =   313
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton opt1 
            Caption         =   "alle Artikel, die seit "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   120
            TabIndex        =   312
            Top             =   1440
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.Label lbl2 
            Caption         =   "Tagen im Zugang waren."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   1800
            TabIndex        =   320
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label lbl2 
            Caption         =   "Tagen verkauft worden sind."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   1800
            TabIndex        =   319
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label lbl6 
            Caption         =   "Berechnung Lagerumschlag"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   315
            Top             =   120
            Width           =   4215
         End
      End
      Begin VB.CheckBox Check56 
         Caption         =   "Gesamt EK Wert anzeigen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8280
         TabIndex        =   369
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox Check52 
         Caption         =   "Geburtstagerinnerung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8280
         TabIndex        =   447
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   9600
         MaxLength       =   2
         TabIndex        =   370
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check69 
         Caption         =   "auch Adresse drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8520
         TabIndex        =   448
         Top             =   3600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Rechts
         Caption         =   "Anzahl Tage:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   8160
         TabIndex        =   371
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Rechts
         Caption         =   "Inventureingaben"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   8160
         TabIndex        =   309
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Rechts
         Caption         =   "Stammdatenübernahmepause (in sec)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   282
         Top             =   5400
         Width           =   2775
      End
   End
   Begin VB.Frame fraKasse 
      Caption         =   "Kasse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   261
      Tag             =   "1"
      Top             =   7560
      Width           =   1455
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   500
         Text            =   "0"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   11
         Left            =   120
         TabIndex        =   365
         Top             =   360
         Width           =   2655
         Begin VB.ComboBox Combo4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1800
            Style           =   2  'Dropdown-Liste
            TabIndex        =   367
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox Combo5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown-Liste
            TabIndex        =   366
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbl6 
            Caption         =   "Welche Waage?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   368
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.CheckBox Check109 
         Caption         =   "Terminpreise bonusfähig"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3120
         TabIndex        =   349
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox Check109 
         Caption         =   "Neukunden anzeigen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3120
         TabIndex        =   345
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox Check108 
         Caption         =   "Sterne anzeigen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   342
         Top             =   1080
         Width           =   2535
      End
      Begin sevCommand3.Command Command26 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   331
         Top             =   3000
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "TSE - Einstellungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command26 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   295
         Top             =   5400
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "Einstellungen an der Kasse"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "alle ""geparkten Artikel"" mit dieser Farbe versehen(Farbnr eintragen)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   3120
         TabIndex        =   501
         Top             =   1680
         Width           =   2775
      End
   End
   Begin VB.Frame fraTagesabschluss 
      Caption         =   "Tagesabschluss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   239
      Tag             =   "1"
      Top             =   7200
      Width           =   2415
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   519
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7800
         TabIndex        =   393
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton opt1 
            Caption         =   "Listendrucker"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   240
            TabIndex        =   395
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Bondrucker"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   394
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbl6 
            Caption         =   "Zählbeleg drucken"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   396
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.CheckBox Check62 
         Caption         =   "Protokoll ""Artikel kumuliert"" mit Summen drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   377
         Top             =   5640
         Width           =   5055
      End
      Begin VB.CheckBox Check57 
         Caption         =   "mit Export (Z-Bon)"
         Height          =   315
         Left            =   240
         TabIndex        =   376
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox Check93 
         Caption         =   "Protokoll ""Artikel kumuliert"" sortiert nach WGN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   326
         Top             =   5400
         Width           =   4815
      End
      Begin VB.CheckBox Check83 
         Caption         =   "AGN Zusammenfassung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   297
         Top             =   5160
         Width           =   4095
      End
      Begin VB.CheckBox Check79 
         Caption         =   "an diesem Rechner kein Abschluss"
         Height          =   225
         Left            =   6000
         TabIndex        =   277
         Top             =   2160
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   6000
         TabIndex        =   259
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110886914
         CurrentDate     =   38457.8333333333
      End
      Begin VB.CheckBox Check47 
         Caption         =   "Kassendateien sofort versenden"
         Height          =   225
         Left            =   6000
         TabIndex        =   258
         Top             =   2520
         Width           =   3855
      End
      Begin sevCommand3.Command Command21 
         Height          =   360
         Left            =   6720
         TabIndex        =   257
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   400
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
         Caption         =   "Achtung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   252
         Top             =   1680
         Width           =   5415
         Begin VB.CheckBox Check41 
            Caption         =   "ohne Warengruppen"
            Height          =   360
            Left            =   2760
            TabIndex        =   518
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox Check72 
            Caption         =   "Filialtäusche"
            Height          =   360
            Left            =   360
            TabIndex        =   273
            Top             =   1440
            Width           =   3375
         End
         Begin VB.CheckBox Check45 
            Caption         =   "Ein- und Auszahlungen"
            Height          =   360
            Left            =   360
            TabIndex        =   256
            Top             =   1080
            Width           =   3495
         End
         Begin VB.CheckBox Check44 
            Caption         =   "Kreditkartenzahlungen"
            Height          =   360
            Left            =   360
            TabIndex        =   255
            Top             =   720
            Width           =   3495
         End
         Begin VB.CheckBox Check43 
            Caption         =   "Artikel kumuliert "
            Height          =   360
            Left            =   360
            TabIndex        =   254
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lbl6 
            Caption         =   "Reporte beim schnellen Abschluss"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   253
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.CheckBox Check39 
         Caption         =   "EC Lastschriften an die Zentrale"
         Height          =   240
         Left            =   6000
         TabIndex        =   250
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox Check25 
         Caption         =   "schneller Abschluss"
         Height          =   360
         Left            =   240
         TabIndex        =   249
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6000
         TabIndex        =   245
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox chk_ZBON_DINA4_HOCH 
            Caption         =   "Hochformat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   521
            Top             =   610
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Bondrucker"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   247
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Listendrucker"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   246
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lbl6 
            Caption         =   "Z-Bon drucken"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   248
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Abschlussnummer"
         Height          =   240
         Left            =   7680
         TabIndex        =   243
         Top             =   3720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Abschlussdatum"
         Height          =   240
         Left            =   7680
         TabIndex        =   242
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin sevCommand3.Command Command19 
         Height          =   240
         Left            =   6000
         TabIndex        =   241
         Top             =   3960
         Width           =   495
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   400
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
         Caption         =   "P"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Abschluss mit Bargeldeingabe"
         Height          =   240
         Left            =   240
         TabIndex        =   240
         Top             =   480
         Width           =   3375
      End
      Begin sevCommand3.Command Command98 
         Height          =   360
         Left            =   3480
         TabIndex        =   433
         Top             =   4800
         Width           =   405
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
         ToolTipTitle    =   "Spaltenanordung"
         ButtonStyle     =   2
         Caption         =   ""
         Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
         Picture         =   "frmWKL53.frx":04CE
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "Z-Bon an diese Emailadresse senden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   520
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lbl6 
         Caption         =   "Z-Bon zusammenstellen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   72
         Left            =   240
         TabIndex        =   296
         Top             =   4800
         Width           =   3015
      End
      Begin VB.Label lbl6 
         Caption         =   "Die Kassendateien versenden um:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   6000
         TabIndex        =   260
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lbl6 
         Caption         =   "Z-Bon:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   240
         TabIndex        =   244
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraNacht 
      Caption         =   "Nachtverarbeitung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   205
      Tag             =   "1"
      Top             =   8760
      Width           =   1815
      Begin VB.CheckBox Check104 
         Caption         =   "nur Winkiss ausschalten, Rechner bleibt an"
         Height          =   240
         Left            =   240
         TabIndex        =   333
         Top             =   1800
         Width           =   5415
      End
      Begin VB.CheckBox Check92 
         Caption         =   "Externe Sicherung anlegen"
         Height          =   240
         Left            =   240
         TabIndex        =   325
         Top             =   4560
         Width           =   3975
      End
      Begin VB.CheckBox Check90 
         Caption         =   "Mindestbestände errechnen"
         Height          =   240
         Left            =   5760
         TabIndex        =   322
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox Check84 
         Caption         =   "Artikelumsätze neu "
         Height          =   240
         Left            =   2520
         TabIndex        =   298
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox Checkbox1 
         Caption         =   "Übertragung von Programmupdates"
         Height          =   240
         Left            =   240
         TabIndex        =   207
         Top             =   4200
         Width           =   3975
      End
      Begin VB.CheckBox Checkbox3 
         Caption         =   "Übertragung von Statistikdateien"
         Height          =   240
         Left            =   240
         TabIndex        =   208
         Top             =   3840
         Width           =   4095
      End
      Begin VB.CheckBox Checkbox6 
         Caption         =   "Einlesen der Kassendateien"
         Height          =   240
         Left            =   240
         TabIndex        =   211
         Top             =   5520
         Width           =   4455
      End
      Begin VB.CheckBox Checkbox5 
         Caption         =   "Übertragung der Kassendateien"
         Height          =   240
         Left            =   240
         TabIndex        =   212
         Top             =   5280
         Width           =   3975
      End
      Begin VB.CheckBox Checkbox2 
         Caption         =   "Übertragung von Stammdatenänderungen"
         Height          =   240
         Left            =   240
         TabIndex        =   213
         Top             =   3480
         Width           =   4695
      End
      Begin VB.CheckBox Checkbox7 
         Caption         =   "Bestellvorschläge laut Bestellrhythmus bereitstellen"
         Height          =   240
         Left            =   240
         TabIndex        =   214
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox Check68 
         Caption         =   "PC ausschalten nach der Nachtverarbeitung"
         Height          =   240
         Left            =   240
         TabIndex        =   216
         Top             =   1560
         Width           =   5415
      End
      Begin VB.CheckBox Check67 
         Caption         =   "Die Nachtverarbeitung startet um:"
         Height          =   240
         Left            =   240
         TabIndex        =   227
         Top             =   360
         Width           =   5415
      End
      Begin VB.CheckBox Checkbox9 
         Caption         =   "Stammdaten einlesen"
         Height          =   240
         Left            =   5760
         TabIndex        =   232
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox Checkbox8 
         Caption         =   "Lagerumschlagswerte"
         Height          =   240
         Left            =   2520
         TabIndex        =   233
         Top             =   960
         Width           =   2775
      End
      Begin sevCommand3.Command Command18 
         Height          =   375
         Left            =   6720
         TabIndex        =   234
         Top             =   2400
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Protokoll der Nachtverarbeitung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         Left            =   7200
         Style           =   2  'Dropdown-Liste
         TabIndex        =   221
         Top             =   5280
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frmWKL53.frx":0B60
         Left            =   5400
         List            =   "frmWKL53.frx":0B62
         Style           =   2  'Dropdown-Liste
         TabIndex        =   220
         Top             =   5280
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   240
         TabIndex        =   206
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110886914
         CurrentDate     =   38457.8333333333
      End
      Begin VB.Label Label5 
         Caption         =   "Wiederholungs- intervall"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   223
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "1. Versuch Dateien von der Zentrale abzuholen nach:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5400
         TabIndex        =   222
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lbl6 
         Caption         =   "bei Filialgeschäften:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   240
         TabIndex        =   215
         Top             =   5040
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lbl6 
         Caption         =   "Nachtverarbeitung beinhaltet:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   240
         TabIndex        =   210
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label lbl6 
         Caption         =   "Die Nachtverabeitung funktioniert nur wenn dieser PC auch die Datenbank enthält und FTP - Rechner ist."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   240
         TabIndex        =   209
         Top             =   720
         Visible         =   0   'False
         Width           =   8175
      End
   End
   Begin VB.Frame fraNoEURO 
      Caption         =   "Fremdwährung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   182
      Tag             =   "1"
      Top             =   9000
      Width           =   1455
      Begin sevCommand3.Command Command13 
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   192
         Top             =   5160
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   400
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   191
         Top             =   4680
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
            Weight          =   400
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text13 
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   190
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox Text13 
         Height          =   360
         Index           =   1
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   188
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text13 
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   186
         Top             =   1440
         Width           =   2895
      End
      Begin VB.ComboBox cboNoEuro 
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   185
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label lblanz 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   193
         Top             =   5280
         Width           =   7095
      End
      Begin VB.Label lbl2 
         Caption         =   "Umrechnunskurs zum EURO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   189
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label lbl2 
         Caption         =   "Währungskürzel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5160
         TabIndex        =   187
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Caption         =   "Fremdwährungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   184
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Caption         =   "Währungsbezeichnung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   183
         Top             =   1200
         Width           =   2775
      End
   End
   Begin VB.Frame fraDruck 
      Caption         =   "Druckeinstellungen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   43
      Tag             =   "1"
      Top             =   7440
      Width           =   1575
      Begin VB.CheckBox Check85 
         Caption         =   "Etikett bei Farbänderung abstellen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   516
         Top             =   4320
         Width           =   3375
      End
      Begin VB.CheckBox Check31 
         Caption         =   "alle Druckansichten automatisch speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   495
         Top             =   5280
         Width           =   4335
      End
      Begin VB.Frame Frame19 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   6240
         TabIndex        =   478
         Top             =   240
         Width           =   2295
         Begin VB.CheckBox Check77 
            Caption         =   "schneller Scanmodus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   482
            Top             =   1320
            Width           =   1935
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Lieferantenbestnr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   120
            TabIndex        =   480
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton opt1 
            Caption         =   "EAN/Artnr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   120
            TabIndex        =   479
            Top             =   600
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.Label lbl6 
            Caption         =   "Etiketten selbst wählen - Startfokus auf"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   96
            Left            =   120
            TabIndex        =   481
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   26
         Left            =   9840
         MaxLength       =   6
         TabIndex        =   407
         Top             =   5040
         Width           =   975
      End
      Begin sevCommand3.Command Command26 
         Height          =   345
         Index           =   2
         Left            =   8640
         TabIndex        =   385
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Einstellungen Kassenbon"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check74 
         Caption         =   "auf dem Bondrucker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   275
         Top             =   4800
         Width           =   2415
      End
      Begin sevCommand3.Command Command20 
         Height          =   345
         Left            =   8640
         TabIndex        =   251
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "letztes Differenzprotokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check37 
         Caption         =   "bei Filialtausch mit EK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   238
         Top             =   5040
         Width           =   2775
      End
      Begin VB.CheckBox Check36 
         Caption         =   "Preisänderungsprotokoll drucken (Einlesen der Zentraldatei aus der Nachtverarbeitung)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   237
         Top             =   4560
         Width           =   7815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   9840
         MaxLength       =   4
         TabIndex        =   230
         Top             =   4320
         Width           =   975
      End
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   3840
         TabIndex        =   151
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton opt1 
            Caption         =   "zum Drucker leiten"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   120
            TabIndex        =   153
            Top             =   600
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton opt1 
            Caption         =   "ignorieren"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   152
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lbl6 
            Caption         =   "Info über erfolgreiche Dateiübertragungen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   41
            Left            =   120
            TabIndex        =   154
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   3615
         Begin VB.CheckBox Check54 
            Caption         =   "EAN statt Libesnr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   522
            Top             =   1920
            Width           =   1935
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Originalreihenfolge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   392
            Top             =   1560
            Width           =   1695
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Lief.-Nr, Linie, Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   100
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Lief-Nr, Lief.-Bestell-Nr, Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   99
            Top             =   840
            Width           =   3255
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Lief.-Nr, Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   98
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   97
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lbl6 
            Caption         =   "Voreinstellung für Etiketten"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   96
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lbl6 
            Caption         =   "- Sortierung nach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   95
            Top             =   360
            Width           =   1575
         End
      End
      Begin sevCommand3.Command Command26 
         Height          =   345
         Index           =   3
         Left            =   8640
         TabIndex        =   523
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Datenschutzblatt"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         Caption         =   "Edeka Etikett"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   88
         Left            =   8520
         TabIndex        =   409
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label lbl6 
         Caption         =   "LiefNr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   82
         Left            =   9840
         TabIndex        =   408
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         Caption         =   "Tabellenfaktor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   8520
         TabIndex        =   231
         Top             =   4320
         Width           =   1215
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sortierungen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   159
      Tag             =   "1"
      Top             =   8280
      Width           =   1335
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   10
         Left            =   120
         TabIndex        =   160
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000004&
            Caption         =   "Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   162
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000004&
            Caption         =   "Lieferant, Linie und Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   161
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label lbl6 
            Caption         =   "allg. Artikelsortierung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   164
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lbl6 
            Caption         =   "nach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   163
            Top             =   480
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   17
      Tag             =   "1"
      Top             =   960
      Width           =   8655
      Begin VB.CheckBox Check24 
         Caption         =   "keine Wochendaten übertragen (Parfümerie)"
         Height          =   240
         Left            =   120
         TabIndex        =   507
         Top             =   5040
         Width           =   5415
      End
      Begin VB.CheckBox Check70 
         Caption         =   "Sicherheit bei der Übernahme von Artikeln. Die Option ""Kassenverkaufspreise übernehmen"" standardmäßig nicht ausgewählt"
         Height          =   720
         Left            =   120
         TabIndex        =   506
         Top             =   4320
         Width           =   9495
      End
      Begin VB.CheckBox Check61 
         Caption         =   "kein Hinweis bei neuen Filialtäuschen"
         Height          =   240
         Left            =   120
         TabIndex        =   375
         Top             =   3840
         Width           =   5295
      End
      Begin VB.CheckBox Check34 
         Caption         =   "kein Hinweis bei neuen Stammdaten und Terminpreisen"
         Height          =   240
         Left            =   120
         TabIndex        =   236
         Top             =   3480
         Width           =   5295
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   2400
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Update holen"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdNeuheit 
         Height          =   375
         Left            =   3000
         TabIndex        =   42
         Top             =   2880
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Was ist neu?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin sevCommand3.Command cmdUpdEinlesen 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Update einlesen"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtUpdatepfad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   5175
      End
      Begin sevCommand3.Command cmdStandardUp 
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   360
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdUpdate 
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zum Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lbl6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   6615
      End
   End
   Begin VB.Frame fraSta 
      Caption         =   "Statistik"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   140
      Tag             =   "1"
      Top             =   8040
      Width           =   1575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   25
         Left            =   8040
         MaxLength       =   9
         TabIndex        =   405
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CheckBox Check49 
         Caption         =   "tägliche Abverkäufe an Vedes"
         Height          =   375
         Left            =   7080
         TabIndex        =   403
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox Text19 
         Height          =   360
         Left            =   240
         MaxLength       =   100
         TabIndex        =   356
         Top             =   5040
         Width           =   5895
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Index           =   14
         Left            =   6960
         TabIndex        =   334
         Top             =   240
         Width           =   3975
         Begin sevCommand3.Command Command34 
            Height          =   375
            Left            =   2640
            TabIndex        =   397
            Top             =   2520
            Width           =   1215
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "11/12/13"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   24
            Left            =   120
            MaxLength       =   100
            TabIndex        =   390
            Top             =   2160
            Width           =   3735
         End
         Begin VB.CheckBox Check35 
            Caption         =   "automatisch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   389
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   340
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   20
            Left            =   960
            MaxLength       =   4
            TabIndex        =   339
            Top             =   1320
            Width           =   615
         End
         Begin sevCommand3.Command Command30 
            Height          =   375
            Left            =   120
            TabIndex        =   338
            Top             =   2520
            Width           =   2415
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "erstellen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   21
            Left            =   120
            MaxLength       =   2
            TabIndex        =   335
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lbl6 
            Caption         =   "für Kalenderwoche"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   504
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Rechts
            Caption         =   "Kunden nummer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   341
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lbl6 
            Caption         =   "GFK Woche"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   80
            Left            =   120
            TabIndex        =   337
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl6 
            Caption         =   "Auswertung für Kalenderwoche:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   79
            Left            =   120
            TabIndex        =   336
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   12
         Left            =   120
         TabIndex        =   299
         Top             =   2400
         Width           =   4815
         Begin sevCommand3.Command Command22 
            Height          =   375
            Left            =   3720
            TabIndex        =   302
            Top             =   720
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "sofort"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command27 
            Height          =   375
            Left            =   3720
            TabIndex        =   308
            Top             =   240
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "1. Mal"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.CheckBox Check87 
            Caption         =   "Dieses Unternehmen nimmt an der statistischen Monatsauswertung teil."
            Height          =   495
            Left            =   120
            TabIndex        =   303
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox Text16 
            Height          =   360
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   301
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text15 
            Height          =   360
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   300
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   307
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lbl6 
            Caption         =   "Monatsauswertung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   74
            Left            =   120
            TabIndex        =   306
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Rechts
            Caption         =   "letzte Auswertung:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   305
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbl6 
            Alignment       =   1  'Rechts
            Caption         =   "Kunden nummer:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   73
            Left            =   120
            TabIndex        =   304
            Top             =   1080
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   9
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   4815
         Begin VB.CheckBox Check59 
            Caption         =   "per Email "
            Height          =   375
            Left            =   1800
            TabIndex        =   374
            Top             =   1560
            Width           =   2775
         End
         Begin VB.TextBox Text20 
            Height          =   360
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   363
            Top             =   1560
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text10 
            Height          =   360
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   148
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Height          =   360
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   146
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin sevCommand3.Command cmd12 
            Height          =   375
            Left            =   3720
            TabIndex        =   145
            Top             =   600
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "sofort"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Dieses Unternehmen nimmt an der statistischen Wochenauswertung teil."
            Height          =   495
            Left            =   120
            TabIndex        =   144
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lbl6 
            Alignment       =   1  'Rechts
            Caption         =   "Zusatz:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   86
            Left            =   120
            TabIndex        =   364
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl6 
            Alignment       =   1  'Rechts
            Caption         =   "Kunden nummer:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   39
            Left            =   120
            TabIndex        =   149
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            Caption         =   "letzte Auswertung:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   147
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lbl6 
            Caption         =   "Wochenauswertung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   143
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lbl6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   142
            Top             =   480
            Width           =   1695
         End
      End
      Begin sevCommand3.Command Command36 
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   475
         Top             =   4200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
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
            Weight          =   400
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
         Caption         =   "erstellen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command36 
         Height          =   360
         Index           =   0
         Left            =   7080
         TabIndex        =   476
         ToolTipText     =   "Kalender"
         Top             =   4080
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ToolTip         =   "Wählen Sie hier das Datum aus."
         ToolTipTitle    =   "Kalender"
         ButtonStyle     =   2
         Caption         =   ""
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command36 
         Height          =   255
         Index           =   2
         Left            =   8400
         TabIndex        =   503
         Top             =   4200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
            Weight          =   400
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
         Caption         =   "FTP"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "Datum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   95
         Left            =   7680
         TabIndex        =   477
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lbl6 
         Caption         =   "Vedes-KdNr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   7080
         TabIndex        =   404
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lbl6 
         Caption         =   "feste Pfadangabe + Dateiname für die Excelausgabe der Kundenanalyse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   85
         Left            =   240
         TabIndex        =   355
         Top             =   4800
         Width           =   6615
      End
   End
   Begin VB.Frame fraECASH 
      Caption         =   "elektronische Zahlungsabwicklung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   130
      Tag             =   "1"
      Top             =   7560
      Width           =   1335
      Begin VB.Frame fraELP 
         Caption         =   "elPAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3120
         TabIndex        =   380
         Top             =   3600
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox Text22 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   496
            Top             =   360
            Width           =   2175
         End
         Begin sevCommand3.Command Command33 
            Height          =   495
            Left            =   240
            TabIndex        =   381
            Top             =   720
            Width           =   2175
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "zu den Einstellungen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin VB.ComboBox cboECASH 
         Height          =   360
         Left            =   3960
         Style           =   2  'Dropdown-Liste
         TabIndex        =   133
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox Check6 
         Caption         =   "diesen Computer für E-Cash zulassen"
         Height          =   240
         Left            =   120
         TabIndex        =   132
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame fraadt 
         Caption         =   "ADT Wellcom GmbH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   6480
         TabIndex        =   131
         Top             =   840
         Visible         =   0   'False
         Width           =   4455
         Begin sevCommand3.Command Command23 
            Height          =   495
            Left            =   120
            TabIndex        =   267
            Top             =   3360
            Width           =   2175
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "Zahlung testen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Frame Frame15 
            Caption         =   "XML"
            Height          =   2655
            Left            =   120
            TabIndex        =   169
            Top             =   600
            Width           =   8655
            Begin VB.TextBox Text18 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   360
               MaxLength       =   15
               TabIndex        =   352
               Top             =   2160
               Width           =   1455
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1920
               MaxLength       =   6
               TabIndex        =   351
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox Text14 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   270
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   360
               MaxLength       =   3
               TabIndex        =   268
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox Check20 
               Caption         =   "EuroCard"
               Height          =   255
               Left            =   6360
               TabIndex        =   180
               Top             =   1440
               Width           =   1455
            End
            Begin VB.CheckBox Check19 
               Caption         =   "American Express"
               Height          =   255
               Left            =   6360
               TabIndex        =   179
               Top             =   1200
               Width           =   2055
            End
            Begin VB.CheckBox Check18 
               Caption         =   "VisaCard"
               Height          =   255
               Left            =   6360
               TabIndex        =   178
               Top             =   960
               Width           =   1335
            End
            Begin VB.CheckBox Check17 
               Caption         =   "Diners"
               Height          =   255
               Left            =   6360
               TabIndex        =   177
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox Text12 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   360
               MaxLength       =   8
               TabIndex        =   170
               Top             =   600
               Width           =   5175
            End
            Begin VB.Label lbl6 
               Caption         =   "IP Adresse"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   84
               Left            =   360
               TabIndex        =   354
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label lbl6 
               Caption         =   "Port"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   62
               Left            =   1920
               TabIndex        =   353
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label lbl6 
               Caption         =   "Limit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   67
               Left            =   1920
               TabIndex        =   271
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lbl6 
               Caption         =   "Client - ID"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   65
               Left            =   360
               TabIndex        =   269
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lbl2 
               Caption         =   "zugelassene Karten"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   5760
               TabIndex        =   181
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lbl6 
               Caption         =   "Terminal - ID"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   46
               Left            =   360
               TabIndex        =   171
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.OptionButton Option2 
            Caption         =   "XML - Verfahren"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   168
            Top             =   360
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.Label lbl6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   36
            Left            =   240
            TabIndex        =   135
            Top             =   2520
            Width           =   3015
         End
      End
      Begin VB.Label lbl2 
         Caption         =   "Vertragspartner:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   134
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fraWE 
      Caption         =   "WE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   121
      Tag             =   "1"
      Top             =   8040
      Width           =   1095
      Begin VB.CheckBox Check50 
         Caption         =   "bei Wareneingang aus Bestellung: generell keine Etiketten abstellen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   502
         Top             =   3480
         Width           =   6375
      End
      Begin VB.CheckBox Check33 
         Caption         =   "bei Preisänderungen: Etiketten nur für diese Filiale abstellen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   498
         Top             =   3240
         Width           =   6375
      End
      Begin VB.CheckBox Check63 
         Caption         =   "automatisch zwischenspeichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   379
         Top             =   3000
         Width           =   3015
      End
      Begin VB.CheckBox Check51 
         Caption         =   "automatisch geführt = ""J"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   378
         Top             =   2760
         Width           =   3135
      End
      Begin VB.CheckBox Check15 
         Caption         =   "schneller Scan - Modus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   167
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   17
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   166
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   7
         Left            =   120
         TabIndex        =   122
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox Check21 
            Caption         =   "keine negativen Zugänge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   386
            Top             =   1440
            Width           =   2775
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000004&
            Caption         =   "Lieferscheinnummer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   125
            Top             =   1080
            Width           =   3015
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000004&
            Caption         =   "EAN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   124
            Top             =   720
            Value           =   -1  'True
            Width           =   3015
         End
         Begin VB.Label lbl6 
            Caption         =   "Startfokus auf: "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   126
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lbl6 
            Caption         =   "aus Einzellieferung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   123
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Label lbl6 
         Caption         =   "Voreinstellung für Zu/Abgang:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   120
         TabIndex        =   165
         Top             =   2160
         Width           =   2895
      End
   End
   Begin VB.Frame fraPfade 
      Caption         =   "Pfade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Tag             =   "1"
      Top             =   7560
      Width           =   855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   28
         Left            =   240
         TabIndex        =   453
         Top             =   4320
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   27
         Left            =   240
         TabIndex        =   452
         Top             =   5040
         Width           =   5175
      End
      Begin sevCommand3.Command Command4 
         Height          =   350
         Index           =   3
         Left            =   3120
         TabIndex        =   40
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   350
         Index           =   3
         Left            =   4320
         TabIndex        =   39
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   2280
         Width           =   5175
      End
      Begin sevCommand3.Command Command3 
         Height          =   350
         Index           =   2
         Left            =   4320
         TabIndex        =   35
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   350
         Index           =   2
         Left            =   3120
         TabIndex        =   34
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   5175
      End
      Begin sevCommand3.Command Command3 
         Height          =   350
         Index           =   1
         Left            =   4320
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   350
         Index           =   1
         Left            =   3120
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   350
         Index           =   0
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   350
         Index           =   0
         Left            =   4320
         TabIndex        =   27
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "LITE!.LZH, Pfad zum Programmupdate "
         Top             =   600
         Width           =   5175
      End
      Begin sevCommand3.Command Command4 
         Height          =   345
         Index           =   4
         Left            =   3120
         TabIndex        =   450
         Top             =   4680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   345
         Index           =   4
         Left            =   4320
         TabIndex        =   451
         Top             =   4680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   345
         Index           =   5
         Left            =   4320
         TabIndex        =   454
         Top             =   3960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   345
         Index           =   5
         Left            =   3120
         TabIndex        =   455
         Top             =   3960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zur Webcam-Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   92
         Left            =   240
         TabIndex        =   457
         ToolTipText     =   "Y-Dateien aus der Zentrale"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zu den Fotos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   91
         Left            =   240
         TabIndex        =   456
         ToolTipText     =   "F-Dateierstellung an die Zentrale"
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zu den Ausgangsdateien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   41
         ToolTipText     =   "F-Dateierstellung an die Zentrale"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zu den Kassendateien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   37
         ToolTipText     =   "Y-Dateien aus der Zentrale"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zum Wareneingang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "WV*.dbf , Pfad zu den Warenverteilungen"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfad zum Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame fraDaba 
      Caption         =   "Datenbank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   102
      Tag             =   "1"
      Top             =   7680
      Width           =   975
      Begin VB.CheckBox Check42 
         Caption         =   "ohne Auswertungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   512
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CheckBox Check27 
         Caption         =   "ohne Anzeige"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   387
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox Check105 
         Caption         =   """Penner"" schwarz färben"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   383
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   347
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   346
         Top             =   840
         Width           =   495
      End
      Begin sevCommand3.Command Test 
         Height          =   240
         Left            =   9480
         TabIndex        =   116
         ToolTipText     =   "Reparatur mit Jetcomp"
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Reparatur"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check103 
         Caption         =   "an diesem Rechner keine automatische Komprimierung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   332
         Top             =   2880
         Width           =   5535
      End
      Begin sevCommand3.Command Command28 
         Height          =   255
         Left            =   9480
         TabIndex        =   324
         Top             =   5520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "extern abholen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command25 
         Height          =   255
         Left            =   240
         TabIndex        =   283
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Speicher anzeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command24 
         Height          =   255
         Left            =   9480
         TabIndex        =   278
         Top             =   5160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "extern sichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command17 
         Height          =   255
         Left            =   9480
         TabIndex        =   229
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Datenbankablauf"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Left            =   9480
         TabIndex        =   219
         Top             =   4440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Programmablauf"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command16 
         Height          =   255
         Left            =   9480
         TabIndex        =   218
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Abmeldefehler"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command15 
         Height          =   255
         Left            =   9480
         TabIndex        =   217
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Programmfehler"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdReindizieren 
         Height          =   255
         Left            =   9480
         TabIndex        =   204
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Reindizieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   255
         Left            =   9480
         TabIndex        =   203
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Protokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   10515
         TabIndex        =   202
         Top             =   2280
         Width           =   10575
      End
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   195
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check16 
         Caption         =   "automatische Datenbankkomprimierung immer bei Programmstart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   175
         Top             =   3120
         Width           =   5535
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2640
         TabIndex        =   118
         Top             =   480
         Width           =   975
      End
      Begin sevCommand3.Command cmdKompNow 
         Height          =   255
         Left            =   9480
         TabIndex        =   117
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "Komprimieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   113
         Top             =   4560
         Width           =   975
      End
      Begin sevCommand3.Command cmdNow 
         Height          =   255
         Left            =   2040
         TabIndex        =   111
         Top             =   4560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
            Weight          =   400
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
         Caption         =   "sofort"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmWKL53.frx":0B64
         Left            =   240
         List            =   "frmWKL53.frx":0B7D
         Style           =   2  'Dropdown-Liste
         TabIndex        =   109
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   107
         Text            =   "1000"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   106
         Text            =   "100"
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lbl6 
         Caption         =   "Anzahl:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   87
         Left            =   240
         TabIndex        =   382
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lbl6 
         Caption         =   "Pause:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   83
         Left            =   240
         TabIndex        =   348
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbl6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   240
         TabIndex        =   201
         Top             =   2040
         Width           =   8895
      End
      Begin VB.Label lbl6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   200
         Top             =   1800
         Width           =   8895
      End
      Begin VB.Label lbl6 
         Caption         =   "Datenbank komprimieren (auch reindizieren und sichern)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   240
         TabIndex        =   199
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Label lbl6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   8040
         TabIndex        =   198
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   66
         Left            =   8040
         TabIndex        =   197
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl6 
         Caption         =   "vorher  in MB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   6480
         TabIndex        =   196
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   194
         Top             =   5520
         Width           =   9015
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   240
         X2              =   10800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lbl6 
         Caption         =   "letzte Komprimierung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   68
         Left            =   240
         TabIndex        =   120
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lbl6 
         Caption         =   "jetzt      in MB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   6480
         TabIndex        =   119
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl6 
         Caption         =   "Datenbankpfad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   1560
         TabIndex        =   115
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lbl6 
         Caption         =   " Uhr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   5520
         TabIndex        =   114
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label lbl6 
         Caption         =   "letzte Aktualisierung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   4560
         TabIndex        =   112
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label lbl6 
         Caption         =   " min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   1560
         TabIndex        =   110
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lbl6 
         Caption         =   "Aktualisierung der Datenbank alle:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   108
         Top             =   4320
         Width           =   4095
      End
      Begin VB.Label lbl6 
         Caption         =   "Abstand in ms "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   105
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lbl6 
         Caption         =   "Aktualisierungsversuche"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   104
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label lbl6 
         Caption         =   "Datenbankpfad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   103
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraSicher 
      Caption         =   "Sicherung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   84
      Tag             =   "1"
      Top             =   9120
      Width           =   1455
      Begin VB.Frame Frame10 
         Height          =   4215
         Left            =   120
         TabIndex        =   86
         Top             =   720
         Visible         =   0   'False
         Width           =   6135
         Begin VB.OptionButton Option9 
            Caption         =   "während der Nachtverarbeitung"
            Height          =   240
            Index           =   2
            Left            =   240
            TabIndex        =   402
            Top             =   2400
            Value           =   -1  'True
            Width           =   3015
         End
         Begin VB.OptionButton Option9 
            Caption         =   "täglich um:"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   401
            Top             =   2760
            Width           =   1575
         End
         Begin VB.OptionButton Option9 
            Caption         =   "täglich bei Programmstart"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   400
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   92
            Top             =   1800
            Width           =   1695
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Left            =   3720
            TabIndex        =   91
            Top             =   3480
            Width           =   1695
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "Test Sichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox txtSicherPfad 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   89
            Top             =   840
            Width           =   5175
         End
         Begin sevCommand3.Command Command8 
            Height          =   375
            Left            =   4320
            TabIndex        =   88
            Top             =   360
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "Standard"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command9 
            Height          =   375
            Left            =   3120
            TabIndex        =   87
            Top             =   360
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
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
               Weight          =   400
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
            Caption         =   "Ändern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1920
            TabIndex        =   399
            Top             =   2760
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   110886914
            CurrentDate     =   38457.8333333333
         End
         Begin VB.Label lbl6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   240
            TabIndex        =   150
            Top             =   1200
            Width           =   3495
         End
         Begin VB.Label lbl6 
            Caption         =   "letzte Sicherung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   3720
            TabIndex        =   93
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lbl6 
            Caption         =   "Pfad zur Sicherung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   90
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Dieser PC führt die Datenbanksicherung durch."
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   360
         Width           =   5055
      End
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   3720
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFarben 
      Caption         =   "Design"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Tag             =   "1"
      Top             =   8760
      Width           =   975
      Begin VB.CheckBox Check81 
         Caption         =   "als Demo - Programm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   508
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CheckBox Check60 
         Caption         =   "keine Sprüche beim Starten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   484
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CheckBox Check78 
         Caption         =   "Sounds abspielen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   483
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Frame Frame17 
         Caption         =   "Schaltflächen-Design"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   3360
         TabIndex        =   414
         Top             =   240
         Width           =   4335
         Begin sevCommand3.Command cmdBeispiel 
            Height          =   375
            Left            =   2880
            TabIndex        =   415
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16777215
            BackColorTo     =   12632256
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   14737632
            BackColorHoverTo=   8421504
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   255
            ForeColor       =   4210752
            ForeColorDisabled=   9609633
            ForeColorHover  =   0
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            MousePointer    =   99
            BorderStyle     =   2
            ButtonStyle     =   2
            Caption         =   "Beispiel"
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   5
            Left            =   2760
            TabIndex        =   424
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   6
            Left            =   2760
            TabIndex        =   425
            Top             =   600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   7
            Left            =   2760
            TabIndex        =   426
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   8
            Left            =   2760
            TabIndex        =   427
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   9
            Left            =   2760
            TabIndex        =   428
            Top             =   1320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   10
            Left            =   2760
            TabIndex        =   429
            Top             =   1560
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   11
            Left            =   2760
            TabIndex        =   430
            Top             =   1800
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmd2 
            Height          =   135
            Index           =   12
            Left            =   2760
            TabIndex        =   431
            Top             =   2040
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
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
               Weight          =   400
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
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdStandard 
            Height          =   375
            Left            =   240
            TabIndex        =   432
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16777215
            BackColorTo     =   12632256
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   14737632
            BackColorHoverTo=   8421504
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   255
            ForeColor       =   4210752
            ForeColorDisabled=   9609633
            ForeColorHover  =   0
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            MousePointer    =   99
            BorderStyle     =   2
            ButtonStyle     =   2
            Caption         =   "Beispiel"
            Version3        =   -1  'True
         End
         Begin VB.Label lbl2 
            Caption         =   "Schriftfarbe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   423
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lbl2 
            Caption         =   "Mouseover Schriftfarbe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   422
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label lbl2 
            Caption         =   "Mouseover Rahmenfarbe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   421
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label lbl2 
            Caption         =   "Mouseover Hintergrund zu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   420
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label lbl2 
            Caption         =   "Mouseover Hintergrund von"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   419
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lbl2 
            Caption         =   "Rahmenfarbe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   418
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lbl2 
            Caption         =   "Hintergrund zu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   417
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lbl2 
            Caption         =   "Hintergrund von"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   416
            Top             =   600
            Width           =   1335
         End
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   412
         Top             =   2520
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   21
         Left            =   1920
         TabIndex        =   410
         Top             =   1800
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   7800
         MaxLength       =   20
         TabIndex        =   327
         Top             =   600
         Width           =   2655
      End
      Begin sevCommand3.Command Command12 
         Height          =   255
         Left            =   2160
         TabIndex        =   172
         Top             =   3120
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd5 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
            Weight          =   400
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
         ButtonType      =   93
         Caption         =   "Wiederherstellen"
         Image           =   6421
         UseDefaultMaskColor=   -1  'True
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmd2 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label lbl2 
         Caption         =   "Warnschrift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   413
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl2 
         Caption         =   "Link (Mouse over)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   1920
         TabIndex        =   411
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl6 
         Caption         =   "Programm - Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   77
         Left            =   7800
         TabIndex        =   328
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         Caption         =   "Arial"
         Height          =   255
         Index           =   49
         Left            =   2760
         TabIndex        =   176
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lbl6 
         Caption         =   "Arial"
         Height          =   255
         Index           =   48
         Left            =   240
         TabIndex        =   174
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lbl6 
         Caption         =   "Schriftart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   240
         TabIndex        =   173
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lbl2 
         Caption         =   "Eingabefeld"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Caption         =   "2. Hintergrund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl5 
         Caption         =   "Standard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Caption         =   "Schrift 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Caption         =   "Hintergrund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Caption         =   "Überschrift 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame fraVerbindung 
      Caption         =   "Verbindungseinstellungen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   1320
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2520
      Width           =   9855
      Begin VB.CheckBox Check40 
         Caption         =   "optimierte Stammdatenpflege (Drogerie, Spielwaren)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   398
         Top             =   360
         Width           =   4455
      End
      Begin VB.CheckBox Check58 
         Caption         =   "auto Artikelexport"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   373
         Top             =   4920
         Width           =   2055
      End
      Begin sevCommand3.Command Command32 
         Height          =   255
         Left            =   8520
         TabIndex        =   372
         Top             =   5175
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check55 
         Caption         =   "Texte an Überwachungssoftware senden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8520
         TabIndex        =   362
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Frame Frame18 
         Height          =   2535
         Left            =   8520
         TabIndex        =   357
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox Text2 
            Height          =   360
            Index           =   9
            Left            =   120
            MaxLength       =   6
            TabIndex        =   359
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Index           =   8
            Left            =   120
            MaxLength       =   15
            TabIndex        =   358
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Port"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   361
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "IP"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   360
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CheckBox Check110 
         Caption         =   "Passive Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   350
         Top             =   4920
         Width           =   2775
      End
      Begin VB.CheckBox Check91 
         Caption         =   "optimierte Stammdatenpflege (Parfümerie)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   323
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
      End
      Begin sevCommand3.Command Command10 
         Height          =   375
         Left            =   8520
         TabIndex        =   226
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "Verbinden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   375
         Left            =   9600
         TabIndex        =   225
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
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
            Weight          =   400
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
         Caption         =   "Trennen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo6 
         Height          =   360
         Left            =   9360
         Style           =   2  'Dropdown-Liste
         TabIndex        =   224
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check9 
         Caption         =   "ständige Internetverbindung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox Check8 
         Caption         =   "automatisch "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   138
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Kassendateien über FTP an die Zentrale übertragen (wenn die Zentrale räumlich getrennt, dann Haken setzen!)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   136
         Top             =   1200
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Frame Frame4 
         Height          =   3015
         Left            =   4320
         TabIndex        =   71
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CheckBox Check88 
            Caption         =   "keine Warenverteilungen abholen (übernimmt KissLive)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   525
            Top             =   2520
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   3240
            MaxLength       =   3
            TabIndex        =   262
            Top             =   2130
            Width           =   615
         End
         Begin sevCommand3.Command Command11 
            Height          =   375
            Left            =   120
            TabIndex        =   137
            Top             =   2040
            Width           =   1335
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "sofort"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Index           =   5
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Index           =   4
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox Text2 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   72
            Top             =   1680
            Width           =   3735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            Caption         =   "Alias Nr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   263
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "FTP - Adresse (Zentrale)"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Benutzername"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Passwort"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dieser PC führt die FTP - Übertragung durch."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   4095
         Begin sevCommand3.Command cmdFtpCheckNow 
            Height          =   375
            Left            =   2280
            TabIndex        =   69
            Top             =   2880
            Width           =   1455
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
               Weight          =   400
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
            Caption         =   "sofort"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option1 
            Caption         =   "alle 7 Tage"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   67
            Top             =   3000
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "alle 3 Tage"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   66
            Top             =   2760
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "täglich"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   2520
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   61
            Top             =   1800
            Width           =   3615
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   1200
            Width           =   3615
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label2 
            Caption         =   "auto FTP - Verbindung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Passwort"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Benutzername"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   63
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "FTP - Adresse "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Was ist das?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   6720
         MouseIcon       =   "frmWKL53.frx":0B9C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   449
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip tabWK 
      Height          =   6855
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12091
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   18
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Verbindung"
            Key             =   "Verbindung"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tagesabschluss"
            Key             =   "Tagesabschluss"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Nachtverarbeitung"
            Key             =   "Nachtverarbeitung"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistik"
            Key             =   "Statistik"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fremdwährung"
            Key             =   "Fremdwährung"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "E-Cash"
            Key             =   "ECASH"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "WE"
            Key             =   "WE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sicherung"
            Key             =   "Sicherung"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Design"
            Key             =   "Farben"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            Key             =   "Update"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pfade"
            Key             =   "Pfade"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Druckeinstellungen"
            Key             =   "Druckeinstellungen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Voreinstellungen"
            Key             =   "Unternehmen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datenbank"
            Key             =   "Datenbank"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sortierungen"
            Key             =   "Sortierungen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Kasse"
            Key             =   "Kasse"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "KissLive"
            Key             =   "KissLive"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Webshop"
            Key             =   "Webshop"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   1
      Top             =   7080
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Weight          =   400
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
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Weight          =   400
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
      Caption         =   "Übernehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
End
Attribute VB_Name = "frmWKL53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sdfu As String
Private Sub cboECASH_Click()
    On Error GoTo LOKAL_ERROR
    
    Select Case cboECASH.Text
        Case Is = "ADT Wellcom GmbH"
            fraadt.Visible = True
            fraELP.Visible = False
            
        Case Is = "elPAY"
            fraELP.Caption = "elPAY"
            fraELP.Visible = True
            fraadt.Visible = False
            
        Case Is = "ZV2"
            fraELP.Caption = "ZV2"
            fraadt.Visible = False
            fraELP.Visible = True
            
        Case Is = "ZVT"
            fraELP.Caption = "ZVT"
            fraadt.Visible = False
            fraELP.Visible = True
            
        
        Case Else
            fraadt.Visible = False
            fraELP.Visible = False
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboECASH_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub einschalt()
On Error GoTo LOKAL_ERROR

        byteSortReihen = 2
   
        DTPicker1.Visible = True
        Check104.Visible = True
        Check68.Visible = True
        Check90.Visible = True
        Check92.Visible = True
        Checkbox7.Visible = True
        Checkbox8.Visible = True
        Check84.Visible = True
        Checkbox9.Visible = True
        lbl6(55).Visible = True
        Command18.Visible = True
        
        If gbFtpYes Then
            Checkbox1.Visible = True
            Checkbox2.Visible = True
            Checkbox3.Visible = True
            
            If gbFtpZENT Then
                lbl6(56).Visible = True
                Checkbox5.Visible = True
                Checkbox6.Visible = True
                
                Label4.Visible = True
                Label5.Visible = True
                Combo2.Visible = True
                Combo3.Visible = True
                
                Combo3.Clear
                Combo2.Clear
                
                Combo3.AddItem "2"
                Combo3.AddItem "3"
                Combo3.AddItem "4"
                Combo3.AddItem "5"
                Combo3.AddItem "10"
                Combo3.AddItem "15"
                Combo3.AddItem "20"
                Combo3.AddItem "25"
                Combo3.AddItem "30"
                
                Combo2.AddItem "2"
                Combo2.AddItem "20"
                Combo2.AddItem "30"
                Combo2.AddItem "40"
                Combo2.AddItem "50"
                Combo2.AddItem "60"
                Combo2.AddItem "120"
                Combo2.AddItem "150"
                Combo2.AddItem "180"
                
                If giSTARTMIN = 5 Or giSTARTMIN = 10 Or giSTARTMIN = 15 Or giSTARTMIN = 20 Or giSTARTMIN = 25 Or giSTARTMIN = 30 _
                Or giSTARTMIN = 120 Or giSTARTMIN = 150 Or giSTARTMIN = 180 Then
                    Combo2.Text = giSTARTMIN
                Else
                    Combo2.Text = "2"
                End If
                
                If giINTERV = 20 Or giINTERV = 30 Or giINTERV = 40 _
                Or giINTERV = 2 Or giINTERV = 3 Or giINTERV = 4 Or giINTERV = 5 Or giINTERV = 10 Or giINTERV = 15 Then
                    Combo3.Text = giINTERV
                Else
                    Combo3.Text = "5"
                End If
            Else
                Label4.Visible = False
                Label5.Visible = False
                Combo2.Visible = False
                Combo3.Visible = False
                
                Combo3.Clear
                Combo2.Clear
                
                Combo3.AddItem "2"
                Combo3.AddItem "3"
                Combo3.AddItem "4"
                Combo3.AddItem "5"
                Combo3.AddItem "10"
                Combo3.AddItem "15"
                Combo3.AddItem "20"
                Combo3.AddItem "25"
                Combo3.AddItem "30"
                
                Combo2.AddItem "2"
                Combo2.AddItem "20"
                Combo2.AddItem "30"
                Combo2.AddItem "40"
                Combo2.AddItem "50"
                Combo2.AddItem "60"
                Combo2.AddItem "120"
                Combo2.AddItem "150"
                Combo2.AddItem "180"
                
                Combo2.Text = "2"
                Combo3.Text = "5"
                
            End If
        Else
            Label4.Visible = False
            Label5.Visible = False
            Combo2.Visible = False
            Combo3.Visible = False
            
            Combo3.AddItem "5"
            Combo3.AddItem "10"
            Combo3.AddItem "15"
            Combo3.AddItem "20"
            Combo3.AddItem "25"
            Combo3.AddItem "30"
            
            Combo2.AddItem "2"
            Combo2.AddItem "20"
            Combo2.AddItem "30"
            Combo2.AddItem "40"
            Combo2.AddItem "50"
            Combo2.AddItem "60"
            
            Combo2.Text = "2"
            Combo3.Text = "5"
        End If
        
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "einschalt"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

    Fehlermeldung1
End Sub
Private Sub ausschalt()
On Error GoTo LOKAL_ERROR


        byteSortReihen = 1
        DTPicker1.value = "20:00:00"
        DTPicker1.Visible = False
        Check68.Visible = False
        Check104.Visible = False
        Check68.value = vbUnchecked
        Check90.Visible = False
        Check92.Visible = False
        Checkbox7.Visible = False
        Checkbox8.Visible = False
        Check84.Visible = False
        Checkbox9.Visible = False
        Command18.Visible = False
        lbl6(55).Visible = False
        
        Checkbox1.Visible = False
        Checkbox2.Visible = False
        Checkbox3.Visible = False
        
        lbl6(56).Visible = False
        Checkbox5.Visible = False
        Checkbox6.Visible = False
        
        Label4.Visible = False
        Label5.Visible = False
        Combo2.Visible = False
        Combo3.Visible = False
    

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ausschalt"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

    Fehlermeldung1
End Sub



Private Sub Check23_Click()
On Error GoTo LOKAL_ERROR

    If Check23.value = vbChecked Then
        Check12.value = vbUnchecked
        Check11.value = vbUnchecked
    Else
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check23_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Check25_Click()
On Error GoTo LOKAL_ERROR

    If Check25.value = vbChecked Then
        Frame13.Visible = True
        Check57.Visible = True

        If Check57.value = vbChecked Then
            lbl6(8).Visible = True
            Text8.Visible = True
        Else
            lbl6(8).Visible = False
            Text8.Visible = False
        End If
        
        
    Else
        Frame13.Visible = False
        Check57.Visible = False

        lbl6(8).Visible = False
        Text8.Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check25_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Check29_Click()
On Error GoTo LOKAL_ERROR
    
    If Check29.value = vbChecked Then
        Check48.Enabled = True
    Else
        Check48.Enabled = False
        Check48.value = vbUnchecked
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check29_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Check35_Click()
On Error GoTo LOKAL_ERROR
    
    If Check35.value = vbChecked Then
    
        Check35.Caption = "automatisch an diesen FTP User"
        Text1(24).Visible = True
        Text1(24).Text = ""
        Text1(24).SetFocus
    Else
        Check35.Caption = "automatisch"
        
        Text1(24).Visible = False
        Text1(24).Text = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check35_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check36_Click()
On Error GoTo LOKAL_ERROR
    
    If Check36.value = vbChecked Then
        Check74.Visible = True
        Check74.value = vbUnchecked
        
        
    ElseIf Check36.value = vbUnchecked Then
        Check74.Visible = False
        Check74.value = vbUnchecked
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check36_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check47_Click()
On Error GoTo LOKAL_ERROR
    
    If Check47.value = vbChecked Then
        lbl6(59).Visible = False
        DTPicker2.Visible = False
        
    Else
        lbl6(59).Visible = True
        DTPicker2.Visible = True
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check47_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check49_Click()
On Error GoTo LOKAL_ERROR

    If Check49.value = vbChecked Then
        Text1(25).Visible = True
        lbl6(31).Visible = True
        
    Else
        Text1(25).Visible = False
        lbl6(31).Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check49_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Check52_Click()
On Error GoTo LOKAL_ERROR

    If Check52.value = vbChecked Then
        Text1(23).Visible = True
        lbl2(20).Visible = True
        Check69.Visible = True
    Else
        Text1(23).Visible = False
        lbl2(20).Visible = False
        Check69.Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check52_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Check55_Click()
On Error GoTo LOKAL_ERROR

    If Check55.value = vbChecked Then
        Frame18.Visible = True
    Else
        Frame18.Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check55_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Check57_Click()
On Error GoTo LOKAL_ERROR
    
    If Check57.value = vbChecked Then
        lbl6(8).Visible = True
        Text8.Visible = True
        
    Else
        lbl6(8).Visible = False
        Text8.Visible = False
        
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check57_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check58_Click()
On Error GoTo LOKAL_ERROR
    
    If Check58.value = vbChecked Then
        speicherAuto_Export_Artikelbestand
        Command32.Visible = True
    Else
        speicherAuto_Export_Artikelbestand
        Command32.Visible = False
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check58_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check71_Click()
On Error GoTo LOKAL_ERROR
    
    If Check71.value = vbChecked Then
        Label1(18).Visible = True
        Label1(19).Visible = True
        Label1(20).Visible = True
        Text21(8).Visible = True
        Text21(9).Visible = True
        Text21(10).Visible = True
        Frame20.Visible = True
    Else
        Label1(18).Visible = False
        Label1(19).Visible = False
        Label1(20).Visible = False
        Text21(8).Visible = False
        Text21(9).Visible = False
        Text21(10).Visible = False
        Frame20.Visible = False
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check71_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check80_Click()
On Error GoTo LOKAL_ERROR
    
    If Check80.value = vbChecked Then
        opt1(15).Visible = True
        opt1(16).Visible = True
        opt1(21).Visible = True
        opt1(4).Visible = True
        Command1(3).Visible = True
        Command1(5).Visible = True
        Command1(9).Visible = True
        Command1(0).Visible = True
    Else
        opt1(15).Visible = False
        opt1(16).Visible = False
        opt1(21).Visible = False
        opt1(4).Visible = False
        Command1(3).Visible = False
        Command1(5).Visible = False
        Command1(9).Visible = False
        Command1(0).Visible = False
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check80_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check88_Click()
On Error GoTo LOKAL_ERROR
    
    If Check88.value = vbChecked Then
        gbWVNOT = True
    ElseIf Check88.value = vbUnchecked Then
        gbWVNOT = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check88_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Check9_Click()
On Error GoTo LOKAL_ERROR
    
    If Check9.value = vbChecked Then
        Check91.Visible = True
        Check91.value = vbUnchecked
        Check40.Visible = True
        Check40.value = vbUnchecked
    Else
        Check40.Visible = False
        Check40.value = vbUnchecked
        Check91.Visible = False
        Check91.value = vbUnchecked
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check9_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Command10_Click()

    
    Dim iRet As Long
    
    iRet = 1
    ConID = 5
    
    If Combo6.Text = "keine DFÜ vorhanden" Then
        MsgBox "keine DFÜ vorhanden", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    Else
        sdfu = Trim$(Combo6.Text)
        iRet = InternetDial(Me.hwnd, sdfu, &H2000, ConID, 0)
        
        If iRet = 0 Then
            Command7.Visible = True
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command17_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "DABAABL.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command17_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Command18_Click()
On Error GoTo LOKAL_ERROR

'    Dim cPfad       As String
'    Dim cpfadZiel   As String
'
'
'    cPfad = gcDBPfad
'    If Right$(cPfad, 1) <> "\" Then
'        cPfad = cPfad & "\"
'    End If
'
'    cpfadZiel = cPfad & "ENDZIPIN\"
'
'    Zip_Unzip "", cPfad, cpfadZiel & "end.lzh", txtStatus
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "NACHTV.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command18_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command19_Click()
On Error GoTo LOKAL_ERROR

    

    Screen.MousePointer = 11
    zeigeHilfeDabapfad "ABPRO", "ABPROTO.txt"
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command19_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command20_Click()
On Error GoTo LOKAL_ERROR

Dim sSQL As String

If NewTableSuchenDBKombi("diffta", gdBase) Then

    loeschNEW "DIFFDRUCK", gdBase
    
    sSQL = "Select * into DIFFDRUCK from diffta order by bezeich"
    gdBase.Execute sSQL, dbFailOnError
    
    If Not SpalteInTabellegefundenNEW("DIFFDRUCK", "liefbez", gdBase) Then
        SpalteAnfuegenNEW "DIFFDRUCK", "liefbez", "Text(35)", gdBase
    
        sSQL = "Update DIFFDRUCK inner join LISRT on DIFFDRUCK.linr = lisrt.linr "
        sSQL = sSQL & " set DIFFDRUCK.liefbez = LISRT.liefbez "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("DIFFDRUCK", "EAN", gdBase) Then
        SpalteAnfuegenNEW "DIFFDRUCK", "EAN", "Text(13)", gdBase
    
        sSQL = "Update DIFFDRUCK inner join Artikel on DIFFDRUCK.artnr = Artikel.artnr "
        sSQL = sSQL & " set DIFFDRUCK.EAN = Artikel.ean "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Update DIFFDRUCK set BESTSYS = 0 where BESTSYS is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DIFFDRUCK set LWEKSYS = 0 where LWEKSYS is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DIFFDRUCK set diffbest = 0 where diffbest is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DIFFDRUCK set diffLWEK = 0 where diffLWEK is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from DIFFDRUCK where bestheut = bestsys "
    gdBase.Execute sSQL, dbFailOnError
    
    reportbildschirm "", "aWKL46ds" 'Differenz
    
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command20_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command22_Click()
On Error GoTo LOKAL_ERROR
    
    'kisskundennr vorhanden
    If Trim(gsMStatkundnr) = "" Then
        MsgBox "Geben Sie erst eine Kiss - Kundennummer ein und drücken Sie dann übernehmen! Danach drücken Sie auf sofort!", vbInformation, "Winkiss Hinweis:"
        Text15.SetFocus
        Exit Sub
    End If
    
    If unistatMonat Then
        
        Text16.Text = DatumLastSuniM
        
        If gbFtpYes Then
            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
            frmWKL38.Show 1
        Else
            gsAnzeigeText = "Die statistischen Daten sind erstellt. Bitte übertragen Sie diese."
            frmWK21l.Show 1
        End If
    End If

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command22_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Command23_Click()
    On Error GoTo LOKAL_ERROR
    
    frmWKL117.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command24_Click()
On Error GoTo LOKAL_ERROR

ExternSichern txtStatus, lbl6(28)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command24_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command25_Click()
On Error GoTo LOKAL_ERROR

    Dim strText As String
    Dim udtMem As MEMORYSTATUS

    GlobalMemoryStatus udtMem

    With udtMem
        strText = strText & "Speicher belegt in Prozent : " & _
            Format$(.dwMemoryLoad, "@@@@@@@@@@@") & vbCrLf & vbCrLf

        strText = strText & "Totaler physischer Speicher : " & _
            Format$(.dwTotalPhys, "@@@@@@@@@@@") & vbCrLf

        strText = strText & "Davon noch frei : " & _
            Format$(.dwAvailPhys, "@@@@@@@@@@@") & vbCrLf & vbCrLf

        strText = strText & "Bytes in gepageten Dateien :  " & _
            Format$(.dwTotalPageFile, "@@@@@@@@@@@") & vbCrLf

        strText = strText & "Davon noch frei : " & _
            Format$(.dwAvailPageFile, "@@@@@@@@@@@") & vbCrLf & vbCrLf

        strText = strText & "Totaler virtueller Speicher : " & _
            Format$(.dwTotalVirtual, "@@@@@@@@@@@") & vbCrLf

        strText = strText & "Davon noch frei : " & _
            Format$(.dwAvailVirtual, "@@@@@@@@@@@") & vbCrLf
    End With

    MsgBox strText
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command25_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command26_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case index
        Case 0
            frmWKL131.Show 1
            If gbEDITKASSNR = False Then
                frmWKL00.Command1(9).Enabled = False
            Else
                frmWKL00.Command1(9).Enabled = True
            End If
        Case 1
'            frmWKL146.Show 1
            'frmWKL218.Show 1
            TseEinstellungen.Left = (Me.ScaleWidth - TseEinstellungen.Width) / 2
            TseEinstellungen.Top = (Me.ScaleHeight - TseEinstellungen.Height) / 2
            TseEinstellungen.Show 1
        Case 2
            frmWKL179.Show 1
            
        Case 3
            frmWKL215.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command26_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command27_Click()
On Error GoTo LOKAL_ERROR
    
    'kisskundennr vorhanden
    If Trim(gsMStatkundnr) = "" Then
        MsgBox "Geben Sie erst eine Kiss - Kundennummer ein und drücken Sie dann übernehmen! Danach drücken Sie auf sofort!", vbInformation, "Winkiss Hinweis:"
        Text15.SetFocus
        Exit Sub
    End If
    
    If unistatMonat1Mal(txtStatus, picprogress) Then
        Text16.Text = DatumLastSuniM
        
        If gbFtpYes Then
            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
            frmWKL38.Show 1
        Else
            gsAnzeigeText = "Die statistischen Daten sind erstellt. Bitte übertragen Sie diese."
            frmWK21l.Show 1
        End If
    End If

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command22_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Command28_Click()
On Error GoTo LOKAL_ERROR

ExternAbholenDABA lbl6(28), txtStatus, lbl6(53)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command28_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command29_Click()
On Error GoTo LOKAL_ERROR

SpeicherCaption frmWKL147
SpeicherCaption frmWKL148
SpeicherCaption frmWKL149
SpeicherCaption frmWKL150
SpeicherCaption frmWKL151
SpeicherCaption frmWKL152
SpeicherCaption frmWKL153
SpeicherCaption frmWKLao

SpeicherCaption frmWKL145
SpeicherCaption frmWKL146
SpeicherCaption frmWKL141
SpeicherCaption frmWKL142
SpeicherCaption frmWKL143
SpeicherCaption frmWKL144
SpeicherCaption frmWKL127
SpeicherCaption frmWKL128
SpeicherCaption frmWKL129
SpeicherCaption frmWKL130
SpeicherCaption frmWKL131
SpeicherCaption frmWKL132
SpeicherCaption frmWKL133
SpeicherCaption frmWKL134
SpeicherCaption frmWKL135
SpeicherCaption frmWKL136
SpeicherCaption frmWKL137
SpeicherCaption frmWKL138
SpeicherCaption frmWKL139
SpeicherCaption frmWKL140
SpeicherCaption frmWKL112
SpeicherCaption frmWKL113
SpeicherCaption frmWKL114
SpeicherCaption frmWKL115
SpeicherCaption frmWKL117
SpeicherCaption frmWKLah
SpeicherCaption frmWKL118
SpeicherCaption frmWKL119
SpeicherCaption frmWKL120
SpeicherCaption frmWKL121
SpeicherCaption frmWKL122
SpeicherCaption frmWKL123
SpeicherCaption frmWKL124
SpeicherCaption frmWKL125
SpeicherCaption frmWKL126
SpeicherCaption frmWKL65
SpeicherCaption frmWKL66
SpeicherCaption frmWKL67
SpeicherCaption frmWKL68
SpeicherCaption frmWKL69
SpeicherCaption frmWKL70
SpeicherCaption frmWK25i
SpeicherCaption frmWK25j
SpeicherCaption frmWK25k
SpeicherCaption frmWKL71
SpeicherCaption frmWKL72
SpeicherCaption frmWKL73
SpeicherCaption frmWKL74
SpeicherCaption frmWKL75
SpeicherCaption frmWKL76
SpeicherCaption frmWKL77
SpeicherCaption frmWKL78
SpeicherCaption frmWKL79
SpeicherCaption frmWKL80
SpeicherCaption frmWKL84
SpeicherCaption frmWKL85
SpeicherCaption frmWKL86
SpeicherCaption frmWKL87
SpeicherCaption frmWKL88
SpeicherCaption frmWKL89
SpeicherCaption frmWKL91
SpeicherCaption frmWKL92
SpeicherCaption frmWKL93
SpeicherCaption frmWKL100
SpeicherCaption frmWKL110
SpeicherCaption frmWKL101
SpeicherCaption frmWKL102
SpeicherCaption frmWKL103
SpeicherCaption frmWKL111
SpeicherCaption frmWKL62
SpeicherCaption frmWKL00
SpeicherCaption frmWKL10
SpeicherCaption frmWKL12
SpeicherCaption frmWKL13
SpeicherCaption frmWKL18
SpeicherCaption frmWKL19
SpeicherCaption frmWKL20
SpeicherCaption frmWKL40
SpeicherCaption frmWKL41
SpeicherCaption frmWKL42
SpeicherCaption frmWKL21
SpeicherCaption frmWKL22
SpeicherCaption frmWKL23
SpeicherCaption frmWKL52
SpeicherCaption frmWKL15
SpeicherCaption frmWKL29
SpeicherCaption frmWKL11
SpeicherCaption frmWKL99
SpeicherCaption frmWKL43
SpeicherCaption frmWKL30
SpeicherCaption frmWKL24
SpeicherCaption frmWKL28
SpeicherCaption frmWK25a
SpeicherCaption frmWKL50
SpeicherCaption frmWKL31
SpeicherCaption frmWKL82
SpeicherCaption frmWKL81
SpeicherCaption frmWK25c
SpeicherCaption frmWKL83
SpeicherCaption frmWKL44

SpeicherCaption frmWK21b
SpeicherCaption frmWK00a
SpeicherCaption frmWKL16
SpeicherCaption frmWK81a
SpeicherCaption frmWK81b
SpeicherCaption frmWK81c
SpeicherCaption frmWKL45
SpeicherCaption frmWKL57
SpeicherCaption frmWKL46
SpeicherCaption frmWKL01
SpeicherCaption frmWKL17
SpeicherCaption frmWKL58
SpeicherCaption frmWKL59
SpeicherCaption frmWK20a
SpeicherCaption frmWKLab
SpeicherCaption frmWK20b
SpeicherCaption frmWK21d
SpeicherCaption frmWK15a
SpeicherCaption frmWK20c
SpeicherCaption frmWKL02
SpeicherCaption frmWK40c
SpeicherCaption frmWK25d
SpeicherCaption frmWK24a
SpeicherCaption frmWKLae
SpeicherCaption frmWKLaf
SpeicherCaption frmWKLai
SpeicherCaption frmWKLaj
SpeicherCaption frmWKLak
SpeicherCaption frmWKLal
SpeicherCaption frmWKL00b
SpeicherCaption frmWK24b
SpeicherCaption frmWK20d
SpeicherCaption frmWK25f
SpeicherCaption frmWK25g
SpeicherCaption frmWK24c
SpeicherCaption frmWKLam
SpeicherCaption frmWKLan
SpeicherCaption frmWK25h
SpeicherCaption frmWK10a
SpeicherCaption frmWK21f
SpeicherCaption frmWK20h
SpeicherCaption frmWK20e
SpeicherCaption frmWK20g
SpeicherCaption frmWKL48
SpeicherCaption frmWK20f
SpeicherCaption frmWKL56
SpeicherCaption frmWKLar
SpeicherCaption frmWKLas
SpeicherCaption frmWKLau
SpeicherCaption frmWKLav
SpeicherCaption frmWK21k
SpeicherCaption frmWK21l
SpeicherCaption frmWKL53
SpeicherCaption dlgAbfrage
SpeicherCaption frmWKL25
SpeicherCaption frmWKL26
SpeicherCaption frmWKL27

SpeicherCaption frmWKL33
SpeicherCaption frmWKL34
SpeicherCaption dlgAbfrage3
SpeicherCaption frmWKL35
SpeicherCaption frmWKL36
SpeicherCaption frmWKL03
SpeicherCaption frmWKL38
SpeicherCaption frmWKL37
SpeicherCaption frmWKL09
SpeicherCaption frmWK11a
SpeicherCaption frmWK12a
SpeicherCaption frmWKL54
SpeicherCaption frmWKL47
SpeicherCaption frmWKL49
SpeicherCaption frmWKL06
SpeicherCaption frmWKL07
SpeicherCaption frmWKL156
SpeicherCaption frmWKL14
SpeicherCaption frmWKL05
SpeicherCaption frmWKL39
SpeicherCaption frmWKL51
SpeicherCaption frmWKL55
SpeicherCaption frmWKL60
SpeicherCaption frmWK21m
SpeicherCaption frmWKL61
SpeicherCaption frmWK21n
SpeicherCaption frmWK21o
SpeicherCaption frmWKL63
SpeicherCaption frmWKL64

SpeicherCaption frmWKL32


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command29_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command30_Click()
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer

    If Trim(Text1(21).Text) = "" Then
    
        MsgBox "Bitte geben Sie die Kalenderwoche an, die Sie auswerten möchten!", vbInformation, "Winkiss Hinweis:"
        Text1(21).SetFocus
        Exit Sub
    ElseIf Trim(Text1(21).Text) = "0" Then
        MsgBox "Bitte geben Sie die Kalenderwoche an, die Sie auswerten möchten!", vbInformation, "Winkiss Hinweis:"
        Text1(21).Text = ""
        Text1(21).SetFocus
        Exit Sub
    End If
    
    If Trim(Text1(22).Text) = "" Then
        MsgBox "Bitte geben Sie die Kundennummer an!", vbInformation, "Winkiss Hinweis:"
        Text1(22).SetFocus
        Exit Sub
    ElseIf Trim(Text1(22).Text) = "0" Then
        MsgBox "Bitte geben Sie die Kundennummer an!", vbInformation, "Winkiss Hinweis:"
        Text1(22).Text = ""
        Text1(22).SetFocus
        Exit Sub
    End If
    
    If Check35.value = vbChecked Then
        If Trim(Text1(24).Text) = "" Then
            MsgBox "Bitte geben Sie eine Email-Adresse an!", vbInformation, "Winkiss Hinweis:"
            Text1(24).SetFocus
            Exit Sub
        End If
    End If
    
    While Len(Trim(Text1(21).Text)) < 2
        Text1(21).Text = "0" & Text1(21).Text
    Wend
    
    GFKstat Trim(Text1(22).Text), Check35.value, Trim(Text1(24).Text)
    GFKerstellen Trim(Text1(21).Text), CInt(Text1(20).Text), Trim(Text1(22).Text), Check35.value, Trim(Text1(24).Text)
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command30_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command32_Click()
On Error GoTo LOKAL_ERROR

    frmWKL05.Show 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command32_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command33_Click()
On Error GoTo LOKAL_ERROR
    
    Select Case fraELP.Caption
    
        Case "elPAY"
            frmWKL116.Show 1
        Case "ZVT"
            
            frmWKL198.Show 1
            
        Case "ZV2"
            
            frmWKL210.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command33_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command34_Click()
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    
    If Trim(Text1(22).Text) = "" Then
        MsgBox "Bitte geben Sie die Kundennummer an!", vbInformation, "Winkiss Hinweis:"
        Text1(22).SetFocus
        Exit Sub
    ElseIf Trim(Text1(22).Text) = "0" Then
        MsgBox "Bitte geben Sie die Kundennummer an!", vbInformation, "Winkiss Hinweis:"
        Text1(22).Text = ""
        Text1(22).SetFocus
        Exit Sub
    End If
    
    If Check35.value = vbChecked Then
        If Trim(Text1(24).Text) = "" Then
            MsgBox "Bitte geben Sie eine Email-Adresse an!", vbInformation, "Winkiss Hinweis:"
            Text1(24).SetFocus
            Exit Sub
        End If
    End If
    
    GFKerstellenJahr 2011, Trim(Text1(22).Text), Check35.value, Trim(Text1(24).Text)
    GFKerstellenJahr 2012, Trim(Text1(22).Text), Check35.value, Trim(Text1(24).Text)
    GFKerstellenJahr 2013, Trim(Text1(22).Text), Check35.value, Trim(Text1(24).Text)
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command34_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command36_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case index
        Case 0
            lbl6(95).Caption = Format(Datumschreiben11a(Command36(0).Top, Command36(0).Left), "DD.MM.YY")
        Case 1

            If NewTableSuchenDBKombi("VEDESSTAT", gdBase) Then
                If CBool(leseVEDESstat("live")) = True Then 'Teilnahme an VEDES _Abverkaufszahlen
                    If leseVEDESstat("marktnr") <> "" Then
                    
                        If CLng(DateValue(lbl6(95).Caption)) < CLng(DateValue(Now)) Then
                            VEDES_AUSW_erstellen CLng(DateValue(lbl6(95).Caption))
                            VEDES_AUSW_uebertragen
                        End If
                    End If
                End If
            End If
        Case 2
            frmWKL211.Show 1
    End Select



Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command36_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command37_Click()
On Error GoTo LOKAL_ERROR

    speicherKissLive

    If IfOnline = "Offline" Then
        MsgBox "Es besteht keine Verbindung zum Internet." & vbCrLf & "Bitte stellen Sie eine Online-Verbindung her und versuchen Sie es erneut.", vbCritical, "Winkiss Hinweis:"
        Exit Sub
    End If

    If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
        MsgBox "Die Datenbank wurde erfolgreich geöffnet.", vbInformation + vbOKOnly, "Winkiss Hinweis:"
    Else
        MsgBox "Leider konnte die Datenbank nicht geöffnet werden.", vbCritical + vbOKOnly, "Winkiss Hinweis:"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command37_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command39_Click()
On Error GoTo LOKAL_ERROR

    speicherWebshop

    If IfOnline = "Offline" Then
        MsgBox "Es besteht keine Verbindung zum Internet." & vbCrLf & "Bitte stellen Sie eine Online-Verbindung her und versuchen Sie es erneut.", vbCritical, "Winkiss Hinweis:"
        Exit Sub
    End If

'    If fTestLogin_MySQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
'        MsgBox "Die Datenbank wurde erfolgreich geöffnet.", vbInformation + vbOKOnly, "Winkiss Hinweis:"
'    Else
'        MsgBox "Leider konnte die Datenbank nicht geöffnet werden.", vbCritical + vbOKOnly, "Winkiss Hinweis:"
'    End If
    
    If Text21(5).Text <> "" And Text21(4).Text <> "" Then
        MsgBox OpenURL(gsMySQL_PHP_SCRIPT_PFAD & "/bestand.php?quelle=kiss&tab=" & gsMySQL_BESTAND_TAB & "&spalteb=" & gsMySQL_BESTAND_SPALTE & "&spaltea=" & gsMySQL_BESTAND_INDEXSPALTE & "&artnr=" & Text21(4).Text & "&menge=" & Text21(5).Text & "", IOTDirect)
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command39_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command40_Click()
On Error GoTo LOKAL_ERROR

    frmWKL90.Show 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command40_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1




'On Error GoTo LOKAL_ERROR
'
'    schreibe_Php_Script_Bestand
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Command40_Click"
'    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
End Sub


Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR

    Dim t As Integer
    
    If Trim(sdfu) <> "" Then
        Do Until t = 20
            If DFÜStatus Then 'ist noch online
                Call HangUp(sdfu)
                Pause (1)
                t = t + 1
            Else 'ist getrennt
                t = 20
            End If
        Loop
    End If
    
    If DFÜStatus = False Then
        MsgBox "Die Verbindung wurde getrennt.", vbInformation, "Winkiss Hinweis:"
        Command7.Visible = False
    Else
        MsgBox "Sie sind noch mit dem Internet verbunden. Bitte schließen Sie die Verbindung selbstständig!", vbInformation, "Winkiss Hinweis:"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = ""
    gstab = "ZBON"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    
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
Private Sub cboNoEuro_Click()
On Error GoTo LOKAL_ERROR
    
    Dim sUK     As String
    Dim sWK     As String
    Dim sWB     As String
    Dim cSatz   As String
    
    cSatz = cboNoEuro.Text
   
    If cSatz = "bitte auswählen" Then
        zustandx
    Else
    
        sWK = Mid(cSatz, 1, 6)
        sWK = Trim(sWK)
        Text13(1).Text = sWK
        
        sWB = Mid(cSatz, 19, 30)
        sWB = Trim(sWB)
        Text13(0).Text = sWB
        
        sUK = Mid(cSatz, 11, 10)
        sUK = Trim(sUK)
        Text13(2).Text = sUK
        
    End If
        
        

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboNoEuro_Click"
    Fehler.gsFehlertext = "Beim Anzeigen der Fremdwährungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR
    Dim iFilnr As Integer
    
    If Check1.value = vbChecked Then
        iFilnr = CInt(gcFilNr)
        If iFilnr > 0 Then
            Check7.Visible = True
        End If
        
        If NewTableSuchenDBKombi("StammFTP", gdBase) Then
            LeseStammFtp
            zeigestammftp
        End If
    Else
    
        Text2(2).Text = ""
        Text2(1).Text = ""
        Text2(0).Text = ""
    
        Frame4.Visible = False
        Check7.Visible = False
        Check7.value = vbUnchecked
    
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Check10_Click()
    On Error GoTo LOKAL_ERROR

    If Check10.value = vbUnchecked Then
        cmd12.Visible = False
        Text11.Visible = False
        Label3.Visible = False
        lbl6(39).Visible = False
        lbl6(86).Visible = False
        Text10.Visible = False
        Text20.Visible = False
        Check59.Visible = False
    ElseIf Check10.value = vbChecked Then
        
        LeseStatistWoche
        
        Check10.value = vbChecked
        cmd12.Visible = True
        Label3.Visible = True
        Text11.Text = gdateStatlast
        Text11.Visible = True
        Text10.Text = gsStatkundnr
        Text10.Visible = True
        Text20.Visible = True
        lbl6(39).Visible = True
        lbl6(86).Visible = True
        
        Check59.Visible = True
        If gbDSL Then
            Check59.Enabled = True
        Else
            Check59.Enabled = False
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check10_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Check87_Click()
    On Error GoTo LOKAL_ERROR

    If Check87.value = vbUnchecked Then
        Command22.Visible = False
        Command27.Visible = False
        Text16.Visible = False
        Label9.Visible = False
        lbl6(73).Visible = False
        Text15.Visible = False
    ElseIf Check87.value = vbChecked Then
        
        LeseStatistMonat
        
        Check87.value = vbChecked
        
        Label9.Visible = True
        
        If gdateMStatlast = 0 Then
            Command27.Visible = True
        Else
            Command22.Visible = True
        End If
        
        Text16.Text = gdateMStatlast
        Text16.Visible = True
        Text15.Text = gsMStatkundnr
        Text15.Visible = True
        lbl6(73).Visible = True
        
        
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check87_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Check11_Click()
    On Error GoTo LOKAL_ERROR

    If Check11.value = vbChecked Then
        Check12.value = vbChecked
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check11_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Check12_Click()
On Error GoTo LOKAL_ERROR

    If Check12.value = vbUnchecked Then
        Check11.value = vbUnchecked
        
    Else
        Check23.value = vbUnchecked
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check12_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Check2_Click()
    On Error GoTo LOKAL_ERROR

    If Check2.value = vbUnchecked Then
        Check5.value = vbUnchecked
        Check28.value = vbUnchecked
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check4_Click()
    On Error GoTo LOKAL_ERROR
    Dim iFilnr As Integer
    
    If Check4.value = vbChecked Then
        Frame10.Visible = True
        txtSicherPfad.Text = gsSicherPfad
        Option9(0).value = True
        
    Else
        Frame10.Visible = False
        txtSicherPfad.Text = ""
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check5_Click()
    On Error GoTo LOKAL_ERROR

    If Check5.value = vbChecked Then
        Check2.value = vbChecked
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check5_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check28_Click()
    On Error GoTo LOKAL_ERROR

    If Check28.value = vbChecked Then
        Check2.value = vbChecked
        Check5.value = vbChecked
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check28_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check6_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check6.value = vbChecked Then
        lbl2(9).Visible = True
        combofuell
        cboECASH.Visible = True
        
    ElseIf Check6.value = vbUnchecked Then
        cboECASH.Visible = False
        gsEPartner = ""
        fraadt.Visible = False
        lbl2(9).Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check6_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf."
    
    Fehlermeldung1
End Sub
Private Sub Check7_Click()
    On Error GoTo LOKAL_ERROR
    Dim iFilnr As Integer
    Dim sSQL As String

    If Check7.value = vbChecked Then
        iFilnr = CInt(gcFilNr)
        If iFilnr > 0 Then
            Frame4.Visible = True
        End If
        Option1(0).value = True
        
        If NewTableSuchenDBKombi("StammFTP", gdBase) Then
            sSQL = "Update StammFTP Set ftpoft = 0 "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        LeseStammFtp
        zeigestammftp_OnlyZentrale
'        zeigestammftp

        Option1(1).Enabled = False
        Option1(2).Enabled = False
    Else
        Frame4.Visible = False
        Option1(1).Enabled = True
        Option1(2).Enabled = True
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check7_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
    
End Sub
Private Sub Check8_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check8.value = vbChecked Then
        gbFTPautomatic = True
    ElseIf Check8.value = vbUnchecked Then
        gbFTPautomatic = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check8_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Check67_Click()
    On Error GoTo LOKAL_ERROR

    If byteSortReihen = 1 Then
        einschalt
    Else
        ausschalt
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check67_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

    Fehlermeldung1
End Sub
Private Sub cmd12_Click()
    On Error GoTo LOKAL_ERROR
    
    'kisskundennr vorhanden
    If Trim(gsStatkundnr) = "" Then
        MsgBox "Geben Sie erst eine Kiss - Kundennummer ein und drücken Sie dann übernehmen! Danach drücken Sie auf sofort!", vbInformation, "Winkiss Hinweis:"
        Text10.SetFocus
        Exit Sub
    End If
    
    If unistatweek(txtStatus, picprogress) Then

        Text11.Text = DatumLastSuniW

        If gbFtpYes Then
            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
            frmWKL38.Show 1
        Else
            gsAnzeigeText = "Die statistischen Daten sind erstellt. Bitte übertragen Sie diese."
            frmWK21l.Show 1
        End If
    End If
    
'    If unistatweek_new(txtStatus, picprogress) Then
'
'        Text11.Text = DatumLastSuniW
'
'        If gbFtpYes Then
'            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
'            frmWKL38.Show 1
'        Else
'            gsAnzeigeText = "Die statistischen Daten sind erstellt. Bitte übertragen Sie diese."
'            frmWK21l.Show 1
'        End If
'    End If


Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmd12_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub cmd2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    cdl1.ShowColor
    cmd2(index).BackColor = cdl1.Color
    
    Select Case index
        Case 5
            cmdBeispiel.ForeColor = cdl1.Color
        Case 6
            cmdBeispiel.BackColorFrom = cdl1.Color
        Case 7
            cmdBeispiel.BackColorTo = cdl1.Color
        Case 8
            cmdBeispiel.BorderColor = cdl1.Color
        Case 9
            cmdBeispiel.HoverColorFrom = cdl1.Color
        Case 10
            cmdBeispiel.HoverColorTo = cdl1.Color
        Case 11
            cmdBeispiel.BorderColorHover = cdl1.Color
        Case 12
            cmdBeispiel.ForeColorHover = cdl1.Color
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmd2_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub cmd5_Click()
    On Error GoTo LOKAL_ERROR

    cmd2(0).BackColor = 12777980
    cmd2(1).BackColor = 12566463
    cmd2(2).BackColor = 2039583
    cmd2(3).BackColor = 12615680
    cmd2(4).BackColor = 65280
    cmd2(21).BackColor = 8404992
    cmd2(22).BackColor = 255
    
    cmdBeispiel.BackColorFrom = cmdStandard.BackColorFrom
    cmdBeispiel.BackColorTo = cmdStandard.BackColorTo
    cmdBeispiel.HoverColorFrom = cmdStandard.HoverColorFrom
    cmdBeispiel.HoverColorTo = cmdStandard.HoverColorTo
    cmdBeispiel.BorderColorHover = cmdStandard.BorderColorHover
    cmdBeispiel.BorderColor = cmdStandard.BorderColor
    cmdBeispiel.ForeColorHover = cmdStandard.ForeColorHover
    cmdBeispiel.ForeColor = cmdStandard.ForeColor
    
    cmd2(5).BackColor = cmdBeispiel.ForeColor
    cmd2(6).BackColor = cmdBeispiel.BackColorFrom
    cmd2(7).BackColor = cmdBeispiel.BackColorTo
    cmd2(9).BackColor = cmdBeispiel.HoverColorFrom
    cmd2(10).BackColor = cmdBeispiel.HoverColorTo
    cmd2(12).BackColor = cmdBeispiel.ForeColorHover
    cmd2(8).BackColor = cmdBeispiel.BorderColor
    cmd2(11).BackColor = cmdBeispiel.BorderColorHover
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Cmd5_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub cmdFtpCheckNow_Click()
    On Error GoTo LOKAL_ERROR
    
    FTPprüfung
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdFtpCheckNow_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1

    
End Sub
Private Sub cmdKompNow_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    

    Screen.MousePointer = 11
        
    If BistDualleineinderDatenbank Then
        CompactmyDaba gcDBPfad, "KISSDATA.MDB", gdBase, lbl6(53), txtStatus, lbl6(28), gbOhneAnzeige
        
        sSQL = "update dbeinste set lastkomp='" & Date & "'"
        sSQL = sSQL & " ,lastkomptime='" & TimeValue(Now) & "'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
    End If
    
    lbl6(28).ForeColor = vbBlack
    lbl6(28).Caption = "3 Datenbanken werden noch bearbeitet..."
    lbl6(28).Refresh
    
    lbl6(53).ForeColor = vbBlack
    lbl6(53).Caption = "KISSAPP wird komprimiert"
    lbl6(53).Refresh
    
    If BistDualleineinderDatenbankApp Then
        dbApp_Compri "Kissapp.MDB"
    Else
        anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
    End If
    
    lbl6(53).ForeColor = vbBlack
    lbl6(53).Caption = "GDPDU wird komprimiert"
    lbl6(53).Refresh
    
    If BistDualleineinderDatenbankGDPDU Then
        GDPDU_GLAGER_KLEINHALTEN lbl6(28)
        dbGDPDU_Compri "GDPDU.MDB", lbl6(28)
    Else
        anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
    End If
    
    
    
    lbl6(53).ForeColor = vbBlack
    lbl6(53).Caption = "KASSBON wird komprimiert"
    lbl6(53).Refresh
    
    If BistDualleineinderDatenbankKASSBON Then
        dbKASSBON_Compri "KASSBON.MDB"
    Else
        anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
    End If
    
    
    
    
    
    
    
    Dim sTabc As String
    sTabc = kassetabcheck(gdBase, lbl6(53), lbl6(28))
            
        
    If sTabc = "" Then

    Else
'        MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden.", vbInformation, "Winkiss Hinweis:"
'                End
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    lbl6(53).ForeColor = vbBlack
    lbl6(53).Caption = "Alles Fertig"
    lbl6(53).Refresh

    lbl6(28).ForeColor = vbBlack
    lbl6(28).Caption = "Alles Fertig"
    lbl6(28).Refresh
    
    Text7(0).Text = DatumLastKompAnzeigen
    Text7(2).Text = DatumLastKompZeitAnzeigen
    
    FileSizeAnzeigen
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdKompNow_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub cmdNeuheit_Click()
    On Error GoTo LOKAL_ERROR
    
    URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/neuigkeiten.html"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdNeuheit_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler beim Anzeigen der Neuheit.doc auf. "
    
    Fehlermeldung1
End Sub
Private Sub cmdNow_Click()
    On Error GoTo LOKAL_ERROR
    
    
    Screen.MousePointer = 11
    
    Set dabalokal = Nothing
    gsAnforderung = "ALLES"

    Kopiere
    Me.Refresh
    gsAnforderung = ""
    Zeitanzeigen
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
   
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "cmdNow_Click"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
   
End Sub
Private Sub FileSizeAnzeigen()
    On Error GoTo LOKAL_ERROR
    
    Dim dFilesize As Double
    Dim cPfad As String
    
    If lbl6(64).Caption <> "" Then
        lbl6(51).Visible = True
        lbl6(66).Visible = True
        
        
        lbl6(66).Caption = lbl6(64).Caption
    Else
        lbl6(66).Caption = ""
    
        lbl6(51).Visible = False
        lbl6(66).Visible = False
    End If
    
    cPfad = gcDBPfad      'Dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    dFilesize = FileLen(cPfad & "Kissdata.mdb")     'in BYTE
    dFilesize = dFilesize / 1024                    'in KBYTE
    dFilesize = dFilesize / 1024                    'in MBYTE
    
    lbl6(64).Caption = Format$(dFilesize, "####0.00")
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FileSizeAnzeigen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Zeitanzeigen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sTime As String
    Dim rsrs As Recordset
    Dim lTimeDat As Long
    
    sSQL = "Select * from wkeinste "
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!localtime) Then
            lTimeDat = rsrs!localtime
        Else
            lTimeDat = 0
        End If
        
        sTime = CStr(lTimeDat)
    
        If Len(sTime) = 3 Then
            sTime = Left(sTime, 1) & ":" & Right(sTime, 2)
        Else
            sTime = Left(sTime, 2) & ":" & Right(sTime, 2)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    Text6.Text = sTime
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeitanzeigen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1

End Sub

Private Sub cmdReindizieren_Click()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
        
    If BistDualleineinderDatenbank Then
'        indexDel gdBase
        db_Reindizieren gdBase, lbl6(53), txtStatus, lbl6(28)
        FileSizeAnzeigen
    Else
        anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
    End If
    
    Screen.MousePointer = 0
    
    
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdReindizieren_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo LOKAL_ERROR

    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    
    sTitle = "Speichern des Updatepfades"
    sFilter = "LZH - Dateien (*.lzh)| WK*.lzh"
    sOldpfad = txtUpdatepfad.Text
    gsUpdPfad = pfadaendern(sTitle, sFilter, sOldpfad)
    
    txtUpdatepfad.Text = gsUpdPfad
    speicherpfad
    checkPupdate
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdUpdate_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1

    
End Sub
 Public Sub cmdUpdEinlesen_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet        As Long
    Dim lfail       As Long
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim cZiel2      As String
    Dim t           As Integer
    Dim cPfad       As String
    Dim cPfad1      As String
    Dim stabelle    As String
    Dim Task$, hProcess&, result&
    Dim i           As Integer
    Dim m           As Integer
    Dim cSysPfad    As String
    Dim lWert       As Long
    Dim iStep       As Integer
    
    
    frmWKL00.picprogress.Visible = True
    frmWKL00.txtStatus.Text = 0

    
    iStep = 1
    frmWKL00.txtStatus.Text = iStep
    
    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = ShortPath(cPfad)
    cPfad = cPfad & "Update\"
    
    iStep = 2
    frmWKL00.txtStatus.Text = iStep
    
    cPfad1 = gcDBPfad      'Datenbankpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    Screen.MousePointer = 11
    
    iStep = 3
    frmWKL00.txtStatus.Text = iStep
    
    Kill cPfad & "*.*"
    VerzVorhanden "Update", App.Path & "\"

    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = ShortPath(cPfad)
    cPfad = cPfad & "Update\"
    
    Kill cPfad & "*.*"
    
    
    cQuelle = gsUpdPfad & "\" & gsUpdDatName
    cZiel = cPfad & "LITE!.lzh"
    lRet = CopyFile(cQuelle, cZiel, lfail)
        
    If lRet = 0 Then
        Screen.MousePointer = 0
        MsgBox "Das Kopieren der Datei " & gsUpdDatName & " ist gescheitert.", vbCritical, "Winkiss Fehler:"
        Exit Sub
    End If
    
    If glpVers > 1469 Then
        Zip_Unzip "ICHAG", cPfad, cPfad & "LITE!.lzh", txtStatus
    Else
        t = 2
        Do Until t = 5
            If FileExists(cPfad & "LITE!.lzh") Then
                Task = Shell("LHA" & " e" & Space(1) & cPfad & "LITE!.lzh" & Space(1) & cPfad)
    
                hProcess = OpenProcess(SYNCHRONIZE, False, Task)
                result = WaitForSingleObject(hProcess, INFINITE)
                result = CloseHandle(hProcess)
                t = 5
            End If
        Loop
    End If
    
    Kill cPfad & "LITE!.lzh"
    
    'REPO
    
    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = ShortPath(cPfad)
    cPfad = cPfad & "Update\"
    
    Dim cRepoName As String
    cRepoName = "REPO" & Mid(gsUpdDatName, 3, 4) & ".lzh"

    If FileExists(gsUpdPfad & "\" & cRepoName) Then
    
       cQuelle = gsUpdPfad & "\" & cRepoName
       cZiel = cPfad & "REPO!.lzh"
       lRet = CopyFile(cQuelle, cZiel, lfail)
           
       If lRet = 0 Then
           Screen.MousePointer = 0
           MsgBox "Das Kopieren der Datei " & cRepoName & " ist gescheitert.", vbCritical, "Winkiss Fehler:"
           Exit Sub
       End If
       
        If glpVers > 1469 Then
             Zip_Unzip "ICHAG", cPfad, cPfad & "REPO!.lzh", txtStatus
        Else
            t = 2
            Do Until t = 5
                If FileExists(cPfad & "REPO!.lzh") Then
                    Task = Shell("LHA" & " e" & Space(1) & cPfad & "REPO!.lzh" & Space(1) & cPfad)
        
                    hProcess = OpenProcess(SYNCHRONIZE, False, Task)
                    result = WaitForSingleObject(hProcess, INFINITE)
                    result = CloseHandle(hProcess)
                    t = 5
                End If
            Loop
        End If
       
        Kill cPfad & "REPO!.lzh"
       
        File1.Path = cPfad
        File1.Refresh
       
        For i = 0 To File1.ListCount - 1
       
            m = i
            If iStep + m > 98 Then
            
            Else
                frmWKL00.txtStatus.Text = iStep + m
            End If
            
            If File1.list(i) = "WINKISS.EXE" Then
               
            Else
               stabelle = File1.list(i)
               
               cQuelle = cPfad & stabelle
               cZiel = cPfad1 & stabelle
               cZiel2 = App.Path & "\" & stabelle
               lRet = CopyFile(cQuelle, cZiel, lfail)
               If lRet = 0 Then
                   If i = File1.ListCount - 1 Then
                       Exit For
                   Else
                       i = i + 1
                   End If
               End If
               
               lRet = CopyFile(cQuelle, cZiel2, lfail)
               If lRet = 0 Then
                   If i = File1.ListCount - 1 Then
                       Exit For
                   Else
                       i = i + 1
                       
                   End If
               End If
               
           End If
           
       Next i
    End If
    
    'ENDE REPO

    iStep = 12
    frmWKL00.txtStatus.Text = iStep
    
    
    '*****************Updater persönlich kopieren
    Dim cPfad4 As String
    
    cPfad4 = App.Path      'Anwendungspfad
    If Right(cPfad4, 1) <> "\" Then
        cPfad4 = cPfad4 & "\"
    End If
    
    cQuelle = cPfad & "uWin.exe"
    cZiel = cPfad4 & "uWin.exe"
    lRet = CopyFile(cQuelle, cZiel, lfail)
    
    iStep = 13
    frmWKL00.txtStatus.Text = iStep
    AbmeldungDabaNew
    
    '********************Updater persönlich kopieren Ende
    
    Task = Shell(cPfad4 & "uWin.exe", 1) 'Updater öffnen
    

    Screen.MousePointer = 0
    
    gdBase.Close
    schreibeProtokoll "Abmeldung: meldet sich ab(kissdata.mdb)."
    schreibeProtokollBENUTZERablauf "Abmeldung"
    gdApp.Close
    schreibeProtokoll "Abmeldung: meldet sich ab(kissapp.mdb)."

    frmWKL00.picprogress.Visible = False
    frmWKL00.txtStatus.Text = 0


    End                                             'Winkiss beenden
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "cmdUpdEinlesen_Click"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
        
        Resume Next
        
    End If
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cBeispiele  As String
    Dim iRet        As Integer
    
    Select Case index
        Case Is = 2
            Unload frmWKL53
            
            
            
        Case Is = 3
            Frame7.Visible = True
            lbl6(69).Caption = "Rundungsbeispiele für Variante 1"
            cBeispiele = "10,12 --> 10,10" & vbCrLf
            cBeispiele = cBeispiele & "14,17 --> 14,15" & vbCrLf
            cBeispiele = cBeispiele & "19,75 --> 19,75" & vbCrLf
            cBeispiele = cBeispiele & "51,00 --> 51,00" & vbCrLf
            cBeispiele = cBeispiele & "40,00 --> 39,95" & vbCrLf
            cBeispiele = cBeispiele & "26,02 --> 26,00" & vbCrLf
            cBeispiele = cBeispiele & "28,04 --> 28,05" & vbCrLf
            cBeispiele = cBeispiele & "30,02 --> 29,95" & vbCrLf
            lbl6(70).Caption = cBeispiele
        Case Is = 5
            Frame7.Visible = True
            lbl6(69).Caption = "Rundungsbeispiele für Variante 2"
            cBeispiele = "18,56 --> 18,60" & vbCrLf
            cBeispiele = cBeispiele & "16,24 --> 16,20" & vbCrLf
            cBeispiele = cBeispiele & "16,25 --> 16,30" & vbCrLf
            cBeispiele = cBeispiele & "18,03 --> 18,00" & vbCrLf
            cBeispiele = cBeispiele & "19,05 --> 19,10" & vbCrLf
            cBeispiele = cBeispiele & "15,99 --> 16,00" & vbCrLf
            cBeispiele = cBeispiele & "28,93 --> 28,90" & vbCrLf
            cBeispiele = cBeispiele & "56,50 --> 56,50" & vbCrLf
            lbl6(70).Caption = cBeispiele
        Case Is = 9
            Frame7.Visible = True
            lbl6(69).Caption = "Rundungsbeispiele für Variante 3"
            cBeispiele = "18,56 --> 19,00" & vbCrLf
            cBeispiele = cBeispiele & "16,24 --> 16,50" & vbCrLf
            cBeispiele = cBeispiele & "16,25 --> 16,50" & vbCrLf
            cBeispiele = cBeispiele & "18,03 --> 18,50" & vbCrLf
            cBeispiele = cBeispiele & "19,05 --> 19,50" & vbCrLf
            cBeispiele = cBeispiele & "15,99 --> 16,00" & vbCrLf
            cBeispiele = cBeispiele & "28,93 --> 29,00" & vbCrLf
            cBeispiele = cBeispiele & "56,50 --> 56,50" & vbCrLf
            lbl6(70).Caption = cBeispiele
        Case Is = 0
            Frame7.Visible = True
            lbl6(69).Caption = "Rundungsbeispiele für Variante 4"
            cBeispiele = "18,56 --> 18,95" & vbCrLf
            cBeispiele = cBeispiele & "16,24 --> 16,45" & vbCrLf
            cBeispiele = cBeispiele & "16,45 --> 16,45" & vbCrLf
            cBeispiele = cBeispiele & "18,03 --> 18,45" & vbCrLf
            cBeispiele = cBeispiele & "119,05 --> 119,50" & vbCrLf
            cBeispiele = cBeispiele & "15,99 --> 15,95" & vbCrLf
            cBeispiele = cBeispiele & "28,96 --> 28,95" & vbCrLf
            cBeispiele = cBeispiele & "156,51 --> 157,00" & vbCrLf
            lbl6(70).Caption = cBeispiele
        Case Is = 4
            If lbl6(69).Caption = "Rundungsbeispiele für Variante 1" Then
                Text1(13).Text = RundenS(CDbl(Text1(13).Text))
            ElseIf lbl6(69).Caption = "Rundungsbeispiele für Variante 2" Then
                Text1(13).Text = RundenS2(CDbl(Text1(13).Text))
            ElseIf lbl6(69).Caption = "Rundungsbeispiele für Variante 3" Then
                Text1(13).Text = RundenS3(CDbl(Text1(13).Text))
            ElseIf lbl6(69).Caption = "Rundungsbeispiele für Variante 4" Then
                Text1(13).Text = RundenS4(CDbl(Text1(13).Text))
            End If
            
        Case Is = 6
            Frame7.Visible = False
        Case Is = 7
            iRet = MsgBox("Möchten Sie wirklich die bisherigen Inventureingaben löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbYes Then
                delARTTOINV
            End If
        Case Is = 8
            loeschNEW "ALLARTLU", gdBase
            loeschNEW "GDLAGER", gdBase
            CreateTableT2 "GDLAGER", gdBase

        Case Is = 1
            Select Case tabWK.SelectedItem.Key
            
                Case "Webshop"
                    speicherWebshop
                Case "KissLive"
                    speicherKissLive
                    speicherLive
            
                Case "Kasse"
                    speicherAUSBLEND
'                    speicherParknetto
                    speicherWaage
                    speicherFarbebeiPark
                    
                Case "Tagesabschluss"
                    speicherBargeldEingabe
                    speicherAbNummer
                    speicherZBonWKL53
                    speicherQZBON
                    speicherECTOZ
                    speicherKSF
                    speicherKaMail
                    
                Case "Nachtverarbeitung"
                    speicherNacht
                Case "Sortierung"
                
                Case "Verbindung"
                    speicherDSLandiesemRechner
                    speicherFtpYesNo
                    LeseStammFtp
                    
                    If Check1.value = vbChecked Then
                        zeigestammftp
                    End If
                    
'                    speicherdfu
                    speicherAliasFil
                    speicherOptimierteStamdatenpflege
                    speicherÜberwachung
                    speicherAuto_Export_Artikelbestand
                    
                Case "Farben"
                    speicherfarbe
                    speicherfont
                    speicherpname
                    speicherNoSpruch
                    
                    Modul6.Farbform frmWKL00, frmWKL00.Label1(0)
                    Modul6.Schrift frmWKL00
                    
                Case "Update"
                    speicherpfad
                    speicherREME
                    
                Case "Druckeinstellungen"
                    speicher2BKOPIE
                    speicheretisort
'                    speicherEtiOpt
                    speicherErrDruck
                    speichertabfak
                    speicherDruck27
                    speicherfilmEK
                    speicherEdekaLief
                    
                Case "Sicherung"
                    speicherSicherung
                    
                Case "Unternehmen" 'alias Voreinstellungen
                    speichersSpanne
                    speicherRundung
                    speicherSpezRunden
                    speicherMDE
                    speicherLocalSecurityYesNo
                    speicherAutoLocalModusYesNo
                    speicherArtNrBeg
                    speicherBedienerKarte
                    speicherMWSTBeg
                    speicherBILDTAST
                    
                    speicherFILBONI
                    speicherAlteStada
                    speicherStadapause
                    speicherLUGBERECHNUNG
                    
                Case "Datenbank"
                    speicherUpdCountTime
                    speicherLokalAktualisierungszeit
                    speicherDabakompWann
                    speicherDabakompautono
                    
                Case "WE"
                    speicherWeEinzelFokus
                    speicherWeMenge
                    speicherscanmodi
                    
                Case "ECASH"
                    speicherECASH cboECASH.Text
                    
                Case "Statistik"
                    speicherStatistik
                    GFKstat Trim(Text1(22).Text), Check35.value, Trim(Text1(24).Text)
                    
                    If VEDESstat(Trim(Text1(25)), Check49.value) = True Then
'                        Command35.Visible = True
                    Else
                        loeschNEW "VEDESSTAT", gdBase
                    End If
                
            End Select
            
    End Select
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub speicherLive()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check48.value = vbChecked Then
        sSQL = "Update WKEINSTE Set KL_LIVEBESTAND_DIFF = true"
        gdApp.Execute sSQL, dbFailOnError
        gbKL_LIVEBESTAND_DIFF = True
        
    ElseIf Check48.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set KL_LIVEBESTAND_DIFF = False"
        gdApp.Execute sSQL, dbFailOnError
        gbKL_LIVEBESTAND_DIFF = False
        
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherLive"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub speicherSicherung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim sQuell      As String
    Dim sZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    
    If Check4.value = vbChecked Then
        sSQL = "Update WKEINSTE Set Sichern = True "
        gdApp.Execute sSQL, dbFailOnError
        gbSichernYes = True
    
        If Trim(txtSicherPfad.Text) <> "" Then
            gsSicherPfad = txtSicherPfad.Text
        Else
            gsSicherPfad = gcDBPfad & "\Sicherung"
        End If
        
        sQuell = gcPfad & "\kisslite.ini"
        sZiel = gsSicherPfad & "\kisslite.ini"

        lRet = CopyFile(sQuell, sZiel, lfail)
        If lRet = 0 Then
            lbl6(40).ForeColor = vbRed
            lbl6(40).Caption = "Dies ist kein gültiger Pfad"
            lbl6(40).Refresh
            txtSicherPfad.Text = ""
            txtSicherPfad.SetFocus
            Exit Sub
        End If
        
        speicherSicherungpfad
        lbl6(40).Caption = ""
        lbl6(40).Refresh
        
        If Option9(0).value = True Then
            giSICHTYP = 1
        ElseIf Option9(2).value = True Then
            giSICHTYP = 2
        ElseIf Option9(1).value = True Then
            giSICHTYP = 3
        End If
        
        
        
        
        
        
        
        
        
        sSQL = "Update WKEINSTE Set Sichtyp = " & giSICHTYP & " "
        gdApp.Execute sSQL, dbFailOnError
        
        If giSICHTYP = 3 Then
            sSQL = "Update WKEINSTE Set Sichtime = '" & Right(DTPicker3.value, 8) & "'"
            gdApp.Execute sSQL, dbFailOnError
            gsSICHTIME = DTPicker3.value
        Else
            sSQL = "Update WKEINSTE Set Sichtime = ''"
            gdApp.Execute sSQL, dbFailOnError
            gsSICHTIME = ""
        End If
        
    Else
        sSQL = "Update WKEINSTE Set Sichern = False "
        gdApp.Execute sSQL, dbFailOnError
        gbSichernYes = False
        
        gsSicherPfad = ""
        txtSicherPfad.Text = ""
        
        giSICHTYP = 0
        gsSICHTIME = ""
        
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherSicherung"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub speichertabfak()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    If Trim$(Text1(10).Text) = "" Then
        gdTabfak = 1.3
    Else
        If IsNumeric(Text1(10).Text) Then
            gdTabfak = CDbl(Text1(10).Text)
        Else
            gdTabfak = 1.3
        End If
    End If
    
    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!Tabfak = gdTabfak
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichertabfak"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Sub
Private Sub cmdStandardUp_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    txtUpdatepfad.Text = gcDBPfad & "\In"
    gsUpdPfad = gcDBPfad & "\In"
    
    speicherpfad
    checkPupdate
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdStandardUp_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub



Private Sub Command11_Click()
    On Error GoTo LOKAL_ERROR
    
    giKissFtpMode = 9 'FTPMODE= 9 , Kombimode Kassendateien holen und schicken
    frmWKL38.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1

    
End Sub
Private Sub Command12_Click()
    On Error GoTo LOKAL_ERROR
    
'    cdl1.Flags = cdlCFBoth
    cdl1.Flags = cdlCFPrinterFonts + cdlCFTTOnly + cdlCFANSIOnly + cdlCFLimitSize
    cdl1.Min = 8
    cdl1.Max = 12
    cdl1.FontName = lbl6(48).Caption
    cdl1.FontSize = lbl6(49).Caption
    cdl1.ShowFont
    
    lbl6(48).Caption = cdl1.FontName
    lbl6(49).Caption = cdl1.FontSize
    
     
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command12_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1

End Sub

Private Sub Command13_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case index
        Case 0
            If speichern Then
                cbofuellenNE cboNoEuro
                zustandx
            End If
        Case 1
            If delNoEuro Then
                cbofuellenNE cboNoEuro
                zustandx
            End If
            
    End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Function speichern() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sUK     As String
    Dim sWB     As String
    Dim sWK     As String
    Dim cSQL    As String
    
    speichern = False
    
    If IsNumeric(Text13(2).Text) Then
        sUK = Text13(2).Text
    Else
        anzeigeNew "rot", "Bitte vervollständigen", lblanz
        Text13(2).SetFocus
        Exit Function
    End If
    
    If Text13(0).Text <> "" Then
        sWB = Text13(0).Text
    Else
        anzeigeNew "rot", "Bitte vervollständigen", lblanz
        Text13(0).SetFocus
        Exit Function
    
    End If
    
    If Text13(1).Text <> "" Then
        sWK = Trim(Text13(1).Text)
        
        If SeekInNoEuro(sWK) Then
            
            delNoEuro
            
        Else
        
        End If
    Else
        anzeigeNew "rot", "Bitte vervollständigen", lblanz
        Text13(1).SetFocus
        Exit Function
    End If
    
    cSQL = "Insert into NOEURO (BEDNU,BEDNAME,LASTDATE,WKUERZEL,WBEZEICH,UKURSEUR) values  "
    cSQL = cSQL & " ( " & gcBedienerNr
    cSQL = cSQL & ", '" & gcUserName & "'"
    cSQL = cSQL & ", '" & DateValue(Now) & "'"
    cSQL = cSQL & ", '" & sWK & "'"
    cSQL = cSQL & ", '" & sWB & "'"
    cSQL = cSQL & ", '" & sUK & "'"
    cSQL = cSQL & " ) "
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    speichern = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Function
Private Function delNoEuro() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sWK     As String
    Dim cSQL    As String
    
    delNoEuro = False
    
    If Text13(1).Text <> "" Then
        sWK = Trim(Text13(1).Text)
    Else
        anzeigeNew "rot", "Bitte eine Fremdwährung auswählen!", lblanz
        Exit Function
    End If
    
    cSQL = "Delete from NOEURO where WKuerzel = '" & sWK & "'"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    delNoEuro = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delNoEuro"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Function
Private Function SeekInNoEuro(cwk As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim rs      As Recordset
    Dim cSQL    As String
    
    SeekInNoEuro = False

    cSQL = "Select * from NOEURO where WKuerzel = '" & cwk & "'"
    Set rs = gdBase.OpenRecordset(cSQL)
    If Not rs.EOF Then
        SeekInNoEuro = True
    End If
    rs.Close: Set rs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SeekInNoEuro"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Function
Private Sub zustandx()
    On Error GoTo LOKAL_ERROR
    
    Text13(0).Text = ""
    Text13(1).Text = ""
    Text13(2).Text = ""
    Text13(0).SetFocus
    
    lblanz.ForeColor = vbBlack
    lblanz.Caption = "Bitte eine Fremdwährung auswählen!"
    lblanz.Refresh
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zustandx"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Command14_Click()
On Error GoTo LOKAL_ERROR


    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "Datenbank.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command14_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command15_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "PROFEHLER.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command15_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command16_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "DBERRMEL.txt"
    Screen.MousePointer = 0
    
    Screen.MousePointer = 11
    zeigeHilfeAPPpfad "BIGERR", "RED60.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command16_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
    
    giKissFtpMode = 4 ' FTPMODE= 3 , Programmupdates/ Stammdaten holen
                      ' aus Programmeinstellungen/Update einlesen
    frmWKL38.Show 1
    
    checkPupdate
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    Select Case index
        Case Is = 0 'Programmupdatepfad
            
            Text1(0).Text = gcDBPfad & "\In"
            gsUpdPfad = gcDBPfad & "\In"
    
        Case Is = 1 'Wareneingangspfad
            Text1(1).Text = gcDBPfad & "\Kissdata.mdb"
            gsZinPfad = gcDBPfad & "\Kissdata.mdb"
    
        Case Is = 2 'Kassendateipfad
            Text1(2).Text = gcDBPfad & "\In"
            gsKinPfad = gcDBPfad & "\In"
        Case Is = 3 'Ausgangsdateipfad
            Text1(3).Text = gcDBPfad & "\Kassout"
            gsZoutPfad = gcDBPfad & "\Kassout"
            
        Case 4
            Text1(27).Text = "C:\Fotos"
            gsFotoPfad = "C:\Fotos"
        Case 5
            Text1(28).Text = ""
            gsWebcamPfad = ""
    End Select
    
    speicherpfad
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Command4_Click(index As Integer)
   On Error GoTo LOKAL_ERROR

    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    
    Select Case index
        Case Is = 0 'Programmupdatepfad
        
            sTitle = "Speichern des Updatepfades"
            sFilter = "LZH - Dateien (*.lzh)| WK*.lzh"
            sOldpfad = Text1(0).Text
            gsUpdPfad = pfadaendern(sTitle, sFilter, sOldpfad)
            
            Text1(0).Text = gsUpdPfad
            speicherpfad
        Case Is = 1 'Wareneingangspfad
        
            sTitle = "Speichern des Wareneingangspfades"
            sFilter = "Access - Dateien (*.mdb)|*.mdb"
            
            sOldpfad = gcDBPfad & "\Kissdata.mdb"
            gsZinPfad = pfadaendern(sTitle, sFilter, sOldpfad)
            gsZinPfad = frmWKL00.cdlopen.FileName
   
            
            Text1(1).Text = gsZinPfad
            speicherpfad
        Case Is = 2 'Kassendateipfad
        
            sTitle = "Speichern des Kassendateipfades"
            sFilter = "LZH - Dateien (*.lzh)| Y*.lzh| "
            sOldpfad = Text1(2).Text
            gsKinPfad = pfadaendern(sTitle, sFilter, sOldpfad)
            
            Text1(2).Text = gsKinPfad
            speicherpfad
            
        Case Is = 3 'Zentralausgangspfad F*.lzh
        
            sTitle = "Speichern des Ausgangspfades"
            sFilter = "LZH - Dateien (*.lzh)| F*.lzh| "
            sOldpfad = Text1(2).Text
            gsKinPfad = pfadaendern(sTitle, sFilter, sOldpfad)
            
            Text1(3).Text = gsZoutPfad
            speicherpfad
        Case Is = 4 'Fotopfad
        
            sTitle = "Speichern des Fotopfades"
            sFilter = "JPEG - Dateien (*.JPG)| *.JPG"
            
            sOldpfad = "C:\"
            gsFotoPfad = pfadaendern(sTitle, sFilter, sOldpfad)
            
   
            Text1(27).Text = gsFotoPfad
            speicherpfad
        Case Is = 5 'Webcampfad
        
            sTitle = "Speichern des Webcampfades"
            sFilter = "exe - Dateien (*.exe)| *.exe"
            
            sOldpfad = "C:\"
            gsWebcamPfad = pfadaendern(sTitle, sFilter, sOldpfad)
            gsWebcamPfad = frmWKL00.cdlopen.FileName
   
            Text1(28).Text = gsWebcamPfad
            speicherpfad
    End Select
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Command5_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Dim sQuell      As String
    Dim sZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    

    If Trim(txtSicherPfad.Text) <> "" Then
        gsSicherPfad = txtSicherPfad.Text
    Else
        gsSicherPfad = gcDBPfad & "\Sicherung"
    End If
    
    sQuell = gcPfad & "\kisslite.ini"
    sZiel = gsSicherPfad & "\kisslite.ini"

    lRet = CopyFile(sQuell, sZiel, lfail)
    If lRet = 0 Then
        lbl6(40).ForeColor = vbRed
        lbl6(40).Caption = "Dies ist kein gültiger Pfad"
        lbl6(40).Refresh
        txtSicherPfad.Text = ""
        txtSicherPfad.SetFocus
        Exit Sub
    End If
    
    lbl6(40).ForeColor = glS1
    lbl6(40).Caption = "Sicherung wird durchgeführt..."
    lbl6(40).Refresh
    
    DabaSicherung
    
    gbSichernHeut = False
    Text3.Text = LeselastSicherung
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "PROABL.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub fraVerbindung_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(13).ForeColor = vbBlack
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fraVerbindung_MouseMove"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf."
    
    Fehlermeldung1
End Sub

Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case index
        Case 13
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/225-optimierte-stammdatenpflege-fuer-parfuemerieartikel.html"
    End Select
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If index = 13 Then
        Label1(13).ForeColor = vbBlue
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf."
    
    Fehlermeldung1
End Sub

Private Sub opt1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    If opt1(3).value = True Then
        Frame3(6).Visible = True
    ElseIf opt1(3).value = False Then
        Frame3(6).Visible = False
    End If
    
    If opt1(9).value = True Then
        chk_ZBON_DINA4_HOCH.Visible = True
    ElseIf opt1(9).value = False Then
        chk_ZBON_DINA4_HOCH.Visible = False
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "opt1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf."

    Fehlermeldung1
End Sub

Private Sub Option9_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    If Option9(0).value = True Then
    
        DTPicker3.Visible = False
    
    ElseIf Option9(2).value = True Then
    
        DTPicker3.Visible = False

    ElseIf Option9(1).value = True Then
        
        DTPicker3.Visible = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option9_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Test_Click()
    On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim ctmp        As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = ShortPath(cPfad)
    
    gdBase.Close
    
    If Datenbankreparatur(cPfad, "kissdata.mdb", gsPasswort, lbl6(53), lbl6(28)) Then
        Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
        FileSizeAnzeigen
        MsgBox "Die Datenbank ist repariert, jetzt können alle anderen Winkiss/Kassenrechner weiterarbeiten.", vbInformation, "Winkiss Hinweis:"
    Else
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Test_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    txtSicherPfad.Text = gcDBPfad & "\Sicherung"
    gsSicherPfad = gcDBPfad & "\Sicherung"
    
    speicherSicherungpfad
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKL53
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me

    tabWK.SelectedItem.Key = "Verbindung"
    tabWK_Click
    
    If gbFtpYes = True Then
        Check1.value = vbChecked
'        Frame2.Visible = True
        
        If gbDSL Then
            Check9.value = vbChecked
        Else
            Check9.value = vbUnchecked
        End If
        
        If gbPASSIVMODE Then
            Check110.value = vbChecked
        Else
            Check110.value = vbUnchecked
        End If
        
        If gbFTPautomatic Then
            Check8.value = vbChecked
        Else
            Check8.value = vbUnchecked
        End If
        
        If gbFtpZENT Then
            Check7.value = vbChecked
            Frame4.Visible = True
            Text2(6).Text = giFILALI
            
            If gbWVNOT Then
                Check88.value = vbChecked
            Else
                Check88.value = vbUnchecked
            End If
        End If
        zeigestammftp
        Command2.Enabled = True
    Else
        Check1.value = vbUnchecked
'        Frame2.Visible = False
    End If
    
    If gbDSL = True Then
        Check9.value = vbChecked
        Check91.Visible = True
        
        If gbOptiStada Then
            Check91.value = vbChecked
        Else
            Check91.value = vbUnchecked
        End If
        
        If gbOptiStadaSpiel Then
            Check40.value = vbChecked
        Else
            Check40.value = vbUnchecked
        End If
    Else
        Check9.value = vbUnchecked
        Check40.Visible = False
        Check40.value = vbUnchecked
        Check91.Visible = False
        Check91.value = vbUnchecked
    End If
    
    If gbSPY Then
        Check55.value = vbChecked
        Frame18.Visible = True
        
        Text2(8).Text = gsServerIP
        Text2(9).Text = gsServerPort
    Else
        Check55.value = vbUnchecked
        Frame18.Visible = False
        
        Text2(8).Text = ""
        Text2(9).Text = ""
    End If
    
    If gbAuto_Export_Artikelbestand Then
        Check58.value = vbChecked
    Else
        Check58.value = vbUnchecked
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub fuellecombo6dfue()
    On Error GoTo LOKAL_ERROR
    
    Dim S&
    Dim LN&
    Dim X%
    Dim R(255)  As RASENTRYNAME95
    Dim lRet    As Long

    Screen.MousePointer = 11

    '### Namen der bestehenden DFÜ-Verbindungen einlesen
    R(0).dwSize = 264
    S = 256 * R(0).dwSize
    lRet = RasEnumEntries(vbNullString, vbNullString, R(0), S, LN)
    
    Combo6.AddItem "keine DFÜ vorhanden"
    
    If lRet = 0 Then
        If LN <> 0 Then
            '### Es besteht mindestens eine DFÜ-Verbindung
            For X = 0 To LN - 1
                
                ConName = StrConv(R(X).szEntryName(), vbUnicode)
                Combo6.AddItem ConName
                
                If Left(ConName, Len(gsDFU)) = gsDFU Then
                
                    Combo6.Text = gsDFU
                    Combo6.RemoveItem 0
                
                End If

            Next X
        Else
            Combo6.Text = "keine DFÜ vorhanden"
        End If
    End If

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 383 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "fuellecombo6dfue"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. globaler DFÜ-Name: " & gsDFU
    
        Fehlermeldung1
    End If
End Sub
Private Sub PositionierenWKL53()
    On Error GoTo LOKAL_ERROR
    
    Dim ltop        As Long
    Dim lleft       As Long
    Dim lwidth      As Long
    Dim lHeight     As Long
    
    Dim i           As Integer
    
    ltop = 840
    lleft = 360
    lwidth = 11055
    lHeight = 6015
    
    frmWKL53.Caption = gsPname & " Programmeinstellungen"
    
    frmWKL53.Height = 8000
    frmWKL53.Width = 11805
    
    For i = 0 To frmWKL53.Controls.Count - 1
        
        If TypeOf frmWKL53.Controls(i) Is Frame Then 'alle Frames
        
            If frmWKL53.Controls(i).Tag = 1 Then
                frmWKL53.Controls(i).Top = ltop
                frmWKL53.Controls(i).Left = lleft
                frmWKL53.Controls(i).Height = lHeight
                frmWKL53.Controls(i).Width = lwidth
                frmWKL53.Controls(i).Visible = False
            End If
        End If
    Next i
   
    Frame7.Height = 5655
    Frame7.Left = 120
    Frame7.Top = 240
    Frame7.Width = 10815
    
    Frame20.Height = 1455
    Frame20.Left = 120
    Frame20.Top = 3360
    Frame20.Width = 2175


    fraadt.Height = 4335
    fraadt.Left = 120
    fraadt.Top = 840
    fraadt.Width = 10815
    
    fraELP.Height = 4335
    fraELP.Left = 120
    fraELP.Top = 840
    fraELP.Width = 10815
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL53"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub zeigestammftp()
    On Error GoTo LOKAL_ERROR

    Option1(giStammFTPOFT).value = True
            
    Text2(2).Text = gsStammFTPPASS
    Text2(1).Text = gsStammFTPUSER
    Text2(0).Text = gsStammFTPAdresse
    
    Text2(3).Text = gsZenFTPPASS
    Text2(4).Text = gsZenFTPUSER
    Text2(5).Text = gsZenFTPAdresse
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigestammftp"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub zeigestammftp_OnlyZentrale()
    On Error GoTo LOKAL_ERROR

    Text2(3).Text = gsZenFTPPASS
    Text2(4).Text = gsZenFTPUSER
    Text2(5).Text = gsZenFTPAdresse
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigestammftp_OnlyZentrale"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Function GetDateFromWeek(ByVal nWeek As Integer, _
  Optional ByVal nDayOfWeek As VBA.VbDayOfWeek = vbMonday, _
  Optional ByVal nYear As Integer = -1) As Date
 
  Dim nCurWeek As Integer
  Dim vStart As Variant
  Dim vMonday As Variant
  Dim vSunday As Variant
  Dim nDay As Integer
 
  ' Kein Jahr angeben? Dann aktuelles Jahr verwenden!
  If nYear = -1 Then nYear = Year(Now)
 
  ' aktuelle Woche im Jahr nYear ermitteln
  vStart = DateSerial(nYear, Month(Now), Day(Now))
  nCurWeek = Val(Format$(vStart, "ww", vbMonday))
 
  ' Datum der gewünschten Woche ermitteln
  vStart = DateAdd("ww", nWeek - nCurWeek, vStart)
 
  ' Wochenanfang ermitteln
  nDay = Weekday(vStart, vbMonday)
 
  ' Datum des gewünschten Wochentags ermitteln
  If nDayOfWeek = vbSunday Then
    GetDateFromWeek = DateAdd("d", -nDay + 7, vStart)
  Else
    GetDateFromWeek = DateAdd("d", -nDay + nDayOfWeek - 1, vStart)
  End If
End Function
Private Sub tabWK_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    Dim cPfad1      As String
    Dim ctmp        As String
    Dim cdatei      As String
    Dim rsrs        As Recordset
    Dim sSQL        As String
    Dim iFileNr     As Integer
    Dim i           As Integer
    Dim cpfaddb     As String
    
    cpfaddb = gcDBPfad
    If Right(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If
    
    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Update"
    
    cPfad1 = gcDBPfad      'dabapfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    For i = 0 To frmWKL53.Controls.Count - 1
        If TypeOf frmWKL53.Controls(i) Is Frame And frmWKL53.Controls(i).Tag = 1 Then   'alle Frames
            frmWKL53.Controls(i).Visible = False
        End If
    Next i
    
    Select Case tabWK.SelectedItem.Key
    
        Case "Webshop"
            fraWebshop.ZOrder
            fraWebshop.Visible = True
            
            Text21(11).Text = gsMySQL_PHP_SCRIPT_PFAD
            
            If gbMySQL_LIVEBESTAND = True Then
                Check71.value = vbChecked
                
                Text21(8).Text = gsMySQL_BESTAND_TAB
                Text21(9).Text = gsMySQL_BESTAND_INDEXSPALTE
                Text21(10).Text = gsMySQL_BESTAND_SPALTE
            Else
                Check71.value = vbUnchecked
                
                Text21(8).Text = ""
                Text21(9).Text = ""
                Text21(10).Text = ""
            End If
    
        Case "KissLive"
            fraKisslive.ZOrder
            fraKisslive.Visible = True
            
            Text21(0).Text = gsKL_ADRESSE
            Text21(1).Text = gsKL_BENUTZER
            Text21(2).Text = gsKL_PASSWORT
            Text21(3).Text = gsKL_DATENBANKNAME
            Text21(6).Text = gsKL_DSN
            
            If gbKL_LIVEBESTAND = True Then
                Check29.value = vbChecked
            Else
                Check29.value = vbUnchecked
            End If
            
            If gbKL_LIVENACHRICHTEN = True Then
                Check86.value = vbChecked
            Else
                Check86.value = vbUnchecked
            End If
            
            If gbKL_LIVEBESTAND = True Then
                If gbKL_LIVEBESTAND_DIFF = True Then
                    Check48.value = vbChecked
                Else
                    Check48.value = vbUnchecked
                End If
            Else
                Check48.Enabled = False
            End If
            
            
            If gbKL_LIVEKVKPR = True Then
                Check26.value = vbChecked
            Else
                Check26.value = vbUnchecked
            End If
            
            If gbKL_LIVEFarbe = True Then
                Check46.value = vbChecked
            Else
                Check46.value = vbUnchecked
            End If
            
            If gbKL_LIVEGefSperr = True Then
                Check38.value = vbChecked
            Else
                Check38.value = vbUnchecked
            End If
            
            
            
            
            
            
            
            
            If gbKL_LIVEGUTSCHEIN = True Then
                Check3.value = vbChecked
            Else
                Check3.value = vbUnchecked
            End If
            
        Case "Kasse"
            fraKasse.ZOrder
            fraKasse.Visible = True
            
            If gbTPbf = True Then
                Check109(2).value = vbChecked
            Else
                Check109(2).value = vbUnchecked
            End If
            
            If gbNeukunden = True Then
                Check109(0).value = vbChecked
            Else
                Check109(0).value = vbUnchecked
            End If
            
            If gbSterne = True Then
                Check108.value = vbChecked
            Else
                Check108.value = vbUnchecked
            End If
            
            
                       
            Text23.Text = giFarbebeiPark
            
            cbocomfuell_Waage
         
        Case "Tagesabschluss"
            
            fraTagesabschluss.ZOrder
            fraTagesabschluss.Visible = True
            
            If gsZählbeleg = "Listendrucker" Then
                opt1(1).value = False
                opt1(10).value = True
            ElseIf gsZählbeleg = "Bondrucker" Then
                opt1(1).value = True
                opt1(10).value = False
            Else
                opt1(1).value = False
                opt1(10).value = False
            End If
            
            If gsZBon = "Listendrucker" Then
                opt1(8).value = False
                opt1(9).value = True
            ElseIf gsZBon = "Bondrucker" Then
                opt1(8).value = True
                opt1(9).value = False
            Else
                opt1(8).value = False
                opt1(9).value = False
            End If
            
            If gbAABSCHL = True Then
                Check79.value = vbChecked
            Else
                Check79.value = vbUnchecked
            End If
            
            If gbKSF = True Then
                Check47.value = vbChecked
                
                lbl6(59).Visible = False
                DTPicker2.Visible = False
                
            Else
                lbl6(59).Visible = True
                DTPicker2.Visible = True
                
                
                Check47.value = vbUnchecked
                If gsKassDatstart <> "" Then
                    DTPicker2.value = gsKassDatstart
                Else
                    DTPicker2.value = "20:00:00"
                End If
            End If
            
            
            If gbMitExport = True Then
                Check57.value = vbChecked
                Text8.Text = gsKaMail
            Else
                Check57.value = vbUnchecked
                Text8.Text = ""
            End If
            
            
            
            
            
            If gbZBONDINA4HOCH = True Then
                chk_ZBON_DINA4_HOCH.value = vbChecked
            Else
                chk_ZBON_DINA4_HOCH.value = vbUnchecked
            End If
            
            
            
            If gbQZBON = True Then
                Frame13.Visible = True
                leseABREport
                Check25.value = vbChecked
                
                If gbARTKUM = True Then
                    Check43.value = vbChecked
                    
                    
                    If gbARTKUM_ohneWGN = True Then
                        Check41.value = vbChecked
                    Else
                        Check41.value = vbUnchecked
                    End If
                    
                    
                    
                Else
                    Check43.value = vbUnchecked
                    Check41.value = vbUnchecked
                End If
                
                If gbTAGFILT = True Then
                    Check72.value = vbChecked
                Else
                    Check72.value = vbUnchecked
                End If
                
                If gbKK = True Then
                    Check44.value = vbChecked
                Else
                    Check44.value = vbUnchecked
                End If
                
                If gbEA = True Then
                    Check45.value = vbChecked
                Else
                    Check45.value = vbUnchecked
                End If
                
                
            Else
                Frame13.Visible = False
                gbTAGFILT = False
                gbARTKUM = False
                gbKK = False
                gbEA = False
                Check25.value = vbUnchecked
            End If
            
            If gbBargeldEingabe = True Then
                Check22.value = vbChecked
            Else
                Check22.value = vbUnchecked
            End If
            
           
            
            If gbAbschlussNummer Then
                Check13.value = vbChecked
            Else
                Check13.value = vbUnchecked
            End If
            
            If gbAbschlussDatum Then
                Check14.value = vbChecked
            Else
                Check14.value = vbUnchecked
            End If
            
            If gbAGNAusw Then
                Check83.value = vbChecked
            Else
                Check83.value = vbUnchecked
            End If
            
            If gbARTKUMWGN Then
                Check93.value = vbChecked
            Else
                Check93.value = vbUnchecked
            End If
            
            If gbKUMSUM Then
                Check62.value = vbChecked
            Else
                Check62.value = vbUnchecked
            End If
            
            If gbECTOZ = True Then
                Check39.value = vbChecked
            Else
                Check39.value = vbUnchecked
            End If
    
        Case "Nachtverarbeitung"
            fraNacht.ZOrder
            fraNacht.Visible = True
            
            If gbNacht = True Then
                Check67.value = vbChecked
                einschalt
                leseNacht
                
                If gbEXTSICH = True Then
                    Check92.value = vbChecked
                Else
                    Check92.value = vbUnchecked
                End If
                
                If gbMB = True Then
                    Check90.value = vbChecked
                Else
                    Check90.value = vbUnchecked
                End If
                
                If gbPCAus = True Then
                    Check68.value = vbChecked
                Else
                    Check68.value = vbUnchecked
                End If
                
                If gbWKAUS = True Then
                    Check104.value = vbChecked
                Else
                    Check104.value = vbUnchecked
                End If
                
                If gbBR = True Then
                    Checkbox7.value = vbChecked
                Else
                    Checkbox7.value = vbUnchecked
                End If
                
                If gbSTAMDA = True Then
                    Checkbox9.value = vbChecked
                Else
                    Checkbox9.value = vbUnchecked
                End If
                
                If gbKABSCH = True Then
                    Checkbox8.value = vbChecked
                Else
                    Checkbox8.value = vbUnchecked
                End If
                
                If gbUmsatzNeu = True Then
                    Check84.value = vbChecked
                Else
                    Check84.value = vbUnchecked
                End If
                
                If gsNachtstart <> "" Then
                    
                    DTPicker1.value = gsNachtstart
                    
                Else
                    DTPicker1.value = "20:00:00"
                End If
                
                
                
                If gbFtpYes = True Then
                
                    If gbUPRO = True Then
                        Checkbox1.value = vbChecked
                    Else
                        Checkbox1.value = vbUnchecked
                    End If
                    
                    If gbUSTADA = True Then
                        Checkbox2.value = vbChecked
                    Else
                        Checkbox2.value = vbUnchecked
                    End If
                
                    If gbUSTAT = True Then
                        Checkbox3.value = vbChecked
                    Else
                        Checkbox3.value = vbUnchecked
                    End If
                    If gbFtpZENT = True Then
                        If gbUKDAT = True Then
                            Checkbox5.value = vbChecked
                        Else
                            Checkbox5.value = vbUnchecked
                        End If
                    
                        If gbEKDAT = True Then
                            Checkbox6.value = vbChecked
                            
                            Combo2.Text = giSTARTMIN
                            Combo3.Text = giINTERV
                        Else
                            Checkbox6.value = vbUnchecked
                        End If
                    End If
                Else
                    Checkbox1.value = vbUnchecked
                    Checkbox2.value = vbUnchecked
                    Checkbox3.value = vbUnchecked
                    Checkbox5.value = vbUnchecked
                    Checkbox6.value = vbUnchecked
                
                End If
                
                
            Else
                Check67.value = vbUnchecked
                ausschalt
            End If
    
        Case "Sortierungen"
            fraSort.ZOrder
            fraSort.Visible = True
            
        Case "Statistik"
            fraSta.ZOrder
            fraSta.Visible = True
            
            Text1(0).Text = gsUpdPfad
            Text1(1).Text = gsZinPfad
            Text1(2).Text = gsKinPfad
            Text1(3).Text = gsZoutPfad
            fraDaba.Visible = False
            fraECASH.Visible = False
            
            Text1(20).Text = DatePart("yyyy", DateValue(Now))
            
            Text1(21).Text = DatePart("ww", DateValue(Now))
            If Year(DateValue(Now)) = 2017 Then
                Text1(21).Text = CStr(CInt(Text1(21).Text) - 1)
            End If
            
            
            
            
            
            
            
            
            Dim vMonday As Date
            Dim vSunday As Date
 
            vMonday = GetDateFromWeek(CInt(Text1(21).Text), vbMonday, Year(Now))
            vSunday = GetDateFromWeek(CInt(Text1(21).Text), vbSunday, Year(Now))
'            lbl6(81).Caption = "aktuelle KW = " & DatePart("ww", DateValue(Now)) & " (" & vMonday & "-" & vSunday & ")"
            lbl6(81).Caption = "aktuelle KW = " & Text1(21).Text & " (" & vMonday & "-" & vSunday & ")"
            
           

            
            
            
            lbl6(95).Caption = Format(DateValue(Now), "DD.MM.YY")
            
            Set Command36(0).Picture = LoadPicture(cpfaddb & "Picture\System\" & "Kalender.jpg")
            Command36(0).BackColorFrom = vbWhite
            Command36(0).BackColorTo = vbWhite
            Command36(0).PictureAlign = 3
            
            'GFKSTAT auslesen
            If NewTableSuchenDBKombi("GFKSTAT", gdBase) Then
                Text1(22).Text = leseGFKstat("kundnr")
                If CBool(leseGFKstat("automatik")) = False Then
                    Check35.value = vbUnchecked
                Else
                    Check35.value = vbChecked
                End If
                Text1(24).Text = leseGFKstat("email")
            Else
                Text1(22).Text = ""
                Text1(24).Text = ""
                Check35.value = vbUnchecked
            End If
            'GFKSTAT auslesen ENDE
            
            'VEDESSTAT auslesen
            If NewTableSuchenDBKombi("VEDESSTAT", gdBase) Then
                
                If CBool(leseVEDESstat("live")) = False Then
                    Check49.value = vbUnchecked
                    Text1(25).Text = ""
                Else
                    Check49.value = vbChecked
                    Text1(25).Text = leseVEDESstat("marktnr")
                    
                End If
                
            Else
                
                Text1(25).Text = ""
                Check49.value = vbUnchecked
            End If
            'VEDESSTAT auslesen ENDE
            
            If gbUnistatWeek = True Then
                Check10.value = vbChecked
                cmd12.Visible = True
                Label3.Visible = True
                Text11.Text = gdateStatlast
                Text11.Visible = True
                Text10.Text = gsStatkundnr
                Text10.Visible = True
                lbl6(39).Visible = True
                Text20.Text = gsStatZusatz
                Text20.Visible = True
                lbl6(86).Visible = True
                Check59.Visible = True
                
                If gbDSL = True Then
                    Check59.Enabled = True
                    
                    If gbStatweekperMail = True Then
                        Check59.value = vbChecked
                    Else
                        Check59.value = vbUnchecked
                    End If
                Else
                    Check59.Enabled = False
                    Check59.value = vbUnchecked
                End If
                
                
            Else
                cmd12.Visible = False
                Text11.Visible = False
                Label3.Visible = False
                lbl6(39).Visible = False
                Text10.Visible = False
                lbl6(86).Visible = False
                Text20.Visible = False
                Check59.Visible = False
            End If
            
            If gbUnistatMonat = True Then
                Check87.value = vbChecked
                If gdateMStatlast = 0 Then
                    Command27.Visible = True
                Else
                    Command22.Visible = True
                End If
                Label9.Visible = True
                Text16.Text = gdateMStatlast
                Text16.Visible = True
                Text15.Text = gsMStatkundnr
                Text15.Visible = True
                lbl6(73).Visible = True
            Else
                Command27.Visible = False
                Command22.Visible = False
                Text16.Visible = False
                Label9.Visible = False
                lbl6(73).Visible = False
                Text15.Visible = False
            End If
            
            Text19.Text = gsKUPFAD
            
        Case "Pfade"
            fraPfade.ZOrder
            fraPfade.Visible = True
            
            Text1(0).Text = gsUpdPfad
            Text1(1).Text = gsZinPfad
            Text1(2).Text = gsKinPfad
            Text1(3).Text = gsZoutPfad
            Text1(27).Text = gsFotoPfad
            Text1(28).Text = gsWebcamPfad
            fraDaba.Visible = False
            fraECASH.Visible = False
            
        Case "Fremdwährung"
            fraNoEURO.ZOrder
            fraNoEURO.Visible = True
            
            cbofuellenNE cboNoEuro
            zustandx
           
            
        Case "Sicherung"
            fraSicher.ZOrder
            fraSicher.Visible = True
            
            lbl6(40).Caption = ""
            lbl6(40).Refresh
            
            If gbSichernYes Then
                Check4.value = vbChecked
                Frame10.Visible = True
                txtSicherPfad.Text = gsSicherPfad
                
                Select Case giSICHTYP
                    Case 0
                        Option9(0).value = True
                    Case 1
                        Option9(0).value = True
                    Case 2
                        Option9(2).value = True
                    Case 3
                        Option9(1).value = True
                    Case Else
                        Option9(0).value = True
                End Select
                
                If giSICHTYP = 3 Then
                    DTPicker3.value = gsSICHTIME
                End If
                
            Else
                Check4.value = vbUnchecked
                Frame10.Visible = False
                txtSicherPfad.Text = ""
            End If
        
            Text3.Text = LeselastSicherung
            
        Case "Verbindung"
            fraVerbindung.ZOrder
            fraVerbindung.Visible = True
            
'            fuellecombo6dfue
           
         Case "ECASH"
            fraECASH.ZOrder
            fraECASH.Visible = True
           
            combofuell
           
            If gbEcash = True Then
                Check6.value = vbChecked
                Select Case gsEPartner
                    Case Is = "ADT"
                        leseadtopt
                        
                        If gsAdtVerfahren = "INOUT" Then
                            

                        ElseIf gsAdtVerfahren = "XML" Then
                           
                            Frame15.Visible = True
                            Option2(1).value = True
                            Text12.Text = Trim(gADTtermId)
                            Text14.Text = gADTLimit
                            Text9.Text = Trim(gADTclientId)
                            Text18.Text = Trim(gADTipAdress)
                            Text17.Text = Trim(gADTport)
                            cboECASH.Text = "ADT Wellcom GmbH"
                            cboECASH.Visible = True
                            fraadt.Visible = True
                            
                            If gbADTAE Then Check19.value = vbChecked
                            If gbADTVI Then Check18.value = vbChecked
                            If gbADTDI Then Check17.value = vbChecked
                            If gbADTEU Then Check20.value = vbChecked
                        End If
                        
                    Case Is = "ELP"
                           
                        cboECASH.Text = "elPAY"
                        cboECASH.Visible = True
                        fraELP.Visible = True
                        fraELP.Caption = "elPAY"
                        
                    Case Is = "ZV2"
                           
                        cboECASH.Text = "ZV2"
                        cboECASH.Visible = True
                        fraELP.Visible = True
                        fraELP.Caption = "ZV2"
                        
                    Case Is = "ZVT"
                           
                        cboECASH.Text = "ZVT"
                        cboECASH.Visible = True
                        fraELP.Visible = True
                        fraELP.Caption = "ZVT"
                        
                    
                            
                    Case Else
                        fraadt.Visible = False
                        Check6.value = vbUnchecked
                End Select
                   
            ElseIf gbEcash = False Then
                Check6.value = vbUnchecked
                cboECASH.Visible = False
                fraadt.Visible = False
                gsEPartner = ""
            End If
            
         Case "Datenbank"
            fraDaba.ZOrder
            fraDaba.Visible = True
            
            If gbDabakompfrueh Then
                Check16.value = vbChecked
            Else
                Check16.value = vbUnchecked
            End If
            
            If gbDabakompautoNo Then
                Check103.value = vbChecked
            Else
                Check103.value = vbUnchecked
            End If
            
            If gbOhneAnzeige Then
                Check27.value = vbChecked
            Else
                Check27.value = vbUnchecked
            End If
            
            
            If gbKopOhneAuswertung Then
                Check42.value = vbChecked
            Else
                Check42.value = vbUnchecked
            End If
            
            
            If gbPenner_faerben Then
                Check105.value = vbChecked
            Else
                Check105.value = vbUnchecked
            End If
            
            Combo1.Text = glLokalAktuZeit
            Text4.Text = glUPDCOUNT
            Text5.Text = glUPDTime
            Zeitanzeigen
            
            Text7(1).Text = Format(gdDBPAUSE, "######0.00")
            Text7(0).Text = DatumLastKompAnzeigen
            Text7(2).Text = DatumLastKompZeitAnzeigen
            FileSizeAnzeigen
            
            
            
            
            If gbLokalModus Then
                gcDBPfad = "C:\aLeer"
            Else
            
                cdatei = "KISSLITE.INI"
                
                iFileNr = FreeFile
                Open gcPfad & "KISSLITE.INI" For Binary As #iFileNr
                If LOF(iFileNr) > 0 Then
                    
                    ctmp = Space$(LOF(iFileNr))
                    Get #iFileNr, 1, ctmp
                    gcDBPfad = ctmp
                    Close iFileNr
                End If
            End If
            lbl6(27).Caption = gcDBPfad
            
        Case "Farben"
            fraFarben.ZOrder
            fraFarben.Visible = True
            
            If gsPname <> "" Then
                Text1(18).Text = gsPname
            Else
                Text1(18).Text = "Winkiss"
            End If
            
            lbl6(48).Caption = gsFont
            lbl6(49).Caption = gsFontsize
            
            cmd2(0).BackColor = glH1
            cmd2(1).BackColor = glU1
            cmd2(2).BackColor = glS1
            cmd2(3).BackColor = glH2
            cmd2(4).BackColor = glSelBack1
            
            cmd2(22).BackColor = glWarn
            cmd2(21).BackColor = glLink
            
            If glButtonForecolor = 0 Then
                cmd2(5).BackColor = cmdBeispiel.ForeColor
            Else
                cmd2(5).BackColor = glButtonForecolor
                cmdBeispiel.ForeColor = glButtonForecolor
            End If
            
            If glButtonHintergrund_from = 0 Then
                cmd2(6).BackColor = cmdBeispiel.BackColorFrom
            Else
                cmd2(6).BackColor = glButtonHintergrund_from
                cmdBeispiel.BackColorFrom = glButtonHintergrund_from
            End If
            
            If glButtonHintergrund_to = 0 Then
                cmd2(7).BackColor = cmdBeispiel.BackColorTo
            Else
                cmd2(7).BackColor = glButtonHintergrund_to
                cmdBeispiel.BackColorTo = glButtonHintergrund_to
            End If
            
            
            If glButtonMouseMove_Hintergrund_from = 0 Then
                cmd2(9).BackColor = cmdBeispiel.HoverColorFrom
            Else
                cmd2(9).BackColor = glButtonMouseMove_Hintergrund_from
                cmdBeispiel.HoverColorFrom = glButtonMouseMove_Hintergrund_from
            End If
            
            If glButtonMouseMove_Hintergrund_to = 0 Then
                cmd2(10).BackColor = cmdBeispiel.HoverColorTo
            Else
                cmd2(10).BackColor = glButtonMouseMove_Hintergrund_to
                cmdBeispiel.HoverColorTo = glButtonMouseMove_Hintergrund_to
            End If
            
            If glButtonMouseMove_Forecolor = 0 Then
                cmd2(12).BackColor = cmdBeispiel.ForeColorHover
            Else
                cmd2(12).BackColor = glButtonMouseMove_Forecolor
                cmdBeispiel.ForeColorHover = glButtonMouseMove_Forecolor
            End If
            
            If glButtonBordercolor = 0 Then
                cmd2(8).BackColor = cmdBeispiel.BorderColor
            Else
                cmd2(8).BackColor = glButtonBordercolor
                cmdBeispiel.BorderColor = glButtonBordercolor
            End If
            
            If glButtonMouseMove_Bordercolor = 0 Then
                cmd2(11).BackColor = cmdBeispiel.BorderColorHover
            Else
                cmd2(11).BackColor = glButtonMouseMove_Bordercolor
                cmdBeispiel.BorderColorHover = glButtonMouseMove_Bordercolor
            End If
            
            If gbNoSpruch = True Then
                Check60.value = vbChecked
            Else
                Check60.value = vbUnchecked
            End If
            
            If gbSound = True Then
                Check78.value = vbChecked
            Else
                Check78.value = vbUnchecked
            End If
            
            If gbISDEMO = True Then
                Check81.value = vbChecked
            Else
                Check81.value = vbUnchecked
            End If
            
            
        Case "Update"
            fraUpdate.ZOrder
            fraUpdate.Visible = True
            
            txtUpdatepfad.Text = gsUpdPfad
            checkPupdate
            
            If gbSTADAP = False Then
                Check34.value = vbChecked
            Else
                Check34.value = vbUnchecked
            End If
            
            If gbFTH = False Then
                Check61.value = vbChecked
            Else
                Check61.value = vbUnchecked
            End If
            
            If gbKVKSicher = True Then
                Check70.value = vbChecked
            Else
                Check70.value = vbUnchecked
            End If
            
            If gbNOWOCHENDATEN = True Then
                Check24.value = vbChecked
            Else
                Check24.value = vbUnchecked
            End If
            
        Case "WE"
            fraWE.ZOrder
            fraWE.Visible = True
            
            Text1(17).Text = gsWeEinzMe
            
            If gbscanmodi Then
                Check15.value = vbChecked
            Else
                Check15.value = vbUnchecked
            End If
            
            If gbWEautoGef Then
                Check51.value = vbChecked
            Else
                Check51.value = vbUnchecked
            End If
            
            If gbNONEGZU Then
                Check21.value = vbChecked
            Else
                Check21.value = vbUnchecked
            End If
            
            If gbAutoZwsp Then
                Check63.value = vbChecked
            Else
                Check63.value = vbUnchecked
            End If
            
            If gbETIONLYME Then
                Check33.value = vbChecked
            Else
                Check33.value = vbUnchecked
            End If
            
            If gbNoETIWeAusBe Then
                Check50.value = vbChecked
            Else
                Check50.value = vbUnchecked
            End If
            
            If Trim(gsWeEinzFo) = "EAN" Then
                Option3(3).value = False
                Option3(2).value = True
            Else
                Option3(3).value = True
                Option3(2).value = False
            End If
            
        Case "Druckeinstellungen"
            fraDruck.ZOrder
            fraDruck.Visible = True
            
            If gbETIBEIFARB Then
                Check85.value = vbChecked
            Else
                Check85.value = vbUnchecked
            End If
            
            If gbDruck27 Then
                Check36.value = vbChecked
                
                Check74.Visible = True
                
                If gbPAEBON Then
                    Check74.value = vbChecked
                Else
                    Check74.value = vbUnchecked
                End If
                    
                
            Else
                Check36.value = vbUnchecked
                Check74.Visible = False
                Check74.value = vbUnchecked
            End If
            
            If gbSaveReport Then
                Check31.value = vbChecked
            Else
                Check31.value = vbUnchecked
            End If
            
            
            
            If gbFILMEK Then
                Check37.value = vbChecked
            Else
                Check37.value = vbUnchecked
            End If
            
            If gbErrPrint = True Then
                opt1(13).value = False
                opt1(14).value = True
            Else
                opt1(13).value = True
                opt1(14).value = False
            End If
            
            If gbEtiFokEan = True Then
                opt1(24).value = True
                opt1(25).value = False
            Else
                opt1(25).value = True
                opt1(24).value = False
            End If
            
'            If tableSuchenDBKombi("VOREINAP", 2) Then
'                Option3(0).Value = LeseVoreinap(Option3(0).Caption)
'                Option3(1).Value = LeseVoreinap(Option3(1).Caption)
'                Option6(0).Value = LeseVoreinap(Option6(0).Caption)
'                Option6(1).Value = LeseVoreinap(Option6(1).Caption)
'                Option6(2).Value = LeseVoreinap(Option6(2).Caption)
'                Option6(3).Value = LeseVoreinap(Option6(3).Caption)
'                Option6(4).Value = LeseVoreinap(Option6(4).Caption)
'                Option6(5).Value = LeseVoreinap(Option6(5).Caption)
'                Option6(6).Value = LeseVoreinap(Option6(6).Caption)
'                Option6(7).Value = LeseVoreinap(Option6(7).Caption)
'
'            End If
            Option5(giSortierung).value = True
            
            Text1(10).Text = gdTabfak
            
            If gbEtiEan Then
                Check54.value = vbChecked
            Else
                Check54.value = vbUnchecked
            End If
            
            Text1(26).Text = gsEdeka
            
            If gbEtiQuickScanM Then
                Check77.value = vbChecked
            Else
                Check77.value = vbUnchecked
            End If
            
        Case "Unternehmen" 'alias Voreinstellungen
            fraUnter.ZOrder
            fraUnter.Visible = True
            
            If gbBEDKARTE = True Then
                Check12.value = vbChecked
                sSQL = "select bedkarte from dbeinste"
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.RecordCount = 0 Then
                
                    If rsrs!bedkarte = True Then
                        Check11.value = vbChecked
                    Else
                        Check11.value = vbUnchecked
                    End If
                    
                End If
                rsrs.Close: Set rsrs = Nothing
            Else
                If gbQPASS = True Then
                    Check23.value = vbChecked
                Else
                    Check23.value = vbUnchecked
                End If
                
                Check12.value = vbUnchecked
                Check11.value = vbUnchecked
            End If
            
            If gbOLDSTADADEL Then
                Check82.value = vbChecked
            Else
                Check82.value = vbUnchecked
            End If
            
            Text1(14).Text = gdStadaPause
            
            If gbGTBON = True Then
                Check73.value = vbChecked
            Else
                Check73.value = vbUnchecked
            End If
            
            If gbNewArt = True Then
                Check65.value = vbChecked
            Else
                Check65.value = vbUnchecked
            End If
            
            If gbNewArtNrVorschlag = True Then
                Check53.value = vbChecked
            Else
                Check53.value = vbUnchecked
            End If
            
            If gbArtEindeut = True Then
                Check66.value = vbChecked
            Else
                Check66.value = vbUnchecked
            End If
            
            
            
            If gbAA = True Then
                Check64.value = vbChecked
            Else
                Check64.value = vbUnchecked
            End If
            
            If gbTagAkt = True Then
                Check76.value = vbChecked
            Else
                Check76.value = vbUnchecked
            End If
            
            If gbBILDTAST = True Then
                Check32.value = vbChecked
            Else
                Check32.value = vbUnchecked
            End If
            
            If gsSpanne = "LEK" Then
                opt1(2).value = False
                opt1(3).value = True
                If gbEKMAX = True Then
                    opt1(22).value = True
                Else
                    opt1(23).value = True
                End If
            Else
                opt1(2).value = True
                opt1(3).value = False
            End If
            
            
            
            If gsMDEGERAET = "FORCOM" Then
                opt1(6).value = True
                opt1(7).value = False
                opt1(0).value = False
                opt1(11).value = False
                opt1(12).value = False
            ElseIf gsMDEGERAET = "REWEMDE" Then
                opt1(0).value = True
                opt1(7).value = False
                opt1(6).value = False
                opt1(11).value = False
                opt1(12).value = False
            ElseIf gsMDEGERAET = "SCANPAL" Then
                opt1(7).value = True
                opt1(0).value = False
                opt1(6).value = False
                opt1(11).value = False
                opt1(12).value = False
            ElseIf gsMDEGERAET = "BELAMDE" Then
                opt1(11).value = True
                opt1(7).value = False
                opt1(6).value = False
                opt1(0).value = False
                opt1(12).value = False
            ElseIf gsMDEGERAET = "CASIOMDE" Then
                opt1(12).value = True
                opt1(7).value = False
                opt1(6).value = False
                opt1(0).value = False
                opt1(11).value = False
            End If
            
            Text1(12).Text = giMDEPAUSE
            
            If gbLocalSec Then
                Check2.value = vbChecked
            Else
                Check2.value = vbUnchecked
            End If
            
            If gbAutoLokalModus Then
                Check5.value = vbChecked
            Else
                Check5.value = vbUnchecked
            End If
            
            If gbAutoSYN Then
                Check28.value = vbChecked
            Else
                Check28.value = vbUnchecked
            End If
            
            Text1(5).Text = giAufrunden
            Text1(6).Text = giAbrunden
            Text1(7).Text = giRundkrit
            
            If gsMWST <> "" Then
                Text1(19).Text = gsMWST
            Else
                Text1(19).Text = "V"
            End If
            
            'Eintrag der Beginner ArtNr
            'Welche Filnr?
            
            If Trim(gcFilNr) = "0" Then
                Text1(9).Text = glArtNrBeg
                Text1(9).Locked = False
            ElseIf Trim(gcFilNr) = "1" Then
                Text1(9).Text = 551000
                Text1(9).Locked = True
            Else
                If Len(gcFilNr) = 2 Then
                    Text1(9).Text = "5" & gcFilNr & "000"
                ElseIf Len(gcFilNr) = 1 Then
                    Text1(9).Text = "50" & gcFilNr & "000"
                End If
                Text1(9).Locked = True
            End If
            
            
            
            If gbSPEZRU = True Then
                Check80.value = vbChecked
                If gbSPEZVAR = 1 Then
                    opt1(15).value = True
                ElseIf gbSPEZVAR = 2 Then
                    opt1(16).value = True
                ElseIf gbSPEZVAR = 3 Then
                    opt1(21).value = True
                ElseIf gbSPEZVAR = 4 Then
                    opt1(4).value = True
                End If
            Else
                Check80.value = vbUnchecked
            End If
            
            If Check80.value = vbChecked Then
                opt1(15).Visible = True
                opt1(16).Visible = True
                opt1(21).Visible = True
                opt1(4).Visible = True
                Command1(3).Visible = True
                Command1(5).Visible = True
                Command1(9).Visible = True
                Command1(0).Visible = True
            Else
                opt1(15).Visible = False
                opt1(16).Visible = False
                opt1(21).Visible = False
                opt1(4).Visible = False
                Command1(3).Visible = False
                Command1(5).Visible = False
                Command1(9).Visible = False
                Command1(0).Visible = False
            End If
            
            
            
            If gbREGEB = True Then
                Check52.value = vbChecked
                Text1(23).Text = giGebTage
                
                If gbGebAdresse = True Then
                    Check69.value = vbChecked
                Else
                    Check69.value = vbUnchecked
                End If
                
            Else
                Check52.value = vbUnchecked
                Check69.value = vbUnchecked
                Text1(23).Text = "2"
            End If
            
            If gbGesEKWert_anzeigen = True Then
                Check56.value = vbChecked
            Else
                Check56.value = vbUnchecked
            End If
            
            If gbyLugBe = 1 Then
                opt1(20).value = True
            ElseIf gbyLugBe = 2 Then
                opt1(19).value = True
            ElseIf gbyLugBe = 3 Then
                opt1(18).value = True
            ElseIf gbyLugBe = 4 Then
                opt1(17).value = True
            End If
            
            Text1(15).Text = giTageVerkauf
            Text1(16).Text = giTageZugang
            
            cbocomfuell
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "tabWK_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    Resume Next
    
End Sub
Private Sub cbocomfuell()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    cbocom.Clear
    cbocom.Visible = True
    
    For i = 1 To 255
        cbocom.AddItem i
    Next i
    cbocom.Visible = True
    
    cbocom.Text = gbYtescanPcom
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbocomfuell"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub cbocomfuell_Waage()
    On Error GoTo LOKAL_ERROR
    
    With Combo4
        .Clear
        .Visible = True
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .Visible = True
        .Text = gbYteWAAGEPcom
    End With
    
    With Combo5
        .Clear
        .Visible = True
        .AddItem "TP-II"
        .AddItem "keine Waage"
        .Visible = True
        .Text = gsWAAGE
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbocomfuell_Waage"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub combofuell()
    On Error GoTo LOKAL_ERROR

    cboECASH.Clear
    cboECASH.Visible = True
    cboECASH.AddItem "ZV2"
    cboECASH.AddItem "ZVT"
    'cboECASH.AddItem "ADT Wellcom GmbH"
    cboECASH.AddItem "elPAY"
    cboECASH.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "combofuell"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub checkPupdate()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet    As Integer
    Dim ctmp    As String
    
    iRet = NEWfnCheck4UpdateDateiWKL00(False)
    If iRet <> 0 Then
        ctmp = ctmp & "Es liegt für Sie ein Programm-Update vor!"
        lbl6(0).ForeColor = vbRed
        lbl6(0).Caption = ctmp
        lbl6(0).Refresh
        cmdUpdEinlesen.Enabled = True
    Else
        lbl6(0).ForeColor = vbBlack
        lbl6(0).Caption = "Nein, Es liegt kein neues Programmupdate vor. (Klicken Sie auf 'Update holen')"
        lbl6(0).Refresh
        cmdUpdEinlesen.Enabled = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkPupdate"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    

End Sub
Private Sub speicherfarbe()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    glH1 = cmd2(0).BackColor
    glU1 = cmd2(1).BackColor
    glS1 = cmd2(2).BackColor
    glH2 = cmd2(3).BackColor
    glSelBack1 = cmd2(4).BackColor
    glLink = cmd2(21).BackColor
    glWarn = cmd2(22).BackColor
    
    glButtonHintergrund_from = cmd2(6).BackColor
    glButtonHintergrund_to = cmd2(7).BackColor
    glButtonMouseMove_Hintergrund_from = cmd2(9).BackColor
    glButtonMouseMove_Hintergrund_to = cmd2(10).BackColor
    glButtonMouseMove_Bordercolor = cmd2(11).BackColor
    glButtonBordercolor = cmd2(8).BackColor
    glButtonMouseMove_Forecolor = cmd2(12).BackColor
    glButtonForecolor = cmd2(5).BackColor
    
    
    
    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!H1 = glH1
        rsrs!U1 = glU1
        rsrs!S1 = glS1
        rsrs!H2 = glH2
        rsrs!SB1 = glSelBack1
        rsrs!WARN = glWarn
        rsrs!LINK = glLink
        
        rsrs!ButtonHintergrund_from = glButtonHintergrund_from
        rsrs!ButtonHintergrund_to = glButtonHintergrund_to
        rsrs!ButtonMouseMove_Hintergrund_from = glButtonMouseMove_Hintergrund_from
        rsrs!ButtonMouseMove_Hintergrund_to = glButtonMouseMove_Hintergrund_to
        rsrs!ButtonMouseMove_Bordercolor = glButtonMouseMove_Bordercolor
        rsrs!ButtonBordercolor = glButtonBordercolor
        rsrs!ButtonMouseMove_Forecolor = glButtonMouseMove_Forecolor
        rsrs!ButtonForecolor = glButtonForecolor
        
        rsrs.Update
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherfarbe"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1

End Sub
Private Sub speicherpname()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    If Text1(18).Text <> "" Then
        gsPname = Text1(18).Text
    Else
        gsPname = "Winkiss"
    End If
    
    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!pname = gsPname
        rsrs.Update
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherpname"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherfont()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    gsFont = lbl6(48).Caption
    gsFontsize = lbl6(49).Caption
    
    If gsFont = "" Then
        gsFont = "Arial"
    End If
    
    If gsFontsize = 0 Then
        gsFontsize = 12
    End If
    
    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!Font = gsFont
        rsrs!FontSize = gsFontsize
        rsrs.Update
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherfont"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherNacht()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check67.value = vbChecked Then
    
        sSQL = "Update WKEINSTE Set NACHT = true "
        gdApp.Execute sSQL, dbFailOnError
    
        gbNacht = True
        speichernachtDetails
    
    ElseIf Check67.value = vbUnchecked Then
    
        sSQL = "Update WKEINSTE Set NACHT = false "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update DBEINSTE set NachtStart = '' "
        gdBase.Execute sSQL, dbFailOnError
    
        loeschNEW "NACHT", gdApp
        leseNacht
        
        gbNacht = False
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherNacht"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speichernachtDetails()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sStart As String
    
    loeschNEW "NACHT", gdApp
    CreateTable "NACHT", gdApp
    
    
    'Nachtstart Text
    sStart = Right(DTPicker1.value, 8)
    
    If sStart <> "" Then
    
        gsNachtstart = sStart
        sSQL = "Insert into Nacht (NachtStart) values ('" & sStart & "')"
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update DBEINSTE set NachtStart = '" & sStart & "'"
        gdBase.Execute sSQL, dbFailOnError
        
    Else
    
        gsNachtstart = ""

    End If
    
    'StartMin
    
    sSQL = "Update Nacht Set StartMin = " & CInt(Combo2.Text)
    gdApp.Execute sSQL, dbFailOnError

    giSTARTMIN = CInt(Combo2.Text)
    
    'INTERV
    
    sSQL = "Update Nacht Set INTERV = " & CInt(Combo3.Text)
    gdApp.Execute sSQL, dbFailOnError

    giINTERV = CInt(Combo3.Text)


   
    
    'Pcaus bit
    
    If Check68.value = vbChecked Then
    
        sSQL = "Update Nacht Set PCAUS = true "
        gdApp.Execute sSQL, dbFailOnError
        
        gbPCAus = True
    
    ElseIf Check68.value = vbUnchecked Then
    
        sSQL = "Update Nacht Set PCAUS = false "
        gdApp.Execute sSQL, dbFailOnError
        
        gbPCAus = False
    
    End If
    
    'WKaus bit
    
    If Check104.value = vbChecked Then
    
        sSQL = "Update Nacht Set WKAUS = true "
        gdApp.Execute sSQL, dbFailOnError
        
        gbWKAUS = True
    
    ElseIf Check104.value = vbUnchecked Then
    
        sSQL = "Update Nacht Set WKAUS = false "
        gdApp.Execute sSQL, dbFailOnError
        
        gbWKAUS = False
    
    End If
    
    'UPRO
    
    If Checkbox1.value = vbChecked Then

        sSQL = "Update Nacht Set UPRO = true "
        gdApp.Execute sSQL, dbFailOnError

        gbUPRO = True

    ElseIf Checkbox1.value = vbUnchecked Then

        sSQL = "Update Nacht Set UPRO = false "
        gdApp.Execute sSQL, dbFailOnError

        gbUPRO = False

    End If
    
    'BR
    
    If Checkbox7.value = vbChecked Then

        sSQL = "Update Nacht Set BR = true "
        gdApp.Execute sSQL, dbFailOnError

        gbBR = True

    ElseIf Checkbox7.value = vbUnchecked Then

        sSQL = "Update Nacht Set BR = false "
        gdApp.Execute sSQL, dbFailOnError

        gbBR = False

    End If
    
    'STAMDA
    
    If Checkbox9.value = vbChecked Then

        sSQL = "Update Nacht Set STAMDA = true "
        gdApp.Execute sSQL, dbFailOnError

        gbSTAMDA = True

    ElseIf Checkbox9.value = vbUnchecked Then

        sSQL = "Update Nacht Set STAMDA = false "
        gdApp.Execute sSQL, dbFailOnError

        gbSTAMDA = False

    End If
    
    'MB
    If Check90.value = vbChecked Then
        sSQL = "Update Nacht Set MB = true "
        gdApp.Execute sSQL, dbFailOnError

        gbMB = True
    ElseIf Check90.value = vbUnchecked Then
        sSQL = "Update Nacht Set MB = false "
        gdApp.Execute sSQL, dbFailOnError

        gbMB = False
    End If
    
    'extern Sichern
    If Check92.value = vbChecked Then
        sSQL = "Update Nacht Set EXTSICH = true "
        gdApp.Execute sSQL, dbFailOnError

        gbEXTSICH = True
    ElseIf Check92.value = vbUnchecked Then
        sSQL = "Update Nacht Set EXTSICH = false "
        gdApp.Execute sSQL, dbFailOnError

        gbEXTSICH = False
    End If
    
    
    'KABSCH
    If Checkbox8.value = vbChecked Then
        sSQL = "Update Nacht Set KABSCH = true "
        gdApp.Execute sSQL, dbFailOnError
        gbKABSCH = True
    ElseIf Checkbox8.value = vbUnchecked Then
        sSQL = "Update Nacht Set KABSCH = false "
        gdApp.Execute sSQL, dbFailOnError
        gbKABSCH = False
    End If
    
    'umsartneu
    If Check84.value = vbChecked Then
        sSQL = "Update Nacht Set UmsatzNeu = true "
        gdApp.Execute sSQL, dbFailOnError
        gbUmsatzNeu = True
    ElseIf Check84.value = vbUnchecked Then
        sSQL = "Update Nacht Set UmsatzNeu = false "
        gdApp.Execute sSQL, dbFailOnError
        gbUmsatzNeu = False
    End If
    
    
    'USTADA
    
    If Checkbox2.value = vbChecked Then

        sSQL = "Update Nacht Set USTADA = true "
        gdApp.Execute sSQL, dbFailOnError

        gbUSTADA = True

    ElseIf Checkbox2.value = vbUnchecked Then

        sSQL = "Update Nacht Set USTADA = false "
        gdApp.Execute sSQL, dbFailOnError

        gbUSTADA = False

    End If
    
    'USTAT
    
    If Checkbox3.value = vbChecked Then

        sSQL = "Update Nacht Set USTAT = true "
        gdApp.Execute sSQL, dbFailOnError

        gbUSTAT = True

    ElseIf Checkbox3.value = vbUnchecked Then

        sSQL = "Update Nacht Set USTAT = false "
        gdApp.Execute sSQL, dbFailOnError

        gbUSTAT = False

    End If
    
    'UKDAT
    
    If Checkbox5.value = vbChecked Then

        sSQL = "Update Nacht Set UKDAT = true "
        gdApp.Execute sSQL, dbFailOnError

        gbUKDAT = True

    ElseIf Checkbox5.value = vbUnchecked Then

        sSQL = "Update Nacht Set UKDAT = false "
        gdApp.Execute sSQL, dbFailOnError

        gbUKDAT = False

    End If
    
    
    'EKDAT
    
    If Checkbox6.value = vbChecked Then

        sSQL = "Update Nacht Set EKDAT = true "
        gdApp.Execute sSQL, dbFailOnError

        gbEKDAT = True

    ElseIf Checkbox6.value = vbUnchecked Then

        sSQL = "Update Nacht Set EKDAT = false "
        gdApp.Execute sSQL, dbFailOnError

        gbEKDAT = False

    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernachtDetails"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherAbschlussdetails()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sStart As String
    
    loeschNEW "ABREPORT", gdApp
    CreateTable "ABREPORT", gdApp

    'artkum bit
    
    If Check43.value = vbChecked Then
        sSQL = "Insert into ABREPORT (artkum) values (true)"
        gdApp.Execute sSQL, dbFailOnError
        gbARTKUM = True
        
        
        If Check41.value = vbChecked Then
            sSQL = "Update ABREPORT Set artkum_OHNEWGN = true "
            gdApp.Execute sSQL, dbFailOnError
        
            gbARTKUM_ohneWGN = True
        ElseIf Check41.value = vbUnchecked Then
            sSQL = "Update ABREPORT Set artkum_OHNEWGN = False "
            gdApp.Execute sSQL, dbFailOnError
            gbARTKUM_ohneWGN = False
        End If
        
        
        
    ElseIf Check43.value = vbUnchecked Then
        sSQL = "Insert into ABREPORT (artkum) values (False)"
        gdApp.Execute sSQL, dbFailOnError
        gbARTKUM = False
        
        sSQL = "Update ABREPORT Set artkum_OHNEWGN = False "
        gdApp.Execute sSQL, dbFailOnError
        gbARTKUM_ohneWGN = False
    End If
    
    'TAGFILT bit
    
    If Check72.value = vbChecked Then
        sSQL = "Update ABREPORT Set TAGFILT = true "
        gdApp.Execute sSQL, dbFailOnError
        gbTAGFILT = True
    ElseIf Check72.value = vbUnchecked Then
        sSQL = "Update ABREPORT Set TAGFILT = false "
        gdApp.Execute sSQL, dbFailOnError
        gbTAGFILT = False
    End If
    
    'kk
    If Check44.value = vbChecked Then
        sSQL = "Update ABREPORT Set kk = true "
        gdApp.Execute sSQL, dbFailOnError
        gbKK = True
    ElseIf Check44.value = vbUnchecked Then
        sSQL = "Update ABREPORT Set kk = false "
        gdApp.Execute sSQL, dbFailOnError
        gbKK = False
    End If
    
    'ea
    If Check45.value = vbChecked Then
        sSQL = "Update ABREPORT Set ea = true "
        gdApp.Execute sSQL, dbFailOnError
        gbEA = True
    ElseIf Check45.value = vbUnchecked Then
        sSQL = "Update ABREPORT Set ea = false "
        gdApp.Execute sSQL, dbFailOnError
        gbEA = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAbschlussdetails"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf."
    
    Fehlermeldung1
End Sub
'Private Sub speicherEtiOpt()
'    On Error GoTo LOKAL_ERROR
'
'    Dim sSQL As String
'
'    If Not tableSuchenDBKombi("VOREINAP", 2) Then
'        sSQL = "Create table VOREINAP ( Schluessel Text(30),Wert Text(30) )"
'        gdApp.Execute sSQL, dbFailOnError
'    End If
'
'    sSQL = "Delete from Voreinap"
'    gdApp.Execute sSQL, dbFailOnError
'
'    If Option3(0).Value = True Then
'        SchreibeVoreinap Option3(0).Caption, "EIN"
'    ElseIf Option3(1).Value = True Then
'        SchreibeVoreinap Option3(1).Caption, "EIN"
'    End If
'
'    If Option6(0).Value = True Then
'        SchreibeVoreinap Option6(0).Caption, "EIN"
'    ElseIf Option6(1).Value = True Then
'        SchreibeVoreinap Option6(1).Caption, "EIN"
'    ElseIf Option6(2).Value = True Then
'        SchreibeVoreinap Option6(2).Caption, "EIN"
'    ElseIf Option6(3).Value = True Then
'        SchreibeVoreinap Option6(3).Caption, "EIN"
'    ElseIf Option6(4).Value = True Then
'        SchreibeVoreinap Option6(4).Caption, "EIN"
'    ElseIf Option6(5).Value = True Then
'        SchreibeVoreinap Option6(5).Caption, "EIN"
'    ElseIf Option6(6).Value = True Then
'        SchreibeVoreinap Option6(6).Caption, "EIN"
'    ElseIf Option6(7).Value = True Then
'        SchreibeVoreinap Option6(7).Caption, "EIN"
'    End If
'
'    Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "speicherEtiOpt"
'    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
'
'    Fehlermeldung1
'End Sub
'Private Sub SchreibeVoreinap(schluessel As String, Wert As String)
'    On Error GoTo LOKAL_ERROR
'        Dim sSQL As String
'
'        sSQL = "Insert into VOREINAP (Schluessel,Wert) values ('" & schluessel & "', '" & Wert & "')"
'        gdApp.Execute sSQL, dbFailOnError
'
'    Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "SchreibeVoreinap"
'    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
'
'    Fehlermeldung1
'End Sub
Private Sub speicherFtpYesNo()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim j As Integer
    
    For j = 0 To 2
        If Option1(j).value = True Then Exit For
    Next j
    
    If NewTableSuchenDBKombi("StammFTP", gdBase) = False Then
        CreateTableT2 "STAMMFTP", gdBase
    End If
    
    
    If Text2(0).Text <> "" And Text2(1).Text <> "" And Text2(2).Text <> "" Then
        sSQL = "Delete from StammFTP where FTPNAME = 'Filiale'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into StammFTP  (FTPAD,FTPUS,FTPPA,ftpoft,lastftp,FTPNAME) values ('1','1','1',0,datevalue(now),'Filiale')"
        gdBase.Execute sSQL, dbFailOnError
            
        sSQL = "Update StammFTP Set FTPAD = '" & Text2(0).Text & "', FTPUS = '" & Text2(1).Text & "', FTPPA = '" & Text2(2).Text & "' , ftpoft = " & j & " , Lastftp = datevalue(now)-1"
        sSQL = sSQL & "  where FTPNAME = 'Filiale' "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set FTP = True "
        gdApp.Execute sSQL, dbFailOnError
        gbFtpYes = True
        
        If Check8.value = vbChecked Then
            gbFTPautomatic = True
            sSQL = "Update WKEINSTE Set FTPauto = True "
            gdApp.Execute sSQL, dbFailOnError
        Else
            gbFTPautomatic = False
            sSQL = "Update WKEINSTE Set FTPauto = False "
            gdApp.Execute sSQL, dbFailOnError
        End If
        
        
        
        If Check110.value = vbChecked Then
            gbPASSIVMODE = True
            sSQL = "Update WKEINSTE Set PASSIVMODE = True "
            gdApp.Execute sSQL, dbFailOnError
        Else
            gbPASSIVMODE = False
            sSQL = "Update WKEINSTE Set PASSIVMODE = False "
            gdApp.Execute sSQL, dbFailOnError
        End If
        
        LeseStammFtp
    End If
        
    If Check1.value = vbChecked Then
        sSQL = "Update WKEINSTE Set FTP = true "
        gdApp.Execute sSQL, dbFailOnError
        gbFtpYes = True
    Else
        sSQL = "Update WKEINSTE Set FTP = False "
        gdApp.Execute sSQL, dbFailOnError
        gbFtpYes = False
    End If
    
    If Text2(3).Text <> "" And Text2(4).Text <> "" And Text2(5).Text <> "" Then
        sSQL = "Delete from StammFTP where FTPNAME = 'ZENTRALE'"
        gdBase.Execute sSQL, dbFailOnError

        sSQL = "Insert into StammFTP  (FTPAD,FTPUS,FTPPA,lastftp,FTPNAME) "
        sSQL = sSQL & "values ('" & Text2(5).Text & "','" & Text2(4).Text & "','" & Text2(3).Text & "',datevalue(now),'Zentrale')"
        gdBase.Execute sSQL, dbFailOnError
        
        If Check88.value = vbChecked Then
            gbWVNOT = True
            sSQL = "Update WKEINSTE Set WVNOT = True "
            gdApp.Execute sSQL, dbFailOnError
        Else
            gbWVNOT = False
            sSQL = "Update WKEINSTE Set WVNOT = False "
            gdApp.Execute sSQL, dbFailOnError
        End If

        LeseStammFtp
    Else
        gbWVNOT = False
        sSQL = "Update WKEINSTE Set WVNOT = False "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    If Check7.value = vbChecked Then
        sSQL = "Update WKEINSTE Set FTPZENT = true "
        gdApp.Execute sSQL, dbFailOnError
        gbFtpZENT = True
    Else
        sSQL = "Update WKEINSTE Set FTPZENT = false "
        gdApp.Execute sSQL, dbFailOnError
        gbFtpZENT = False
    End If

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherFtpYesNo"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherDSLandiesemRechner()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check9.value = vbChecked Then
        gbDSL = True
        sSQL = "Update WKEINSTE Set FTPautoh = True "
        gdApp.Execute sSQL, dbFailOnError
    Else
        gbDSL = False
        sSQL = "Update WKEINSTE Set FTPautoh = False "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDSLandiesemRechner"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherLocalSecurityYesNo()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check2.value = vbChecked Then
        sSQL = "Update WKEINSTE Set LocalSec = True "
        gdApp.Execute sSQL, dbFailOnError
        gbLocalSec = True
    Else
        sSQL = "Update WKEINSTE Set LocalSec = False "
        gdApp.Execute sSQL, dbFailOnError
        gbLocalSec = False
        
        Kill "c:\aleer\kissdata.mdb"
    End If
            
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherLocalSecurityYesNo"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    End If
End Sub
Private Sub speicherLokalAktualisierungszeit()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    glLokalAktuZeit = CLng(Combo1.Text)
    
    sSQL = "Update WKEINSTE Set UPDLOKAL =  " & glLokalAktuZeit
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherLokalAktualisierungszeit"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherDabakompWann()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check16.value = vbChecked Then
        gbDabakompfrueh = True
    Else
        gbDabakompfrueh = False
    End If
    
    If gbDabakompfrueh Then
        sSQL = "Update DBEINSTE Set STORNO = true "
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update DBEINSTE Set STORNO = false "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Check105.value = vbChecked Then
        sSQL = "Update DBEINSTE Set PENNERFARB = true "
        gdBase.Execute sSQL, dbFailOnError
        gbPenner_faerben = True
    Else
        sSQL = "Update DBEINSTE Set PENNERFARB = false "
        gdBase.Execute sSQL, dbFailOnError
        gbPenner_faerben = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDabakompWann"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherDabakompautono()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check103.value = vbChecked Then
        gbDabakompautoNo = True
    Else
        gbDabakompautoNo = False
    End If
    
    If gbDabakompautoNo Then
        sSQL = "Update WKEINSTE Set NOAUTO = true "
        gdApp.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update WKEINSTE Set NOAUTO = false "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    'ohneAnzeige
    
    If Check27.value = vbChecked Then
        gbOhneAnzeige = True
    Else
        gbOhneAnzeige = False
    End If
    
    If gbOhneAnzeige Then
        sSQL = "Update WKEINSTE Set ohneAnzeige = true "
        gdApp.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update WKEINSTE Set ohneAnzeige = false "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    
    If Check42.value = vbChecked Then
        gbKopOhneAuswertung = True
    Else
        gbKopOhneAuswertung = False
    End If
    
    If gbKopOhneAuswertung Then
        sSQL = "Update WKEINSTE Set KopOhneAuswertung = true "
        gdApp.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update WKEINSTE Set KopOhneAuswertung = false "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDabakompautono"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherUpdCountTime()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Text4.Text <> "" Then
        sSQL = "Update DBEINSTE Set UPDCOUNT = " & CLng(Text4.Text)
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Text5.Text <> "" Then
        sSQL = "Update DBEINSTE Set UPDTIME = " & CLng(Text5.Text)
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Text7(1).Text <> "" Then
        sSQL = "Update DBEINSTE Set DBPAUSE = '" & Text7(1).Text & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    glUPDCOUNT = CLng(Text4.Text)
    glUPDTime = CLng(Text5.Text)
    gdDBPAUSE = CDbl(Text7(1).Text)
    
    DBEngine.SetOption dbLockRetry, glUPDCOUNT
    DBEngine.SetOption dbLockDelay, glUPDTime
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherUpdCountTime"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherAutoLocalModusYesNo()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check5.value = vbChecked Then
        sSQL = "Update WKEINSTE Set autoLModus = True "
        gdApp.Execute sSQL, dbFailOnError
        gbAutoLokalModus = True
    Else
        sSQL = "Update WKEINSTE Set autoLModus = False "
        gdApp.Execute sSQL, dbFailOnError
        gbAutoLokalModus = False
    End If
    
    If Check28.value = vbChecked Then
        sSQL = "Update WKEINSTE Set autoSYN = True "
        gdApp.Execute sSQL, dbFailOnError
        gbAutoSYN = True
        
    Else
        sSQL = "Update WKEINSTE Set autoSYN = False "
        gdApp.Execute sSQL, dbFailOnError
        gbAutoSYN = False
    End If
            
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAutoLocalModusYesNo"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub


Private Sub speicherAlteStada()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check82.value = vbChecked Then
        sSQL = "Update DBEINSTE Set OLDSTADADEL = True "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        gbOLDSTADADEL = True
    Else
        sSQL = "Update DBEINSTE Set OLDSTADADEL = False "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        gbOLDSTADADEL = False
        
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAlteStada"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherLUGBERECHNUNG()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    gbyLugBe = 2
    If opt1(20).value = True Then
        gbyLugBe = 1
    ElseIf opt1(19).value = True Then
        gbyLugBe = 2
    ElseIf opt1(18).value = True Then
        gbyLugBe = 3
    ElseIf opt1(17).value = True Then
        gbyLugBe = 4
    End If
    
    giTageVerkauf = 365
    If Text1(15).Text <> "" Then
        If IsNumeric(Text1(15).Text) Then
            giTageVerkauf = CInt(Text1(15).Text)
        End If
    End If
    sSQL = "Update DBEINSTE set LUGTAGV = " & giTageVerkauf
    gdBase.Execute sSQL, dbFailOnError
    
    giTageZugang = 365
    If Text1(16).Text <> "" Then
        If IsNumeric(Text1(16).Text) Then
            giTageZugang = CInt(Text1(16).Text)
        End If
    End If
    sSQL = "Update DBEINSTE set LUGTAGZ = " & giTageZugang
    gdBase.Execute sSQL, dbFailOnError
        
    sSQL = "Update DBEINSTE set LUGBE = " & gbyLugBe
    gdBase.Execute sSQL, dbFailOnError
                        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherLUGBERECHNUNG"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherStadapause()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Text1(14).Text = "" Then
        Text1(14).Text = "0"
    Else
        If IsNumeric(Text1(14).Text) = False Then
            Text1(14).Text = "0"
        End If
    End If
    
    sSQL = "Update DBEINSTE Set STADAPAUSE =  '" & Text1(14).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    gdStadaPause = Text1(14).Text
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherStadapause"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherREME()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check34.value = vbChecked Then
        sSQL = "Update WKEINSTE Set STADAP = False "
        gdApp.Execute sSQL, dbFailOnError
        gbSTADAP = False
    Else
        sSQL = "Update WKEINSTE Set STADAP = true "
        gdApp.Execute sSQL, dbFailOnError
        gbSTADAP = True
    End If
    
    If Check61.value = vbChecked Then
        sSQL = "Update DBEINSTE Set FTH = False "
        gdBase.Execute sSQL, dbFailOnError
        gbFTH = False
    Else
        sSQL = "Update DBEINSTE Set FTH = true "
        gdBase.Execute sSQL, dbFailOnError
        gbFTH = True
    End If
    
    If Check70.value = vbChecked Then
        sSQL = "Update DBEINSTE Set KVKSICHER = True "
        gdBase.Execute sSQL, dbFailOnError
        gbKVKSicher = True
    Else
        sSQL = "Update DBEINSTE Set KVKSICHER = false "
        gdBase.Execute sSQL, dbFailOnError
        gbKVKSicher = False
    End If
    
    If Check24.value = vbChecked Then
        sSQL = "Update DBEINSTE Set NOWOCHENDATEN = True "
        gdBase.Execute sSQL, dbFailOnError
        gbNOWOCHENDATEN = True
    Else
        sSQL = "Update DBEINSTE Set NOWOCHENDATEN = false "
        gdBase.Execute sSQL, dbFailOnError
        gbNOWOCHENDATEN = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "speicherREME"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub speicherDruck27()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check36.value = vbChecked Then
        sSQL = "Update DBEINSTE Set DRUCK27 = true "
        gdBase.Execute sSQL, dbFailOnError
        gbDruck27 = True
        
        
        If Check74.value = vbChecked Then
            sSQL = "Update WKEINSTE Set PAEBON = true "
            gdApp.Execute sSQL, dbFailOnError
            gbPAEBON = True
        Else
            sSQL = "Update WKEINSTE Set PAEBON = false "
            gdApp.Execute sSQL, dbFailOnError
            gbPAEBON = False
        End If
        
        
        
    Else
        sSQL = "Update DBEINSTE Set DRUCK27 = false "
        gdBase.Execute sSQL, dbFailOnError
        gbDruck27 = False
        
        sSQL = "Update WKEINSTE Set PAEBON = false "
        gdApp.Execute sSQL, dbFailOnError
        gbPAEBON = False
    End If
    
    
    
    
    
    
    
    If Check85.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ETIBEIFARB = true "
        gdBase.Execute sSQL, dbFailOnError
        gbETIBEIFARB = True
    Else
        sSQL = "Update DBEINSTE Set ETIBEIFARB = false "
        gdBase.Execute sSQL, dbFailOnError
        gbETIBEIFARB = False
    End If
    
    If Check31.value = vbChecked Then
        sSQL = "Update KASSEIN Set SaveReport = true "
        gdBase.Execute sSQL, dbFailOnError
        gbSaveReport = True
    Else
        sSQL = "Update KASSEIN Set SaveReport = false "
        gdBase.Execute sSQL, dbFailOnError
        gbSaveReport = False
    End If
    
    
    
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDruck27"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherfilmEK()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check37.value = vbChecked Then
        sSQL = "Update DBEINSTE Set FILMEK = true "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        gbFILMEK = True
    Else
        sSQL = "Update DBEINSTE Set FILMEK = false "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        gbFILMEK = False
    End If
    
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherfilmEK"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicher2BKOPIE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    
    If Check54.value = vbChecked Then
        sSQL = "Update WKEINSTE Set ETIEAN = true "
        gdApp.Execute sSQL, dbFailOnError
        gbEtiEan = True
    Else
        sSQL = "Update WKEINSTE Set ETIEAN = false "
        gdApp.Execute sSQL, dbFailOnError
        gbEtiEan = False
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher2BKOPIE"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherErrDruck()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If opt1(14).value = True Then
        sSQL = "Update WKEINSTE Set ErrPrint = true "
        gdApp.Execute sSQL, dbFailOnError
        gbErrPrint = True
    ElseIf opt1(13).value = True Then
        sSQL = "Update WKEINSTE Set ErrPrint = false "
        gdApp.Execute sSQL, dbFailOnError
        gbErrPrint = False
    End If
    
    If opt1(24).value = True Then
        sSQL = "Update WKEINSTE Set EtiFokEan = true "
        gdApp.Execute sSQL, dbFailOnError
        gbEtiFokEan = True
    ElseIf opt1(25).value = True Then
        sSQL = "Update WKEINSTE Set EtiFokEan = false "
        gdApp.Execute sSQL, dbFailOnError
        gbEtiFokEan = False
    End If
    
    If Check77.value = vbChecked Then
        sSQL = "Update WKEINSTE Set EtiQuickScanM = true "
        gdApp.Execute sSQL, dbFailOnError
        gbEtiQuickScanM = True
    Else
        sSQL = "Update WKEINSTE Set EtiQuickScanM = false "
        gdApp.Execute sSQL, dbFailOnError
        gbEtiQuickScanM = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherErrDruck"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherWeEinzelFokus()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Option3(2).value = True Then
    
        sSQL = "Update WKEINSTE Set WeEinzFo = 'EAN' "
        gdApp.Execute sSQL, dbFailOnError
        gsWeEinzFo = "EAN"
        
    ElseIf Option3(3).value = True Then
    
        sSQL = "Update WKEINSTE Set WeEinzFo = 'LS' "
        gdApp.Execute sSQL, dbFailOnError
        gsWeEinzFo = "LS"
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherWeEinzelFokus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherKaMail()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Text8.Text <> "" Then
        sSQL = "Update WKEINSTE Set KAMAIL = '" & Trim(Text8.Text) & "'"
        gdApp.Execute sSQL, dbFailOnError
        gsKaMail = Trim(Text8.Text)
    Else
        sSQL = "Update WKEINSTE Set KAMAIL = '' "
        gdApp.Execute sSQL, dbFailOnError
        gsKaMail = ""
    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherKaMail"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub speicherWeMenge()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Text1(17).Text <> "" Then
        If IsNumeric(Trim(Text1(17).Text)) Then
    
            sSQL = "Update WKEINSTE Set WeEinzMe = " & Trim(Text1(17).Text)
            gdApp.Execute sSQL, dbFailOnError
            gsWeEinzMe = Trim(Text1(17).Text)
        Else
            sSQL = "Update WKEINSTE Set WeEinzMe = '' "
            gdApp.Execute sSQL, dbFailOnError
            gsWeEinzMe = ""
        
        End If
        
    Else
    
        sSQL = "Update WKEINSTE Set WeEinzMe = '' "
        gdApp.Execute sSQL, dbFailOnError
        gsWeEinzMe = ""
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    
       
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherWeMenge"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    
End Sub
Private Sub speicherZBonWKL53()
    On Error GoTo LOKAL_ERROR
    
    If opt1(9).value = True Then
        speicherZbon "Listendrucker"

        gsZBon = "Listendrucker"
    ElseIf opt1(8).value = True Then
        speicherZbon "Bondrucker"

        gsZBon = "Bondrucker"
    End If
    
    If opt1(10).value = True Then
        speicherZählBeleg "Listendrucker"

        gsZählbeleg = "Listendrucker"
    ElseIf opt1(1).value = True Then
        speicherZählBeleg "Bondrucker"

        gsZählbeleg = "Bondrucker"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherZBonWKL53"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speichersSpanne()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If opt1(3).value = True Then
        sSQL = "Update DBEINSTE Set Spanne = 'LEK' "
        gdBase.Execute sSQL, dbFailOnError
        gsSpanne = "LEK"
        
        If opt1(22).value = True Then
            sSQL = "Update KASSEIN Set EKMAX = true "
            gdBase.Execute sSQL, dbFailOnError
            gbEKMAX = True
        ElseIf opt1(22).value = False Then
            sSQL = "Update KASSEIN Set EKMAX = false "
            gdBase.Execute sSQL, dbFailOnError
            gbEKMAX = False
        End If
        
    ElseIf opt1(2).value = True Then
        sSQL = "Update DBEINSTE Set Spanne = 'SEK' "
        gdBase.Execute sSQL, dbFailOnError
        gsSpanne = "SEK"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherWeMenge"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherSpezRunden()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check80.value = vbChecked Then
        sSQL = "Update DBEINSTE Set SPEZRU = true"
        gdBase.Execute sSQL, dbFailOnError
        gbSPEZRU = True
        
        If opt1(15).value = True Then
        
            sSQL = "Update DBEINSTE Set SPEZVAR = 1"
            gdBase.Execute sSQL, dbFailOnError
            gbSPEZVAR = 1
        
        ElseIf opt1(16).value = True Then
        
            sSQL = "Update DBEINSTE Set SPEZVAR = 2"
            gdBase.Execute sSQL, dbFailOnError
            gbSPEZVAR = 2
            
        ElseIf opt1(21).value = True Then
        
            sSQL = "Update DBEINSTE Set SPEZVAR = 3"
            gdBase.Execute sSQL, dbFailOnError
            gbSPEZVAR = 3
            
        ElseIf opt1(4).value = True Then
        
            sSQL = "Update DBEINSTE Set SPEZVAR = 4"
            gdBase.Execute sSQL, dbFailOnError
            gbSPEZVAR = 4
        
        End If
        
    ElseIf Check80.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set SPEZRU = False"
        gdBase.Execute sSQL, dbFailOnError
        gbSPEZRU = False
        
    End If
    
    
    
    
    If Check52.value = vbChecked Then
        If Val(Text1(23).Text) = 0 Then Text1(23).Text = "2"
    
        sSQL = "Update DBEINSTE Set REGEB = true"
        gdBase.Execute sSQL, dbFailOnError
        gbREGEB = True
        
        sSQL = "Update DBEINSTE Set GEBTAGE = " & Text1(23).Text
        gdBase.Execute sSQL, dbFailOnError
        giGebTage = Val(Text1(23).Text)
        
        If Check69.value = vbChecked Then
            sSQL = "Update DBEINSTE Set GebAdresse = true"
            gdBase.Execute sSQL, dbFailOnError
            gbGebAdresse = True
            
        ElseIf Check69.value = vbUnchecked Then
            sSQL = "Update DBEINSTE Set GebAdresse = False"
            gdBase.Execute sSQL, dbFailOnError
            gbGebAdresse = False
        End If
        
    ElseIf Check52.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set REGEB = False"
        gdBase.Execute sSQL, dbFailOnError
        gbREGEB = False
        
        sSQL = "Update DBEINSTE Set GEBTAGE = 2 "
        gdBase.Execute sSQL, dbFailOnError
        giGebTage = 2
        
        sSQL = "Update DBEINSTE Set GebAdresse = False"
        gdBase.Execute sSQL, dbFailOnError
        gbGebAdresse = False
    End If
    
    If Check56.value = vbChecked Then
        sSQL = "Update DBEINSTE Set GESEK = true"
        gdBase.Execute sSQL, dbFailOnError
        gbGesEKWert_anzeigen = True
        
    ElseIf Check56.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set GESEK = False"
        gdBase.Execute sSQL, dbFailOnError
        gbGesEKWert_anzeigen = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherSpezRunden"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherOptimierteStamdatenpflege()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check91.value = vbChecked Then
        sSQL = "Update WKEINSTE Set OptiStada = true"
        gdApp.Execute sSQL, dbFailOnError
        gbOptiStada = True
    ElseIf Check91.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set OptiStada = False"
        gdApp.Execute sSQL, dbFailOnError
        gbOptiStada = False
    End If
    
    If Check40.value = vbChecked Then
        sSQL = "Update WKEINSTE Set OptiStadaSpiel = true"
        gdApp.Execute sSQL, dbFailOnError
        gbOptiStadaSpiel = True
    ElseIf Check40.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set OptiStadaSpiel = False"
        gdApp.Execute sSQL, dbFailOnError
        gbOptiStadaSpiel = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherOptimierteStamdatenpflege"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherÜberwachung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check55.value = vbChecked Then
        sSQL = "Update WKEINSTE Set SPY = true"
        gdApp.Execute sSQL, dbFailOnError
        gbSPY = True
    ElseIf Check55.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set SPY = False"
        gdApp.Execute sSQL, dbFailOnError
        gbSPY = False
    End If
    
    If Text2(8).Text <> "" Then
        sSQL = "Update WKEINSTE Set IPADRESS = '" & Text2(8).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        gsServerIP = Text2(8).Text
    Else
        sSQL = "Update WKEINSTE Set IPADRESS = ''"
        gdApp.Execute sSQL, dbFailOnError
        gsServerIP = ""
    End If
    
    If Text2(9).Text <> "" Then
        sSQL = "Update WKEINSTE Set PORT = '" & Text2(9).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        gsServerPort = Text2(9).Text
    Else
        sSQL = "Update WKEINSTE Set PORT = ''"
        gdApp.Execute sSQL, dbFailOnError
        gsServerPort = ""
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherÜberwachung"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherAuto_Export_Artikelbestand()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check58.value = vbChecked Then
        sSQL = "Update WKEINSTE Set AEA = true"
        gdApp.Execute sSQL, dbFailOnError
        gbAuto_Export_Artikelbestand = True
    ElseIf Check58.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set AEA = False"
        gdApp.Execute sSQL, dbFailOnError
        gbAuto_Export_Artikelbestand = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAuto_Export_Artikelbestand"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherECASH(sVertragspartner As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    If cboECASH.Visible = True Then
        Select Case sVertragspartner
            Case Is = "ADT Wellcom GmbH"
            
                loeschapp "ADTOPT"
                CreateTable "ADTOPT", gdApp
            
                    If Check17.value = vbChecked Then
                        gbADTDI = True
                    Else
                        gbADTDI = False
                    End If

                    If Check18.value = vbChecked Then
                        gbADTVI = True
                    Else
                        gbADTVI = False
                    End If

                    If Check19.value = vbChecked Then
                        gbADTAE = True
                    Else
                        gbADTAE = False
                    End If

                    If Check20.value = vbChecked Then
                        gbADTEU = True
                    Else
                        gbADTEU = False
                    End If
                    
                    If Text9.Text <> "" Then
                        gADTclientId = Text9.Text
                    Else
                        gADTclientId = "0"
                    End If
                    
                    If Text14.Text <> "" Then
                        gADTLimit = Text14.Text
                    Else
                        gADTLimit = 0
                    End If
                    
                    'neu 19.08.2010
                    If Text18.Text <> "" Then
                        gADTipAdress = Text18.Text
                    Else
                        gADTipAdress = ""
                    End If
                    
                    If Text17.Text <> "" Then
                        gADTport = Text17.Text
                    Else
                        gADTport = ""
                    End If
                    'ende neu
                    
                    sSQL = "Insert into ADTOPT (Verfahren,termID "
                    sSQL = sSQL & " , VISA "
                    sSQL = sSQL & " , DINERS "
                    sSQL = sSQL & " , AMEX "
                    sSQL = sSQL & " , EURO "
                    sSQL = sSQL & " , clientId "
                    sSQL = sSQL & " , Limit "
                    sSQL = sSQL & " , IPADRESS "
                    sSQL = sSQL & " , PORT "
                    sSQL = sSQL & " ) "
                    sSQL = sSQL & " values ('XML', '" & Text12.Text & "'"
                    
                    If gbADTVI Then
                        sSQL = sSQL & " , True "
                    Else
                        sSQL = sSQL & " , False "
                    End If
                    
                    If gbADTDI Then
                        sSQL = sSQL & " , True "
                    Else
                        sSQL = sSQL & " , False "
                    End If
                
                    If gbADTAE Then
                        sSQL = sSQL & " , True "
                    Else
                        sSQL = sSQL & " , False "
                    End If
                    
                    If gbADTEU Then
                        sSQL = sSQL & " , True "
                    Else
                        sSQL = sSQL & " , False "
                    End If
                    
                    sSQL = sSQL & " , '" & gADTclientId & "' "
                    sSQL = sSQL & " , " & gADTLimit
                    sSQL = sSQL & " , '" & gADTipAdress & "' "
                    sSQL = sSQL & " , '" & gADTport & "' "
                    sSQL = sSQL & " ) "
                    gdApp.Execute sSQL, dbFailOnError
                
                    leseadtopt
                
                
                sSQL = "Update WKEINSTE Set EPartner = 'ADT' "
                gdApp.Execute sSQL, dbFailOnError
    
                gsEPartner = "ADT"
            
            
                sSQL = "Update WKEINSTE Set ECASH = true "
                gdApp.Execute sSQL, dbFailOnError
    
                gbEcash = True
                
            Case Is = "elPAY"
                
                sSQL = "Update WKEINSTE Set EPartner = 'ELP' "
                gdApp.Execute sSQL, dbFailOnError
    
                gsEPartner = "ELP"
            
                sSQL = "Update WKEINSTE Set ECASH = true "
                gdApp.Execute sSQL, dbFailOnError
    
                gbEcash = True
                
            Case Is = "ZVT"
                
                sSQL = "Update WKEINSTE Set EPartner = 'ZVT' "
                gdApp.Execute sSQL, dbFailOnError
    
                gsEPartner = "ZVT"
            
                sSQL = "Update WKEINSTE Set ECASH = true "
                gdApp.Execute sSQL, dbFailOnError
    
                gbEcash = True
            Case Is = "ZV2"
                
                sSQL = "Update WKEINSTE Set EPartner = 'ZV2' "
                gdApp.Execute sSQL, dbFailOnError
    
                gsEPartner = "ZV2"
            
                sSQL = "Update WKEINSTE Set ECASH = true "
                gdApp.Execute sSQL, dbFailOnError
    
                gbEcash = True
            
            Case Else
                sSQL = "Update WKEINSTE Set EPartner = '', ECASH = false "
                gdApp.Execute sSQL, dbFailOnError
                
                loeschNEW "ELPOPT", gdApp
            
                gsEPartner = ""
                gbEcash = False
                Exit Sub
        End Select
       
    Else
        sSQL = "Update WKEINSTE Set EPartner = '', ECASH = false "
        gdApp.Execute sSQL, dbFailOnError
    
        loeschNEW "ELPOPT", gdApp
        loeschNEW "ZVTOPT", gdApp
    
        gsEPartner = ""
        gbEcash = False
    End If

    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherECASH"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

        Fehlermeldung1
    
End Sub
Private Sub speicherStatistik()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
            
    If Check10.value = vbChecked Then
    
        sSQL = "Update DBEINSTE Set ustatw = true "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Delete from Statist where  art = 'W'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Statist (Art,KissKundnr,KissZusatz,LastAusw,permail)"
        sSQL = sSQL & " Values "
        sSQL = sSQL & " ('W','" & Trim(Text10.Text) & "','" & Trim(Text20.Text) & "', Datevalue('" & Text11.Text & "')"
        
        If Check59.value = vbChecked Then
            gbStatweekperMail = True
            sSQL = sSQL & " , true  "
        Else
            gbStatweekperMail = False
            sSQL = sSQL & " , False  "
        End If
        
        sSQL = sSQL & ")"
        
        gdBase.Execute sSQL, dbFailOnError
        
        gsStatkundnr = Trim(Text10.Text)
        gsStatZusatz = Trim(Text20.Text)
        gdateStatlast = Text11.Text
        gbUnistatWeek = True
        
    ElseIf Check10.value = vbUnchecked Then
    
        sSQL = "Update DBEINSTE Set ustatw = false "
        gdBase.Execute sSQL, dbFailOnError

        gbUnistatWeek = False
        gsStatkundnr = ""
        gsStatZusatz = ""
        gdateStatlast = 0
        
        gbStatweekperMail = False
        
    End If
    
    If Check87.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ustatm = true "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Delete from Statist where  art = 'M'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Statist (Art,KissKundnr,LastAusw)"
        sSQL = sSQL & " Values "
        sSQL = sSQL & " ('M','" & Trim(Text15.Text) & "', Datevalue('" & Text16.Text & "'))"
        gdBase.Execute sSQL, dbFailOnError
        
        gsMStatkundnr = Trim(Text15.Text)
        gdateMStatlast = Text16.Text
        gbUnistatMonat = True
        
    ElseIf Check87.value = vbUnchecked Then
    
        sSQL = "Update DBEINSTE Set ustatm = false "
        gdBase.Execute sSQL, dbFailOnError

        gbUnistatMonat = False
        gsMStatkundnr = ""
        gdateMStatlast = 0
        
    End If
    
    If Text19.Text <> "" Then
    
        sSQL = "Update DBEINSTE Set KUPFAD = '" & Text19.Text & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        gsKUPFAD = Text19.Text
    End If
            
            
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherStatistik"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

    Fehlermeldung1

End Sub
Private Sub speicherFILBONI()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
            
    
    If Check64.value = vbChecked Then
        sSQL = "Update DBEINSTE Set AA = true "
        gdBase.Execute sSQL, dbFailOnError
        gbAA = True
        
    ElseIf Check64.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set AA = false "
        gdBase.Execute sSQL, dbFailOnError

        gbAA = False
    End If
    
    If Check76.value = vbChecked Then
        sSQL = "Update DBEINSTE Set tagAkt = true "
        gdBase.Execute sSQL, dbFailOnError
        gbTagAkt = True
        
    ElseIf Check76.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set tagAkt = false "
        gdBase.Execute sSQL, dbFailOnError

        gbTagAkt = False
    End If
            
            
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherFILBONI"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

    Fehlermeldung1

End Sub
Private Sub speicherECTOZ()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
            
    If Check39.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ECTOZ = true "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        gbECTOZ = True
        
    ElseIf Check39.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set ECTOZ = false "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

        gbECTOZ = False
    End If
            
            
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherECTOZ"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "

    Fehlermeldung1

End Sub
Private Sub speicherMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If opt1(7).value = True Then
    
        If Trim(cbocom.Text) = "" Then
            cbocom.Text = "1"
        Else
            If Not IsNumeric(Trim(cbocom.Text)) Then
                cbocom.Text = "1"
            End If
        End If
        gbYtescanPcom = CByte(Trim(cbocom.Text))
        
        
        If Trim(Text1(12).Text) = "" Then
            Text1(12).Text = "60"
        Else
            If Not IsNumeric(Trim(Text1(12).Text)) Then
                Text1(12).Text = "60"
            End If
        End If
        giMDEPAUSE = CInt(Trim(Text1(12).Text))
        
        gbYtescanPcom = CByte(Trim(cbocom.Text))
        
        
        sSQL = "Update WKEINSTE Set MDEGER = 'SCANPAL' "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDECOM = " & gbYtescanPcom
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDEPAUSE = " & giMDEPAUSE
        gdApp.Execute sSQL, dbFailOnError
        

        schreib232_read (gbYtescanPcom)
        gsMDEGERAET = "SCANPAL"
        
        
        
        
        
    ElseIf opt1(26).value = True Then
    
        If Trim(cbocom.Text) = "" Then
            cbocom.Text = "1"
        Else
            If Not IsNumeric(Trim(cbocom.Text)) Then
                cbocom.Text = "1"
            End If
        End If
        gbYtescanPcom = CByte(Trim(cbocom.Text))
        
        
        If Trim(Text1(12).Text) = "" Then
            Text1(12).Text = "60"
        Else
            If Not IsNumeric(Trim(Text1(12).Text)) Then
                Text1(12).Text = "60"
            End If
        End If
        giMDEPAUSE = CInt(Trim(Text1(12).Text))
        
        gbYtescanPcom = CByte(Trim(cbocom.Text))
        
        
        sSQL = "Update WKEINSTE Set MDEGER = 'CIPHERLAB' "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDECOM = " & gbYtescanPcom
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDEPAUSE = " & giMDEPAUSE
        gdApp.Execute sSQL, dbFailOnError
        

        schreibData_read (gbYtescanPcom)
        gsMDEGERAET = "CIPHERLAB"
        
        
        
        
        
        
        
    ElseIf opt1(6).value = True Then
    
        If Trim(cbocom.Text) = "" Then
            cbocom.Text = "1"
        Else
            If Not IsNumeric(Trim(cbocom.Text)) Then
                cbocom.Text = "1"
            End If
        End If
        gbYtescanPcom = CByte(Trim(cbocom.Text))
    
    
        If Trim(Text1(12).Text) = "" Then
            Text1(12).Text = "60"
        Else
            If Not IsNumeric(Trim(Text1(12).Text)) Then
                Text1(12).Text = "60"
            End If
        End If
        giMDEPAUSE = CInt(Trim(Text1(12).Text))
    
        sSQL = "Update WKEINSTE Set MDEGER = 'FORCOM' "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDECOM = " & gbYtescanPcom
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDEPAUSE = " & giMDEPAUSE
        gdApp.Execute sSQL, dbFailOnError
        
        gsMDEGERAET = "FORCOM"
        
    ElseIf opt1(0).value = True Then 'Rewe-MDE
    
        If Trim(cbocom.Text) = "" Then
            cbocom.Text = "1"
        Else
            If Not IsNumeric(Trim(cbocom.Text)) Then
                cbocom.Text = "1"
            End If
        End If
        gbYtescanPcom = CByte(Trim(cbocom.Text))
    
    
        If Trim(Text1(12).Text) = "" Then
            Text1(12).Text = "60"
        Else
            If Not IsNumeric(Trim(Text1(12).Text)) Then
                Text1(12).Text = "60"
            End If
        End If
        giMDEPAUSE = CInt(Trim(Text1(12).Text))
    
        sSQL = "Update WKEINSTE Set MDEGER = 'REWEMDE' "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDECOM = " & gbYtescanPcom
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDEPAUSE = " & giMDEPAUSE
        gdApp.Execute sSQL, dbFailOnError
        
        gsMDEGERAET = "REWEMDE"
        
    ElseIf opt1(11).value = True Then 'Bela-MDE
    
        If Trim(cbocom.Text) = "" Then
            cbocom.Text = "1"
        Else
            If Not IsNumeric(Trim(cbocom.Text)) Then
                cbocom.Text = "1"
            End If
        End If
        gbYtescanPcom = CByte(Trim(cbocom.Text))
    
    
        If Trim(Text1(12).Text) = "" Then
            Text1(12).Text = "60"
        Else
            If Not IsNumeric(Trim(Text1(12).Text)) Then
                Text1(12).Text = "60"
            End If
        End If
        giMDEPAUSE = CInt(Trim(Text1(12).Text))
    
        sSQL = "Update WKEINSTE Set MDEGER = 'BELAMDE' "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDECOM = " & gbYtescanPcom
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDEPAUSE = " & giMDEPAUSE
        gdApp.Execute sSQL, dbFailOnError
        
        gsMDEGERAET = "BELAMDE"
    ElseIf opt1(12).value = True Then 'Casio-MDE
    
        If Trim(cbocom.Text) = "" Then
            cbocom.Text = "1"
        Else
            If Not IsNumeric(Trim(cbocom.Text)) Then
                cbocom.Text = "1"
            End If
        End If
        gbYtescanPcom = CByte(Trim(cbocom.Text))
    
    
        If Trim(Text1(12).Text) = "" Then
            Text1(12).Text = "60"
        Else
            If Not IsNumeric(Trim(Text1(12).Text)) Then
                Text1(12).Text = "60"
            End If
        End If
        giMDEPAUSE = CInt(Trim(Text1(12).Text))
    
        sSQL = "Update WKEINSTE Set MDEGER = 'CASIOMDE' "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDECOM = " & gbYtescanPcom
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set MDEPAUSE = " & giMDEPAUSE
        gdApp.Execute sSQL, dbFailOnError
        
        gsMDEGERAET = "CASIOMDE"
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherMDE"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    
End Sub
Private Sub speicherWaage()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    gbYteWAAGEPcom = CByte(Trim(Combo4.Text))
    gsWAAGE = Trim(Combo5.Text)
    
    sSQL = "Update WKEINSTE Set Waage =  '" & gsWAAGE & "'"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update WKEINSTE Set WAAGECOM = " & gbYteWAAGEPcom
    gdApp.Execute sSQL, dbFailOnError
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherwaage"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherdfu()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    gsDFU = Trim(Combo6.Text)
    
    If gsDFU = "" Then gsDFU = "keine DFÜ vorhanden"
    
    sSQL = "Update WKEINSTE Set DFU =  '" & gsDFU & "'"
    gdApp.Execute sSQL, dbFailOnError

    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherdfu"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    
End Sub
Private Sub speicheretisort()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Option5(0).value = True Then
        giSortierung = 0
    ElseIf Option5(1).value = True Then
        giSortierung = 1
    ElseIf Option5(2).value = True Then
        giSortierung = 2
    ElseIf Option5(3).value = True Then
        giSortierung = 3
    ElseIf Option5(4).value = True Then
        giSortierung = 4
    End If
    
    sSQL = "Update WKEINSTE Set etisort = " & giSortierung
    gdApp.Execute sSQL, dbFailOnError
        
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicheretisort"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    
End Sub


Private Sub speicherNoSpruch()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check60.value = vbChecked Then
        sSQL = "Update DBEINSTE Set NoSpruch = true"
        gdBase.Execute sSQL, dbFailOnError
        gbNoSpruch = True
    ElseIf Check60.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set NoSpruch = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNoSpruch = False
    End If
    
    If Check78.value = vbChecked Then
        sSQL = "Update WKEINSTE Set SOUND = true"
        gdApp.Execute sSQL, dbFailOnError
        gbSound = True
    ElseIf Check78.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set SOUND = False"
        gdApp.Execute sSQL, dbFailOnError
        gbSound = False
    End If
    
    If Check81.value = vbChecked Then
        sSQL = "Update WKEINSTE Set ISDEMO = true"
        gdApp.Execute sSQL, dbFailOnError
        gbISDEMO = True
    ElseIf Check81.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set ISDEMO = False"
        gdApp.Execute sSQL, dbFailOnError
        gbISDEMO = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherNoSpruch"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherBedienerKarte()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Not SpalteInTabellegefundenNEW("WKEINSTE", "BEDKARTE", gdApp) Then
        SpalteAnfuegenNEW "WKEINSTE", "BEDKARTE", "BIT", gdApp
        
        sSQL = "Update WKEINSTE Set BEDKARTE = False "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("WKEINSTE", "QPASS", gdApp) Then
        SpalteAnfuegenNEW "WKEINSTE", "QPASS", "BIT", gdApp
        
        sSQL = "Update WKEINSTE Set QPASS = False "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("DBEINSTE", "BEDKARTE", gdBase) Then
        SpalteAnfuegenNEW "DBEINSTE", "BEDKARTE", "BIT", gdBase
        
        sSQL = "Update DBEINSTE Set BEDKARTE = False"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    If Check12.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BEDKARTE = True"
        gdApp.Execute sSQL, dbFailOnError
        
        If Check11.value = vbChecked Then
            sSQL = "Update DBEINSTE Set BEDKARTE = True"
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "Update DBEINSTE Set BEDKARTE = False"
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        gbBEDKARTE = True
    Else
        sSQL = "Update WKEINSTE Set BEDKARTE = False"
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update DBEINSTE Set BEDKARTE = False"
        gdBase.Execute sSQL, dbFailOnError
        
        gbBEDKARTE = False
    End If
    
    If Check23.value = vbChecked Then
        sSQL = "Update WKEINSTE Set QPASS = True"
        gdApp.Execute sSQL, dbFailOnError
        
        gbQPASS = True
    Else
        sSQL = "Update WKEINSTE Set QPASS = False"
        gdApp.Execute sSQL, dbFailOnError
        
        gbQPASS = False
    End If
    
    If Check73.value = vbChecked Then
        sSQL = "Update WKEINSTE Set GTBON = True"
        gdApp.Execute sSQL, dbFailOnError
        
        gbGTBON = True
    Else
        sSQL = "Update WKEINSTE Set GTBON = False"
        gdApp.Execute sSQL, dbFailOnError
        
        gbGTBON = False
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBedienerKarte"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherBargeldEingabe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Not SpalteInTabellegefundenNEW("WKEINSTE", "BAREIN", gdApp) Then
        SpalteAnfuegenNEW "WKEINSTE", "BAREIN", "BIT", gdApp
        
        sSQL = "Update WKEINSTE Set BAREIN = False "
        gdApp.Execute sSQL, dbFailOnError
    End If
    If Check22.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BAREIN = True"
        gdApp.Execute sSQL, dbFailOnError
        
        gbBargeldEingabe = True
    Else
        sSQL = "Update WKEINSTE Set BAREIN = False"
        gdApp.Execute sSQL, dbFailOnError
        
        gbBargeldEingabe = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBargeldEingabe"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherAbNummer()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Not SpalteInTabellegefundenNEW("DBEINSTE", "ABda", gdBase) Then
        SpalteAnfuegenNEW "DBEINSTE", "ABda", "BIT", gdBase
        sSQL = "Update DBEINSTE Set ABda = False"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("DBEINSTE", "ABNR", gdBase) Then
        SpalteAnfuegenNEW "DBEINSTE", "ABNr", "BIT", gdBase
        sSQL = "Update DBEINSTE Set ABNr = False"
        gdBase.Execute sSQL, dbFailOnError
    End If

    If Check13.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ABNR = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAbschlussNummer = True
    ElseIf Check13.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set ABNR = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAbschlussNummer = False
    End If
    
    If Check14.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ABDa = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAbschlussDatum = True
    ElseIf Check14.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set ABDa = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAbschlussDatum = False
    End If
    
    If Check83.value = vbChecked Then
        sSQL = "Update DBEINSTE Set AGNAUSW = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAGNAusw = True
    ElseIf Check83.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set AGNAUSW = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAGNAusw = False
    End If
    
    If Check93.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ARTKUMWGN = True"
        gdBase.Execute sSQL, dbFailOnError
        gbARTKUMWGN = True
    ElseIf Check93.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set ARTKUMWGN = False"
        gdBase.Execute sSQL, dbFailOnError
        gbARTKUMWGN = False
    End If
    
    If Check62.value = vbChecked Then
        sSQL = "Update DBEINSTE Set KUMSUM = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUMSUM = True
    ElseIf Check62.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set KUMSUM = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUMSUM = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAbNummer"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherAUSBLEND()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Datendrin("KASSEIN", gdBase) = False Then
        sSQL = "Insert into KASSEIN  (NEUKUNDEN) values (False)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    If Check109(0).value = vbChecked Then
        sSQL = "Update KASSEIN Set NEUKUNDEN = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNeukunden = True
    ElseIf Check109(0).value = vbUnchecked Then
        sSQL = "Update KASSEIN Set NEUKUNDEN = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNeukunden = False
    End If
    
    
    
    If Check109(2).value = vbChecked Then
        sSQL = "Update KASSEIN Set TPBF = True"
        gdBase.Execute sSQL, dbFailOnError
        gbTPbf = True
    ElseIf Check109(2).value = vbUnchecked Then
        sSQL = "Update KASSEIN Set TPBF = False"
        gdBase.Execute sSQL, dbFailOnError
        gbTPbf = False
    End If
    
    If Check108.value = vbChecked Then
        sSQL = "Update KASSEIN Set STERNE = True"
        gdBase.Execute sSQL, dbFailOnError
        gbSterne = True
    ElseIf Check108.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set STERNE = False"
        gdBase.Execute sSQL, dbFailOnError
        gbSterne = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAUSBLEND"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherFarbebeiPark()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update KASSEIN Set FarbebeiPark = 0 "
    gdBase.Execute sSQL, dbFailOnError
    giFarbebeiPark = 0
    
    If Text23.Text <> "" Then
        If IsNumeric(Text23.Text) Then
            giFarbebeiPark = CInt(Text23.Text)
            sSQL = "Update KASSEIN Set FarbebeiPark = " & giFarbebeiPark & " "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherFarbebeiPark"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub speicherAliasFil()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update KASSEIN Set FILALI = 0 "
    gdBase.Execute sSQL, dbFailOnError
    giFILALI = 0
    
    If Text2(6).Text <> "" Then
        If IsNumeric(Text2(6).Text) Then
            sSQL = "Update KASSEIN Set FILALI = " & Text2(6).Text
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            giFILALI = CInt(Text2(6).Text)
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAliasFil"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub speicherQZBON()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check25.value = vbChecked Then
        sSQL = "Update WKEINSTE Set QZBON = true"
        gdApp.Execute sSQL, dbFailOnError
        gbQZBON = True
    ElseIf Check25.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set QZBON = False"
        gdApp.Execute sSQL, dbFailOnError
        gbQZBON = False
    End If
    
    If Check57.value = vbChecked Then
        sSQL = "Update WKEINSTE Set MITEXPORT = true "
        gdApp.Execute sSQL, dbFailOnError
        gbMitExport = True
    ElseIf Check57.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set MITEXPORT = false "
        gdApp.Execute sSQL, dbFailOnError
        gbMitExport = False
    End If
    
    
    
    
    If chk_ZBON_DINA4_HOCH.value = vbChecked Then
        sSQL = "Update WKEINSTE Set ZBONDINA4HOCH = true "
        gdApp.Execute sSQL, dbFailOnError
        gbZBONDINA4HOCH = True
    ElseIf chk_ZBON_DINA4_HOCH.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set ZBONDINA4HOCH = false "
        gdApp.Execute sSQL, dbFailOnError
        gbZBONDINA4HOCH = False
    End If
    
    
    
    
    
    If gbQZBON Then
        speicherAbschlussdetails
    Else
        loeschNEW "ABREPORT", gdApp
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherQZBON"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub


Private Sub speicherKSF()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sStart As String
    
    If Check47.value = vbChecked Then
        sSQL = "Update WKEINSTE Set KSF = true"
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE set Kassdatstart  = '' "
        gdApp.Execute sSQL, dbFailOnError
        
        gsKassDatstart = ""
        
        gbKSF = True
        
    ElseIf Check47.value = vbUnchecked Then
    
        sSQL = "Update WKEINSTE Set KSF = False"
        gdApp.Execute sSQL, dbFailOnError
        
        sStart = Right(DTPicker2.value, 8)
        
        If sStart <> "" Then
            gsKassDatstart = sStart
            sSQL = "Update WKEINSTE set Kassdatstart  = '" & sStart & "'"
            gdApp.Execute sSQL, dbFailOnError
        Else
            gsKassDatstart = ""
        End If
        
        gbKSF = False
        
    End If
    
    If Check79.value = vbChecked Then
        sSQL = "Update WKEINSTE Set AABSCHL = true"
        gdApp.Execute sSQL, dbFailOnError
    
        gbAABSCHL = True
    ElseIf Check79.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set AABSCHL = False"
        gdApp.Execute sSQL, dbFailOnError
        
        gbAABSCHL = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherKSF"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherKissLive()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_DSN", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_DSN", "TEXT(50)", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_ADRESSE", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_ADRESSE", "TEXT(50)", gdApp
        SpalteAnfuegenNEW "WKEINSTE", "KL_BENUTZER", "TEXT(50)", gdApp
        SpalteAnfuegenNEW "WKEINSTE", "KL_PASSWORT", "TEXT(20)", gdApp
        SpalteAnfuegenNEW "WKEINSTE", "KL_DATENBANKNAME", "TEXT(20)", gdApp
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVEBESTAND", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_LIVEKVKPR", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVEKVKPR", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_LIVEGUTSCHEIN", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVEGUTSCHEIN", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_LIVEFARBE", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVEFARBE", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_LIVEGEFSPERR", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVEGEFSPERR", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_LIVEBESTAND_DIFF", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVEBESTAND_DIFF", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "KL_LIVENACHRICHTEN", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "KL_LIVENACHRICHTEN", "BIT", gdApp
    End If
    
    Dim bSpeichernErlaubt As Boolean
    bSpeichernErlaubt = False
    
    If Text21(6).Text <> "" Then
        If Text21(1).Text <> "" Or Text21(2).Text <> "" Then
            bSpeichernErlaubt = True
        End If
    Else
        If Text21(0).Text <> "" Or Text21(1).Text <> "" Or Text21(2).Text <> "" Or Text21(3).Text <> "" Then
            bSpeichernErlaubt = True
        End If

    End If
    
    If bSpeichernErlaubt = True Then
    
        sSQL = "Update WKEINSTE Set KL_DSN = '" & Text21(6).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        
        gsKL_DSN = Text21(6).Text
        
        sSQL = "Update WKEINSTE Set KL_ADRESSE = '" & Text21(0).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        
        gsKL_ADRESSE = Text21(0).Text
        
        sSQL = "Update WKEINSTE Set KL_BENUTZER = '" & Text21(1).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        
        gsKL_BENUTZER = Text21(1).Text
        
        sSQL = "Update WKEINSTE Set KL_PASSWORT = '" & Text21(2).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        
        gsKL_PASSWORT = Text21(2).Text
        
        sSQL = "Update WKEINSTE Set KL_DATENBANKNAME = '" & Text21(3).Text & "'"
        gdApp.Execute sSQL, dbFailOnError
        
        gsKL_DATENBANKNAME = Text21(3).Text
        
        If Check29.value = vbChecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEBESTAND = true"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEBESTAND = True
            
        ElseIf Check29.value = vbUnchecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEBESTAND = False"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEBESTAND = False
            
        End If
        
        If Check86.value = vbChecked Then
            sSQL = "Update WKEINSTE Set KL_LIVENACHRICHTEN = true"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVENACHRICHTEN = True
            
        ElseIf Check86.value = vbUnchecked Then
            sSQL = "Update WKEINSTE Set KL_LIVENACHRICHTEN = False"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVENACHRICHTEN = False
            
        End If
        
        If Check26.value = vbChecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEKVKPR = true"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEKVKPR = True
            
        ElseIf Check26.value = vbUnchecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEKVKPR = False"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEKVKPR = False
            
        End If
        
        If Check46.value = vbChecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEFarbe = true"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEFarbe = True
            
        ElseIf Check46.value = vbUnchecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEFarbe = False"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEFarbe = False
            
        End If
        
        If Check38.value = vbChecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEgefSperr = true"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEGefSperr = True
            
        ElseIf Check38.value = vbUnchecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEgefSperr = False"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEGefSperr = False
            
        End If
        
        If Check3.value = vbChecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEGUTSCHEIN = true"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEGUTSCHEIN = True
            
        ElseIf Check3.value = vbUnchecked Then
            sSQL = "Update WKEINSTE Set KL_LIVEGUTSCHEIN = False"
            gdApp.Execute sSQL, dbFailOnError
            gbKL_LIVEGUTSCHEIN = False
            
        End If
    Else
        gsKL_DSN = ""
        gsKL_ADRESSE = ""
        gsKL_BENUTZER = ""
        gsKL_PASSWORT = ""
        gsKL_DATENBANKNAME = ""
        gbKL_LIVEBESTAND = False
        gbKL_LIVEKVKPR = False
        gbKL_LIVEGUTSCHEIN = False
        gbKL_LIVEFarbe = False
        gbKL_LIVEGefSperr = False
        gbKL_LIVENACHRICHTEN = False
        
        sSQL = "Update WKEINSTE Set KL_LIVEBESTAND = FALSE "
        gdApp.Execute sSQL, dbFailOnError
            
        sSQL = "Update WKEINSTE Set KL_LIVEKVKPR = FALSE "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set KL_LIVEGUTSCHEIN = FALSE "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set KL_LIVEFarbe = FALSE "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set KL_LIVENachrichten = FALSE "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set KL_LIVEGefSperr = FALSE "
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set KL_ADRESSE = '' "
        sSQL = sSQL & ", KL_BENUTZER = '' "
        sSQL = sSQL & ", KL_PASSWORT = '' "
        sSQL = sSQL & ", KL_DATENBANKNAME = '' "
        sSQL = sSQL & ", KL_DSN = '' "
        gdApp.Execute sSQL, dbFailOnError
        
    End If
    
    
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherKissLive"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    
End Sub
Private Sub speicherWebshop()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Datendrin("WEBSHOP_E", gdBase) = False Then
        sSQL = "Insert into WEBSHOP_E  (MySQL_PHP_SCRIPT_PFAD) "
        sSQL = sSQL & " values ('') "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    If Text21(11).Text <> "" Then
        
        sSQL = "Update WEBSHOP_E Set MySQL_PHP_SCRIPT_PFAD = '" & Text21(11).Text & "'"
        gdBase.Execute sSQL, dbFailOnError

        gsMySQL_PHP_SCRIPT_PFAD = Text21(11).Text

        If Check71.value = vbChecked Then
            
            
            If Text21(8).Text <> "" Or Text21(9).Text <> "" Or Text21(10).Text <> "" Then
            
                sSQL = "Update WEBSHOP_E Set MySQL_BESTAND_TAB = '" & Text21(8).Text & "'"
                gdBase.Execute sSQL, dbFailOnError
        
                gsMySQL_BESTAND_TAB = Text21(8).Text
        
                sSQL = "Update WEBSHOP_E Set MySQL_BESTAND_INDEXSPALTE = '" & Text21(9).Text & "'"
                gdBase.Execute sSQL, dbFailOnError
        
                gsMySQL_BESTAND_INDEXSPALTE = Text21(9).Text
        
                sSQL = "Update WEBSHOP_E Set MySQL_BESTAND_SPALTE = '" & Text21(10).Text & "'"
                gdBase.Execute sSQL, dbFailOnError
        
                gsMySQL_BESTAND_SPALTE = Text21(10).Text
                
                sSQL = "Update WEBSHOP_E Set MySQL_LIVEBESTAND = true"
                gdBase.Execute sSQL, dbFailOnError
                gbMySQL_LIVEBESTAND = True
            
            Else
                gsMySQL_BESTAND_TAB = ""
                gsMySQL_BESTAND_INDEXSPALTE = ""
                gsMySQL_BESTAND_SPALTE = ""
                gbMySQL_LIVEBESTAND = False
        
                sSQL = "Update WEBSHOP_E Set MySQL_LIVEBESTAND = FALSE "
                gdBase.Execute sSQL, dbFailOnError
        
                sSQL = "Update WEBSHOP_E Set MySQL_BESTAND_TAB = '' "
                sSQL = sSQL & ", MySQL_BESTAND_INDEXSPALTE = '' "
                sSQL = sSQL & ", MySQL_BESTAND_SPALTE= '' "
                gdBase.Execute sSQL, dbFailOnError
        
            End If
    
            
        ElseIf Check71.value = vbUnchecked Then
            sSQL = "Update WEBSHOP_E Set MySQL_LIVEBESTAND = False"
            gdBase.Execute sSQL, dbFailOnError
            gbMySQL_LIVEBESTAND = False

        End If
    Else
    
        gbMySQL_LIVEBESTAND = False
        gsMySQL_PHP_SCRIPT_PFAD = ""

        sSQL = "Update WEBSHOP_E Set MySQL_LIVEBESTAND = FALSE "
        gdBase.Execute sSQL, dbFailOnError

        sSQL = "Update WEBSHOP_E Set MySQL_PHP_SCRIPT_PFAD = '' "
        gdBase.Execute sSQL, dbFailOnError

    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherWebshop"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    Resume Next
End Sub

Private Sub speicherBILDTAST()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "BILDTAST", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "BILDTAST", "BIT", gdApp
    End If
    
    If Check32.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BILDTAST = true"
        gdApp.Execute sSQL, dbFailOnError
        gbBILDTAST = True
        
    ElseIf Check32.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BILDTAST = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBILDTAST = False
        
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherBILDTAST"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
    
End Sub


Private Sub speicherscanmodi()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check15.value = vbChecked Then
        sSQL = "Update wkEINSTE Set scanmodi = True"
        gdApp.Execute sSQL, dbFailOnError
        gbscanmodi = True
        
    ElseIf Check15.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set scanmodi = False"
        gdApp.Execute sSQL, dbFailOnError
        gbscanmodi = False
        
    End If
    
    If Check21.value = vbChecked Then
        sSQL = "Update DBEINSTE Set NONEGZU = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNONEGZU = True
        
    ElseIf Check21.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set NONEGZU = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNONEGZU = False
        
    End If
    
    If Check51.value = vbChecked Then
        sSQL = "Update DBEINSTE Set WEautoGef = True"
        gdBase.Execute sSQL, dbFailOnError
        gbWEautoGef = True
        
    ElseIf Check51.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set WEautoGef = False"
        gdBase.Execute sSQL, dbFailOnError
        gbWEautoGef = False
        
    End If
    
    If Check63.value = vbChecked Then
        sSQL = "Update DBEINSTE Set AutoZwsp = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAutoZwsp = True
        
    ElseIf Check63.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set AutoZwsp = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAutoZwsp = False
        
    End If
    
    If Check33.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ETIONLYME = True"
        gdBase.Execute sSQL, dbFailOnError
        gbETIONLYME = True
        
    ElseIf Check33.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set ETIONLYME = False"
        gdBase.Execute sSQL, dbFailOnError
        gbETIONLYME = False
        
    End If
    
    
    If Check50.value = vbChecked Then
        sSQL = "Update DBEINSTE Set NoETIWeAusBe = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNoETIWeAusBe = True
        
    ElseIf Check50.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set NoETIWeAusBe = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNoETIWeAusBe = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherscanmodi"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
        
        Fehlermeldung1
End Sub
Private Sub speicherArtNrBeg()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update WKEINSTE Set ArtNrBeg = '" & Trim(Text1(9).Text) & "' "
    gdApp.Execute sSQL, dbFailOnError
    
    glArtNrBeg = Trim(Text1(9).Text)
    
    If Check65.value = vbChecked Then
        sSQL = "Update DBEINSTE Set NewArt = True"
        gdBase.Execute sSQL, dbFailOnError
        
        gbNewArt = True
    Else
        sSQL = "Update DBEINSTE Set NewArt = False"
        gdBase.Execute sSQL, dbFailOnError
        
        gbNewArt = False
    End If
    
    If Check53.value = vbChecked Then
        sSQL = "Update DBEINSTE Set NewArtNrVorschlag = True"
        gdBase.Execute sSQL, dbFailOnError
        
        gbNewArtNrVorschlag = True
    Else
        sSQL = "Update DBEINSTE Set NewArtNrVorschlag = False"
        gdBase.Execute sSQL, dbFailOnError
        
        gbNewArtNrVorschlag = False
    End If
    
    If Check66.value = vbChecked Then
        sSQL = "Update DBEINSTE Set ARTEINDEUT = True"
        gdBase.Execute sSQL, dbFailOnError
        
        gbArtEindeut = True
    Else
        sSQL = "Update DBEINSTE Set ARTEINDEUT = False"
        gdBase.Execute sSQL, dbFailOnError
        
        gbArtEindeut = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherArtNrBeg"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherEdekaLief()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update DBEINSTE Set Edeka = '" & Trim(Text1(26).Text) & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    gsEdeka = Trim(Text1(26).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherEdekaLief"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherMWSTBeg()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update DBEINSTE Set MWSTBeg = '" & Trim(Text1(19).Text) & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    gsMWST = Trim(Text1(19).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherMWSTBeg"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherRundung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cWert As String
    
    cWert = IIf(Text1(5).Text <> "", (Text1(5).Text), "0")
    sSQL = "Update DBEINSTE Set Aufrunden = '" & cWert & "'"
    gdBase.Execute sSQL, dbFailOnError
    giAufrunden = CInt(cWert)
    
    cWert = IIf(Text1(6).Text <> "", (Text1(6).Text), "0")
    sSQL = "Update DBEINSTE Set Abrunden = '" & cWert & "'"
    gdBase.Execute sSQL, dbFailOnError
    giAbrunden = CInt(cWert)
    
    cWert = IIf(Text1(7).Text <> "", (Text1(7).Text), "0")
    sSQL = "Update DBEINSTE Set RundKrit = '" & cWert & "'"
    gdBase.Execute sSQL, dbFailOnError
    giRundkrit = CInt(cWert)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherRundung"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(index As Integer)
    On Error GoTo LOKAL_ERROR

    If index = 21 Then
        Dim vMonday As Date
        Dim vSunday As Date
        
        If IsNumeric(Text1(21).Text) = True Then
        
            vMonday = GetDateFromWeek(CInt(Text1(21).Text), vbMonday, CLng(Text1(20).Text))
            vSunday = GetDateFromWeek(CInt(Text1(21).Text), vbSunday, CLng(Text1(20).Text))
            lbl6(81).Caption = "KW = " & Text1(21).Text & " (" & vMonday & "-" & vSunday & ")"
        Else
        
            lbl6(81).Caption = ""
        End If
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = glSelBack1
    Text1(index).SelStart = Len(Text1(index).Text)
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
   
End Sub
Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If index = 13 Then
        If KeyCode = vbKeyReturn Then
            Command1_Click 4
        End If
    ElseIf index = 21 Then
        If KeyCode = vbKeyReturn Then
            Command30_Click
        End If
    End If
    
     Exit Sub
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Text13_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text13(index).BackColor = glSelBack1
    Text13(index).SelStart = Len(Text13(index).Text)
    
     Exit Sub
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text13_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub



Private Sub Text21_GotFocus(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Text21(index).BackColor = glSelBack1
    Text21(index).SelStart = Len(Text21(index).Text)
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text21_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text21_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case index
    Case Is = 4, 5
        cValid = "1234567890" & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    End Select
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text21_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text21_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Text21(index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text21_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Text22_Change()
On Error GoTo LOKAL_ERROR

    If Text22.Text = "a0872cc5" Then
        Command33.Enabled = True
    Else
        Command33.Enabled = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text22_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub



Private Sub Text9_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        If cZeichen = "," Then
            If InStr(Text9.Text, cZeichen) > 0 Then
                KeyAscii = 0
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text9_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        If cZeichen = "," Then
            If InStr(Text14.Text, cZeichen) > 0 Then
                KeyAscii = 0
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text14_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Text2_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(index).BackColor = glSelBack1
    Text2(index).SelStart = Len(Text2(index).Text)
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub


Private Sub txtsicherpfad_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    txtSicherPfad.BackColor = glSelBack1
    txtSicherPfad.SelStart = Len(txtSicherPfad.Text)
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtsicherpfad_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
   
End Sub

Private Sub Text4_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text4.BackColor = glSelBack1
    Text4.SelStart = Len(Text4.Text)
    
     Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
   
End Sub
Private Sub Text5_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text5.BackColor = glSelBack1
    Text5.SelStart = Len(Text4.Text)
    
     Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
   
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case index
    Case Is = 11, 12, 15, 21, 20, 22, 25
        cValid = "1234567890" & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Case Is = 10, 13, 14
        cValid = "1234567890," & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Case Is = 16
        cValid = "1234567890D" & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Case Is = 18
        cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
        cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
        cValid = cValid & "+äÄÜüÖöß"
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Case Is = 19
        cValid = "VEO" & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Case Is = 24
        cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
        cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
        cValid = cValid & "+-.@"
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Case Else
        cValid = "1234567890," & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            If cZeichen = "," Then
                If InStr(Text1(index).Text, cZeichen) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    End Select
    
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
   
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        If cZeichen = "," Then
            If InStr(Text4.Text, cZeichen) > 0 Then
                KeyAscii = 0
            End If
        End If
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
   
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        If cZeichen = "," Then
            If InStr(Text5.Text, cZeichen) > 0 Then
                KeyAscii = 0
            End If
        End If
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text4_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text4.BackColor = vbWhite
    
        Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub txtsicherpfad_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtSicherPfad.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtsicherpfad_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text5_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text5.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub Text2_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(index).BackColor = vbWhite
    
        Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = vbWhite
    
        Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text13_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text13(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text13_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub txtStatus_Change()
On Error GoTo LOKAL_ERROR

    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

End Sub
