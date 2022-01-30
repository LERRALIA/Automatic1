VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL163 
   Caption         =   "Detailinformationen"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "frmWKL163.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11760
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
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
      Height          =   9945
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame3 
         Caption         =   "Konfiguration"
         Height          =   7215
         Left            =   1320
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
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
            Height          =   690
            Index           =   6
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertikal
            TabIndex        =   75
            Top             =   2880
            Width           =   3015
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Standardwarnhinweis verwenden"
            Height          =   495
            Left            =   240
            TabIndex        =   74
            Top             =   2400
            Width           =   2895
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
            Height          =   270
            Index           =   5
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   72
            Top             =   2040
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "beim Export immer regulären Preis verwenden"
            Height          =   495
            Left            =   240
            TabIndex        =   42
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Bei höherem LVK: ""Sonderpreis - Text"" in die Kurzbeschreibung schreiben"
            Height          =   615
            Left            =   240
            TabIndex        =   38
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Lieferantenbestellnummer an die Artikelbezeichnung anhängen"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   3135
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   12
            Left            =   1800
            TabIndex        =   35
            Top             =   6720
            Width           =   1575
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C000&
            Caption         =   "Versand"
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
            Left            =   240
            TabIndex        =   73
            Top             =   2040
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Exportartikel"
         Height          =   10935
         Left            =   -1560
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   12135
         Begin VB.Frame Frame2 
            Caption         =   "Welcher Shop?"
            Height          =   2775
            Left            =   9480
            TabIndex        =   76
            Top             =   5040
            Width           =   1815
            Begin VB.OptionButton Option3 
               Caption         =   "WooCommerce"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   85
               Top             =   1680
               Width           =   1455
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Hitmeister"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   84
               Top             =   1320
               Width           =   1455
            End
            Begin VB.OptionButton Option3 
               Caption         =   "xt commerce V3"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   80
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton Option3 
               Caption         =   "xt commerce"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton Option3 
               Caption         =   "t-online Shop"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   78
               Top             =   960
               Value           =   -1  'True
               Width           =   1455
            End
            Begin sevCommand3.Command Command4 
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   77
               Top             =   2280
               Width           =   1575
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
               Caption         =   "Exportieren"
               PictureAlign    =   2
               PictureVisible  =   0   'False
               Version3        =   -1  'True
            End
         End
         Begin VB.FileListBox File3 
            Height          =   285
            Left            =   10920
            TabIndex        =   69
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   7
            Left            =   10560
            MouseIcon       =   "frmWKL163.frx":0442
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   68
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   6
            Left            =   10080
            MouseIcon       =   "frmWKL163.frx":074C
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   67
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   5
            Left            =   9600
            MouseIcon       =   "frmWKL163.frx":0A56
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   66
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   4
            Left            =   9120
            MouseIcon       =   "frmWKL163.frx":0D60
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   65
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   3
            Left            =   8640
            MouseIcon       =   "frmWKL163.frx":106A
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   64
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   2
            Left            =   8160
            MouseIcon       =   "frmWKL163.frx":1374
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   63
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   1
            Left            =   7680
            MouseIcon       =   "frmWKL163.frx":167E
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   62
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Index           =   0
            Left            =   7200
            MouseIcon       =   "frmWKL163.frx":1988
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   375
            TabIndex        =   61
            Top             =   840
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'Kein
            FontTransparent =   0   'False
            Height          =   615
            Left            =   7200
            MouseIcon       =   "frmWKL163.frx":1C92
            MousePointer    =   99  'Benutzerdefiniert
            ScaleHeight     =   615
            ScaleWidth      =   855
            TabIndex        =   49
            Top             =   1560
            Width           =   855
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
            Height          =   285
            Index           =   4
            Left            =   240
            MaxLength       =   10
            TabIndex        =   46
            Text            =   "Text3"
            Top             =   7440
            Width           =   975
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   45
            Top             =   7440
            Width           =   375
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
            Caption         =   "F2"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6825
            Left            =   120
            MultiSelect     =   2  'Erweitert
            TabIndex        =   44
            Top             =   240
            Width           =   6615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   52
            Top             =   7320
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Bildgröße"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   51
            Top             =   6360
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   50
            Top             =   6840
            Width           =   1935
         End
         Begin VB.Image Image3 
            Height          =   300
            Left            =   7320
            Top             =   480
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C000&
            Caption         =   "Lieferant"
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
            TabIndex        =   47
            Top             =   7200
            Width           =   855
         End
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   11160
         TabIndex        =   53
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Left            =   8280
         MouseIcon       =   "frmWKL163.frx":1F9C
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   855
         TabIndex        =   19
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   40
         Top             =   6720
         Width           =   3015
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   39
         Top             =   6435
         Width           =   1575
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
         Caption         =   "Kat löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   13
         Left            =   9600
         TabIndex        =   36
         Top             =   7920
         Width           =   375
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
         Caption         =   "K"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   8280
         TabIndex        =   26
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton Option1 
            Caption         =   "nach Lieferanten"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   600
            Value           =   -1  'True
            Width           =   1815
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
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
            Caption         =   "Drucken"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option1 
            Caption         =   "nach Kategorie"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   29
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C000&
            Caption         =   "sortiert nach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2295
         End
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   5
         Left            =   8280
         TabIndex        =   25
         Top             =   7440
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "Entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   24
         Top             =   6435
         Width           =   1575
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
         Caption         =   "Kat löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   22
         Top             =   6720
         Width           =   3015
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   3
         Left            =   8280
         TabIndex        =   20
         Top             =   2640
         Width           =   2295
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
         Caption         =   "Bild laden..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   7440
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   120
         MaxLength       =   100
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   2160
         Width           =   8055
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   4080
         Width           =   8055
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   0
         Left            =   10080
         TabIndex        =   8
         Top             =   7440
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   2
         Left            =   10080
         TabIndex        =   4
         Top             =   7920
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         MaxLength       =   255
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   3120
         Width           =   8055
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   6
         Left            =   10080
         TabIndex        =   48
         Top             =   1800
         Width           =   1575
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
         Caption         =   "Exportieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   0
         Left            =   8280
         MouseIcon       =   "frmWKL163.frx":22A6
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   21
         Top             =   3120
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   1
         Left            =   9120
         MouseIcon       =   "frmWKL163.frx":25B0
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   54
         Top             =   3120
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   2
         Left            =   9480
         MouseIcon       =   "frmWKL163.frx":28BA
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   55
         Top             =   3120
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   3
         Left            =   9840
         MouseIcon       =   "frmWKL163.frx":2BC4
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   56
         Top             =   3120
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   4
         Left            =   10200
         MouseIcon       =   "frmWKL163.frx":2ECE
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   57
         Top             =   3120
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   5
         Left            =   10560
         MouseIcon       =   "frmWKL163.frx":31D8
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   58
         Top             =   3120
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   6
         Left            =   10920
         MouseIcon       =   "frmWKL163.frx":34E2
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   59
         Top             =   3120
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   615
         Index           =   7
         Left            =   11280
         MouseIcon       =   "frmWKL163.frx":37EC
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   60
         Top             =   3120
         Width           =   255
      End
      Begin sevCommand3.Command Command4 
         Height          =   210
         Index           =   15
         Left            =   10440
         TabIndex        =   71
         Top             =   3840
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "Bild löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   81
         Top             =   1800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   82
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   83
         Top             =   3720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "vom Lieferant vorhanden"
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
         Index           =   13
         Left            =   720
         TabIndex        =   70
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   7320
         Top             =   840
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "2. Kategorie"
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
         Index           =   17
         Left            =   3480
         TabIndex        =   41
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefNr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   3840
         TabIndex        =   33
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefBestNr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   15
         Left            =   3840
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Bildname"
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
         Index           =   12
         Left            =   8280
         TabIndex        =   31
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "1. Kategorie"
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
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   300
         Left            =   7320
         Top             =   480
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   7320
         Top             =   240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "regulärer Preis"
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
         Index           =   9
         Left            =   3480
         TabIndex        =   18
         Top             =   7560
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "regulärer Preis"
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
         Index           =   8
         Left            =   3480
         TabIndex        =   17
         Top             =   7200
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Shop Preis"
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
         Index           =   5
         Left            =   5880
         TabIndex        =   16
         Top             =   7200
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Anzahl Zeichen"
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
         Index           =   4
         Left            =   5760
         TabIndex        =   14
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Anzahl Zeichen"
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
         Index           =   3
         Left            =   5760
         TabIndex        =   13
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelbeschreibung - bis zu 100 Zeichen"
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
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblanzeige 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   7920
         Width           =   8895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Artnr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeich"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   7335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Kurzbeschreibung - bis zu 255 Zeichen"
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
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Shop Informationen"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   6135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "ausführliche Beschreibung"
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
         Index           =   6
         Left            =   600
         TabIndex        =   5
         Top             =   3600
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmWKL163"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check4_Click()
On Error GoTo LOKAL_ERROR

    If Check4.Value = vbChecked Then
        Text3(6).Enabled = True
    Else
        Text3(6).Enabled = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4_Click"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    WKL163Positionieren

    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Dim i As Integer
    
    Text3(0).Text = ""
    Text3(1).Text = ""
    Text3(2).Text = ""
    Text3(3).Text = ""
    Text3(4).Text = ""
    
    Label5(9).Caption = Format(ermKVKPR1(gsARTNR), "####0.00")
    Text3(3).Text = Label5(9).Caption
    Text3(4).Text = ermLINR(gsARTNR)
    Label5(1).Caption = gsARTNR
    Label5(16).Caption = ermLibesnr(gsARTNR, CLng(Text3(4).Text))
    Label5(7).Caption = fnArtBezSuchen(gsARTNR)
    For i = 0 To 3
        Label5(7).Caption = SwapStr(Label5(7).Caption, "  ", " ")
    Next i
    
    'erst combo füllen
    Combofuellen_TabX "INTERART", Combo1, "KATEGORIE", "bitte wählen"
    
    'erst combo füllen
    Combofuellen_TabX "INTERART", Combo2, "KATEGORIE2", "bitte wählen"
    
    'dann kategorie
    anzeigenInterart gsARTNR
    
    If Externe_Beschreibung_vorhanden(Label5(16).Caption) = True Then
        Label5(13).Visible = True
    Else
        Label5(13).Visible = False
    End If
    
    ZeigeBilder gsARTNR, 0
    
    If NewTableSuchenDBKombi("E163", gdBase) Then
        If SpalteInTabellegefundenNEW("E163", "ShopArt", gdBase) = False Then
            SpalteAnfuegenNEW "E163", "ShopArt", "INTEGER", gdBase
        End If
        voreinstellungladen163
    End If
    
    Frame1.Caption = ermAnzahlShopArt & " Shop-Artikel"
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Externe_Beschreibung_vorhanden(slibesnr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lLiefnr     As Long
    Dim sLiefVerz   As String
    Dim cPfad1      As String
    Dim sTextPfad   As String
    Dim rsZ         As Recordset
    
    Externe_Beschreibung_vorhanden = False
    
    cPfad1 = gcDBPfad    'Dabapfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    '**************Gibt es eine Exceltabelle mit Beschreibungen?
    lLiefnr = ermLINR(Label5(1).Caption)

    sLiefVerz = lLiefnr & Trim(ermLiefBez(lLiefnr))

    VerzVorhanden sLiefVerz, cPfad1 & "Picture\"

    sTextPfad = cPfad1 & "Picture\" & sLiefVerz & "\"

    If FileExists(sTextPfad & "TEXTE.mdb") Then

        'wenn da, dann reinschauen
        
        Dim dbText As Database
        Set dbText = OpenDatabase(sTextPfad & "TEXTE.mdb", False)
        Set rsZ = dbText.OpenRecordset("BESCHREIBUNG")

        If Not rsZ.EOF Then
            rsZ.MoveFirst
            Do While Not rsZ.EOF
                If Not IsNull(rsZ!LIBESNR) Then
                    If Trim(rsZ!LIBESNR) = Trim(slibesnr) Then
                        Externe_Beschreibung_vorhanden = True
                    End If
                End If
                rsZ.MoveNext
            Loop
        End If
        rsZ.Close
        dbText.Close
    End If
    '****************Exceltabelle Ende
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Externe_Beschreibung_vorhanden"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Function Externe_Beschreibung(slibesnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lLiefnr     As Long
    Dim sLiefVerz   As String
    Dim cPfad1      As String
    Dim sTextPfad   As String
    Dim rsZ         As Recordset
    
    Externe_Beschreibung = ""
    
    cPfad1 = gcDBPfad    'Dabapfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    '**************Gibt es eine Exceltabelle mit Beschreibungen?
    lLiefnr = ermLINR(Label5(1).Caption)

    sLiefVerz = lLiefnr & Trim(ermLiefBez(lLiefnr))

    VerzVorhanden sLiefVerz, cPfad1 & "Picture\"

    sTextPfad = cPfad1 & "Picture\" & sLiefVerz & "\"

    If FileExists(sTextPfad & "TEXTE.mdb") Then

        'wenn da, dann reinschauen
        
        Dim dbText As Database
        
        Set dbText = OpenDatabase(sTextPfad & "TEXTE.mdb", False)
        
        Set rsZ = dbText.OpenRecordset("BESCHREIBUNG")

        If Not rsZ.EOF Then
            rsZ.MoveFirst
            Do While Not rsZ.EOF
                
                If Not IsNull(rsZ!LIBESNR) Then
                    If Trim(rsZ!LIBESNR) = Trim(slibesnr) Then
                        If Not IsNull(rsZ!ARTIKELBESCHREIBUNG) Then
                            Externe_Beschreibung = Trim(rsZ!ARTIKELBESCHREIBUNG)
                        End If
                    End If
                End If
                rsZ.MoveNext
            Loop
        End If
        rsZ.Close
        dbText.Close
    End If
    '****************Exceltabelle Ende
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Externe_Beschreibung"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub voreinstellungspeichern163()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    Dim bo0         As Integer
    Dim bo1         As Integer
    Dim bo2         As Integer
    Dim bo3         As Integer
    Dim sVersand    As String
    Dim sSWarnText  As String
    Dim iShopart    As Integer
    
    sVersand = Text3(5).Text
    sSWarnText = Text3(6).Text
    
    loeschNEW "E163", gdBase
    CreateTableT2 "E163", gdBase
    
    If Check1.Value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    If Check2.Value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    If Check3.Value = vbChecked Then
        bo2 = 0
    Else
        bo2 = -1
    End If
    
    If Check4.Value = vbChecked Then
        bo3 = 0
    Else
        bo3 = -1
    End If
    
    If Option3(0).Value = True Then
        iShopart = 0
    ElseIf Option3(1).Value = True Then
        iShopart = 1
    End If
    
    sSQL = "Insert into E163 (bo0,bo1,bo2,bo3,Versand,SWARNTEXT,Shopart) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & ",'" & sVersand & "','" & sSWarnText & "', " & iShopart & ""
   
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern163"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen163()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim iShopart As Integer
    
    Set rs = gdBase.OpenRecordset("E163")
    If Not rs.EOF Then
        
        If rs!bo0 = True Then
            Check1.Value = vbUnchecked
        Else
            Check1.Value = vbChecked
        End If
        
        
        If rs!bo1 = True Then
            Check2.Value = vbUnchecked
        Else
            Check2.Value = vbChecked
        End If
        
        If rs!bo2 = True Then
            Check3.Value = vbUnchecked
        Else
            Check3.Value = vbChecked
        End If
        
        If Not IsNull(rs!Versand) Then
            Text3(5).Text = rs!Versand
        Else
            Text3(5).Text = ""
        End If
        
        If Not IsNull(rs!SWarntext) Then
            Text3(6).Text = rs!SWarntext
        Else
            Text3(6).Text = ""
        End If
        
        If rs!bo3 = True Then
            Check4.Value = vbUnchecked
        Else
            Check4.Value = vbChecked
        End If
        
        If Not IsNull(rs!Shopart) Then
            iShopart = Val(rs!Shopart)
        Else
            iShopart = -1
        End If
        
        If iShopart = 0 Then
            Option3(0).Value = True
        ElseIf iShopart = 1 Then
            Option3(1).Value = True
        Else
            Option3(0).Value = True
        End If
    
    End If
    rs.Close: Set rs = Nothing
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen163"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL163Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame3.Top = 120
    Frame3.Left = 8280
    Frame3.Height = 7215
    Frame3.Width = 3495
    
    Frame4.Top = 0
    Frame4.Left = 0
    Frame4.Height = 7860
    Frame4.Width = 11775
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL163Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Artikel fehlt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub hinzufuegen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cArtNr As String
    Dim cKURZBESCHREIB As String
    Dim cBESCHREIB As String
    Dim cKAT1 As String
    Dim cKAT2 As String
    Dim cArtBez As String
    Dim dSHOPKVK As Double
    
    cArtNr = Label5(1).Caption
    
    If cArtNr = "" Then
        Unload frmWKL163
        Exit Sub
    End If
    
    cArtBez = Trim(Text3(2).Text)
    cArtBez = SwapStr(cArtBez, "'", "$")
    cArtBez = SwapStr(cArtBez, ",", " ")
    cArtBez = SwapStr(cArtBez, "*", " ")
    cArtBez = SwapStr(cArtBez, ";", " ")
    
    If cArtBez = "" Then
        MsgBox "Bitte geben Sie die Artikelbezeichnung ein!", vbInformation + vbOKOnly, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    cKURZBESCHREIB = Trim(Text3(0).Text)
    cKURZBESCHREIB = SwapStr(cKURZBESCHREIB, "'", "$")
    cKURZBESCHREIB = SwapStr(cKURZBESCHREIB, ",", " ")
    cKURZBESCHREIB = SwapStr(cKURZBESCHREIB, "*", " ")
    cKURZBESCHREIB = SwapStr(cKURZBESCHREIB, ";", " ")

    cBESCHREIB = Trim(Text3(1).Text)
    cBESCHREIB = SwapStr(cBESCHREIB, "'", "$")
    cBESCHREIB = SwapStr(cBESCHREIB, Chr(10), " ")
    cBESCHREIB = SwapStr(cBESCHREIB, Chr(13), " ")
    cBESCHREIB = SwapStr(cBESCHREIB, ";", " ")
    cBESCHREIB = SwapStr(cBESCHREIB, Chr(34), " ")
    cBESCHREIB = SwapStr(cBESCHREIB, ",", " ")
    cBESCHREIB = SwapStr(cBESCHREIB, "*", " ")
    
    dSHOPKVK = 0
    If Text3(3).Text <> "" Then
        If IsNumeric(Text3(3).Text) Then
            dSHOPKVK = CDbl(Text3(3).Text)
        End If
    End If
    
    If Combo1.Text <> "" Then
        If Combo1.Text <> "bitte wählen" Then
            cKAT1 = Trim(Combo1.Text)
        End If
    End If
    
    If Combo2.Text <> "" Then
        If Combo2.Text <> "bitte wählen" Then
            cKAT2 = Trim(Combo2.Text)
        End If
    End If
    
    delInterart cArtNr
    
    sSQL = "INSERT into INTERART (ARTNR,ARTBEZ,INTERBEZ,BESCHREIB,LASTDATE,SHOPKVK,KATEGORIE,KATEGORIE2,BILDgr) values "
    sSQL = sSQL & " ('" & cArtNr & "','" & cArtBez & "','" & cKURZBESCHREIB & "','" & cBESCHREIB & "', '" & DateValue(Now) & "'"
    
    sSQL = sSQL & ", '" & dSHOPKVK & "'  "
    sSQL = sSQL & ", '" & cKAT1 & "'  "
    sSQL = sSQL & ", '" & cKAT2 & "'  "
    sSQL = sSQL & ", '" & Label5(12).Caption & "'  "
    
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "erfolgreich gespeichert", lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "hinzufuegen"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub anzeigenInterart(cArtNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim dWert As Double
    
    If cArtNr = "" Then
        Exit Sub
    End If
    
    sSQL = "Select * from INTERART where ARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!ARTBEZ) Then
            Text3(2).Text = rsrs!ARTBEZ
            Text3(2).Text = SwapStr(Text3(2).Text, "$", "'")
        Else
            Text3(2).Text = ""
        End If
        
        If Not IsNull(rsrs!INTERBEZ) Then
            Text3(0).Text = rsrs!INTERBEZ
            Text3(0).Text = SwapStr(Text3(0).Text, "$", "'")
        Else
            Text3(0).Text = ""
        End If
        
        If Not IsNull(rsrs!BESCHREIB) Then
            Text3(1).Text = rsrs!BESCHREIB
            Text3(1).Text = SwapStr(Text3(1).Text, "$", "'")
        Else
            Text3(1).Text = ""
        End If
        
        If Not IsNull(rsrs!SHOPKVK) Then
            dWert = rsrs!SHOPKVK
        Else
            dWert = 0
        End If
        Text3(3).Text = Format$(dWert, "#####0.00")
        
        If Not IsNull(rsrs!KATEGORIE) Then
            If rsrs!KATEGORIE <> "" Then
                Combo1.Text = rsrs!KATEGORIE
            End If
        End If
        
        If Not IsNull(rsrs!KATEGORIE2) Then
            If rsrs!KATEGORIE2 <> "" Then
                Combo2.Text = rsrs!KATEGORIE2
            End If
        End If
        
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigenInterart"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Zeige_Shop_Artikel_inListe(Listx As ListBox, lLinr As Long)
On Error GoTo LOKAL_ERROR

Dim sSQL    As String
Dim rsrs    As Recordset
Dim cFeld   As String
Dim cSatz   As String

Listx.Clear

sSQL = "Select i.Artnr, i.Artbez, i.lastexport "
sSQL = sSQL & " from Interart i inner join Artikel a on i.Artnr = a.Artnr "
sSQL = sSQL & " where a.linr = " & lLinr & " "
sSQL = sSQL & " order by i.artnr "
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    Do While Not rsrs.EOF
    
        cSatz = ""
        cFeld = ""
        If Not IsNull(rsrs!artnr) Then
            cFeld = rsrs!artnr
            cSatz = cFeld
        End If
        
        If Not IsNull(rsrs!ARTBEZ) Then
            cFeld = rsrs!ARTBEZ
        Else
            cFeld = ""
        End If
        
        If Len(cFeld) > 35 Then
            cFeld = Left(cFeld, 32) & "..."
        End If
        
        cSatz = cSatz & " " & cFeld & Space(35 - Len(cFeld))
        
        If Not IsNull(rsrs!lastexport) Then
            cFeld = rsrs!lastexport
        Else
            cFeld = ""
        End If
        cSatz = cSatz & " " & Format(cFeld, "DD.MM.YY")
            
                
        Listx.AddItem cSatz
        
    rsrs.MoveNext
    Loop
End If
rsrs.Close: Set rsrs = Nothing


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_Shop_Artikel_inListe"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index
    Case 0
        hinzufuegen 'speichern

    Case 1
        print_Shop_Art
    Case 2
        Unload frmWKL163
    Case 3
        Pic_hinzu Label5(1).Caption
    Case 4
        Kat_Del Combo1, "Kategorie"
        Combofuellen_TabX "INTERART", Combo1, "KATEGORIE", "bitte wählen"
    Case 11
        Kat_Del Combo2, "Kategorie2"
        Combofuellen_TabX "INTERART", Combo2, "KATEGORIE2", "bitte wählen"
    Case 5
        Art_Del
        Unload frmWKL163
    Case 6
        Command4(13).Visible = False
        Frame4.Visible = True
        
        If Text3(4).Text <> "" Then
            If IsNumeric(Text3(4).Text) Then
            
                Zeige_Shop_Artikel_inListe List1, CLng(Text3(4).Text)
            End If
        End If
    Case 7
        Text3_KeyUp 4, vbKeyF2, 0
    Case 8
        If Check1.Value = vbChecked Then
            'bez + libesnr
            Text3(2).Text = Label5(7).Caption & " " & Label5(16).Caption
        Else
            Text3(2).Text = Label5(7).Caption
        End If
    Case 9
        Text3(0).Text = Text3(2).Text
        If Check2.Value = vbChecked Then
            'Sonderpreissatz
            Dim cKVK As String
            Dim cLVK As String
            cKVK = ermKVKPR1(gsARTNR)
            cLVK = ermVKPR(gsARTNR)
            If CDbl(cKVK) < CDbl(cLVK) Then
                cKVK = Format(cKVK, "####0.00") & " EUR"
                cLVK = Format(cLVK, "####0.00") & " EUR"
                Text3(0).Text = "Sonderpreis statt " & cLVK & " jetzt nur " & cKVK
            End If
        End If
    Case 10
        If Label5(13).Visible = True Then
            Text3(1).Text = Externe_Beschreibung(Label5(16).Caption)
        Else
            If Check2.Value = vbChecked Then
                Text3(1).Text = Text3(2).Text
            Else
                Text3(1).Text = Text3(0).Text
            End If
        End If
        
    Case 12
        voreinstellungspeichern163
        Frame3.Visible = False
    Case 13
        Frame3.Visible = True
    Case 14
    
        voreinstellungspeichern163
        
        'die selektierten in eine Tab
        If Select_inTab Then
            If Text3(4).Text <> "" Then
                If IsNumeric(Text3(4).Text) Then
                    If Option3(0).Value = True Then
                        CSV_Export_TonlineShop CLng(Text3(4).Text)
                    ElseIf Option3(1).Value = True Then
                        CSV_Export_xtCommerceShop CLng(Text3(4).Text)
                    ElseIf Option3(2).Value = True Then
                        CSV_Export_xtCommerceShop_v3 CLng(Text3(4).Text), "§"
                    ElseIf Option3(3).Value = True Then
                        CSV_Export_Hitmeister CLng(Text3(4).Text)
                    ElseIf Option3(4).Value = True Then
                        CSV_Export_WooCommerceShop CLng(Text3(4).Text)
                    End If
                End If
            End If
        Else
            anzeige "rot", "Bitte markieren Sie eine Zeile", lblanzeige
        End If
    Case 15

        DelBild Label5(12).Caption
        ZeigeBilder gsARTNR, 0
        
End Select
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Select_inTab() As Boolean

    Dim bFound      As Boolean
    Dim sSQL        As String
    Dim cArtNr      As String
    Dim lcount      As Long
    Dim lbildgr     As Long
    Dim sPfad       As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    lbildgr = 0
    
    loeschNEW "SHOPTEMP", gdBase
    
    sSQL = "Create Table SHOPTEMP (Artnr Long)"
    gdBase.Execute sSQL, dbFailOnError

    Select_inTab = False
    
    bFound = False
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If Not bFound Then
        Exit Function
    End If
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            cArtNr = Trim(Left(List1.list(lcount), 6))
            
            If FileExists(sPfad & "\" & cArtNr & ".jpg") Then
                lbildgr = lbildgr + fnFileSize(sPfad & "\" & cArtNr & ".jpg")
                
                sSQL = "Update Interart set bildgr = '" & cArtNr & "' & '.jpg' "
                sSQL = sSQL & " where artnr = " & cArtNr & " "
                gdBase.Execute sSQL, dbFailOnError
                
            End If
            
            sSQL = "Insert into SHOPTEMP (Artnr) values  (" & cArtNr & ")"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update Interart set lastexport = '" & DateValue(Now) & "' "
            sSQL = sSQL & " where artnr = " & cArtNr & " "
            gdBase.Execute sSQL, dbFailOnError
            
            
            Select_inTab = True
        End If
    Next
    
    Label1.Caption = Format(lbildgr / 1024, "###,##0.00") & " KB"
    Label1.Refresh
            
    Label3.Caption = Format(lbildgr / 1024 / 1024, "###,##0.00") & " MB"
    Label3.Refresh
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Select_inTab"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub CSV_Export_TonlineShop(lLinr As Long)
On Error GoTo LOKAL_ERROR


    Dim cFeld               As String
    Dim cSatz               As String
    Dim cSatz2              As String
    
    Dim cFeld_Kat           As String
    Dim cSatz_Kat           As String
    
    Dim cFeld_Kat2          As String
    Dim cSatz_Kat2          As String
    
    Dim iFileNr             As Integer
    Dim iFileNr_Kat         As Integer
    Dim rsrs                As Recordset
    Dim lPos                As Long
    Dim lPos_Kat            As Long
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    
    Dim cPfad               As String
    Dim cPfad1              As String
    Dim cdatei              As String
    Dim sAusgabedatname     As String
    
    Dim cdateiKAT           As String
    Dim sAusgabedatnameKAT  As String
    
    Dim sQuelle             As String
    Dim sZiel               As String
    
    Dim sQuellpfad          As String
    Dim sZielpfad           As String
    Dim lfail               As Long
    Dim lRet                As Long
    
    sQuellpfad = gcDBPfad
    sQuellpfad = ShortPath(sQuellpfad)
    If Right(sQuellpfad, 1) <> "\" Then
        sQuellpfad = sQuellpfad & "\"
    End If
    sQuellpfad = sQuellpfad & "PICTURE\ARTIKEL"
    
    Screen.MousePointer = 11
    
    sAusgabedatname = "Produkte.csv"
    sAusgabedatnameKAT = "Kategorie-Produkt-Zuweisung.csv"
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdateiKAT = cPfad1 & "BOX\" & sAusgabedatnameKAT
    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"
    
    VerzVorhanden "Bilder", cPfad & "\"
    
    sZielpfad = gcDBPfad
    sZielpfad = ShortPath(sZielpfad)
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "BOX\Bilder"
    
    Kill sZielpfad & "\*.*"
    Kill cdatei
    Kill cdateiKAT
    
    '***Kat_Datei
    iFileNr_Kat = FreeFile
    Open cdateiKAT For Binary As #iFileNr_Kat
    

    cFeld_Kat = "" & Chr(34) & "Kategorie [Category]" & Chr(34) & ""
    cFeld_Kat = cFeld_Kat & ";" & Chr(34) & "Produkt [Product]" & Chr(34) & ""
    cFeld_Kat = cFeld_Kat & ";" & Chr(34) & "Sortierung [Position]" & Chr(34) & ""
    
    
    cSatz_Kat = cFeld_Kat
    cSatz_Kat = cSatz_Kat & Chr$(13) & Chr$(10)
    
    lPos_Kat = LOF(iFileNr_Kat)
    lPos_Kat = lPos_Kat + 1
    Put #iFileNr_Kat, lPos_Kat, cSatz_Kat
    
    '***Kat_Datei Ende
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cFeld = "" & Chr(34) & "Typ [Class]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bezeichner [Alias]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Sortierung [Position]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Sichtbar [IsVisible]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Listenpreis/EUR/gross [ListPrices/EUR/gross]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Vergleichspreis/EUR/gross [ManufacturerPrices/EUR/gross]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Tagespreisabhängig [IsDailyPrice]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "EcoParticipationCategory [EcoParticipationCategory] " & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "EcoParticipations/EUR/gross [EcoParticipations/EUR/gross]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Steuerklasse [TaxClass]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bestelleinheit [OrderUnit]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Menge für Preis [PriceQuantity]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Mindestbestellmenge [MinOrder]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Staffelung [IntervalOrder]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Referenzeinheit [RefUnit]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Referenzmenge [RefAmount]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Menge im Produkt [RefContentAmount]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Hersteller [Manufacturer]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Hersteller-Produktnr. [ManufacturerSKU]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Produkt-Code [UPCEAN]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Lagerbestand [StockLevel]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Mindestlagerbestand [StockLevelAlert]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Lieferzeitraum [DeliveryPeriod]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Gewichtseinheit [WeightUnit]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Versandgewicht [Weight]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Länge [Length]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Höhe [Height]" & Chr(34) & ""
    
    cFeld = cFeld & ";" & Chr(34) & "Breite [Width]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Versandmethoden [ShippingMethods]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Name/Deutsch [Name/de]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Kurz-URL/Deutsch [URI/de]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Beschreibung/Deutsch [Description/de]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Erweiterte Beschreibung/Deutsch [Text/de]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Schlüsselworte/Deutsch [Keywords/de]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Neu [IsNew]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Markierung Neu bis [NewnessDate]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Käuflich [IsAvailable]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Datum für käuflich ab [AvailabilityDate]" & Chr(34) & ""

'    cFeld = cFeld & ";" & Chr(34) & "PrepaymentType [PrepaymentType]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Hinweis, wenn nicht käuflich/Deutsch [AvailabilityComment/de]" & Chr(34) & ""
'    cFeld = cFeld & ";" & Chr(34) & "PrepaymentValue [PrepaymentValue]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Produkt-Bundle [IsBundleProduct]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Hauptprodukt [SuperProduct]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Variationsattribute [SelectedVariations]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Voreingestellt [IsDefault]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Eigener Preis [HasSubOwnPrices]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bild für Listenansicht [ImageSmall]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bild für Detailansicht [ImageMedium]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bild für vergrößerte Ansicht [ImageLarge]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bild für Aktionsprodukt [ImageHotDeal]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bilder für Galerie/Diaschau [ImagesSlideShowString]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "OwnStyle [OwnStyle]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "SendDescription [SendDescription]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "LinksInNewWindow [LinksInNewWindow]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "WidgetLayout [WidgetLayout]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "WidgetLocale [WidgetLocale]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "WidgetWidth [WidgetWidth]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "WidgetHeight [WidgetHeight]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "WidgetText [WidgetText]" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Einkaufspreis [PurchasePrice]" & Chr(34) & ""
    
    cSatz = cFeld
    cSatz = cSatz & Chr$(13) & Chr$(10)
        
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "SHOPEX", gdBase
    CreateTableT2 "SHOPEX", gdBase
    
    cSQL = "Insert into SHOPEX Select "
    cSQL = cSQL & " i.ARTNR  "
    cSQL = cSQL & ", i.ARTBEZ  "
    cSQL = cSQL & ", i.INTERBEZ  "
    cSQL = cSQL & ", i.BESCHREIB  "
    cSQL = cSQL & ", i.BILDkl  "
    cSQL = cSQL & ", i.BILDmi  "
    cSQL = cSQL & ", i.BILDgr  "
    cSQL = cSQL & ", i.SHOPKVK  "
    cSQL = cSQL & ", i.KATEGORIE  "
    cSQL = cSQL & ", i.KATEGORIE2  "
    cSQL = cSQL & ", " & lLinr & " as LINR "
    cSQL = cSQL & ", 0 as Bestand "
    cSQL = cSQL & " from Interart i inner join SHOPTEMP a on i.Artnr = a.Artnr "
    gdBase.Execute cSQL, dbFailOnError
    
    If Check3.Value = vbChecked Then 'Preis, regulärer
        cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
        cSQL = cSQL & "  Set  SHOPEX.SHOPKVK = Artikel.KVKPR1 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Check4.Value = vbChecked Then 'für alle Standardwarnhinweis
        cSQL = "Update SHOPEX set Beschreib = Beschreib  + ' " & Text3(6).Text & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
    cSQL = cSQL & "  Set  SHOPEX.Bestand = Artikel.Bestand "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update SHOPEX inner join LISRT on SHOPEX.LINR = LISRT.LINR "
    cSQL = cSQL & "  Set  SHOPEX.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    cSatz_Kat2 = ""
    
    Set rsrs = gdBase.OpenRecordset("select * from SHOPEX ")
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            'CLASS
            cFeld = ""
            cSatz = cFeld & ";"
            cSatz2 = cFeld & ";"
            
            '***Kat
            If Not IsNull(rsrs!KATEGORIE) Then
                cFeld_Kat = rsrs!KATEGORIE
            Else
                cFeld_Kat = ""
            End If
            cSatz_Kat = "Categories/" & cFeld_Kat & ";"
            
            'Alias
            If Not IsNull(rsrs!artnr) Then
                cFeld_Kat = rsrs!artnr
            Else
                cFeld_Kat = "0"
            End If
            
            If cFeld_Kat = "" Then cFeld_Kat = "0"
            cSatz_Kat = cSatz_Kat & cFeld_Kat & ";"
            
            'Sortierung
            cFeld_Kat = "0"
            cSatz_Kat = cSatz_Kat & cFeld_Kat & ""
            
            '***Kat Ende
            
            '***2. Kat 2. Kat
            
            
            If Not IsNull(rsrs!KATEGORIE2) Then
                cFeld_Kat2 = rsrs!KATEGORIE2
            Else
                cFeld_Kat2 = ""
            End If
            
            If cFeld_Kat2 <> "" Then
                cSatz_Kat2 = "Categories/" & cFeld_Kat2 & ";"
                
                'Alias
                If Not IsNull(rsrs!artnr) Then
                    cFeld_Kat2 = rsrs!artnr
                Else
                    cFeld_Kat2 = "0"
                End If
                
                If cFeld_Kat2 = "" Then cFeld_Kat2 = "0"
                cSatz_Kat2 = cSatz_Kat2 & cFeld_Kat2 & ";"
                
                'Sortierung
                cFeld_Kat2 = "0"
                cSatz_Kat2 = cSatz_Kat2 & cFeld_Kat2 & ""
            Else
                cSatz_Kat2 = ""
            End If
            
            '***2. Kat 2. Kat Ende
            
            'Alias
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Position
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'IsVisible
            cFeld = "1"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            '[ListPrices/EUR/gross]
            If Not IsNull(rsrs!SHOPKVK) Then
                cFeld = rsrs!SHOPKVK
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cFeld = Format(cFeld, "####0.00")
'            cFeld = SwapStr(cFeld, ",", ".")
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
                
            
            'ManufacturerPrices/EUR/gross
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'IsDailyPrice
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'EcoParticipationCategory
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            
           
            'EcoParticipations/EUR/gross
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'TaxClass
            cFeld = "normal"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'OrderUnit
            cFeld = "piece"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'PriceQuantity
            cFeld = "1"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'MinOrder
            cFeld = "1"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'IntervalOrder
            cFeld = "1"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'RefUnit
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'RefAmount
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'RefContentAmount
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Manufacturer
            cFeld = ermLiefBez(lLinr)
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'ManufacturerSKU
            cFeld = ermLibesnr(rsrs!artnr, lLinr)
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'UPCEAN
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'StockLevel
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'StockLevelAlert
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'DeliveryPeriod
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'WeightUnit
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Weight
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Length
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Height
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Width
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'ShippingMethods
            
            
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
    
            'Name/de
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'URI/de
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Description/de
            If Not IsNull(rsrs!INTERBEZ) Then
                cFeld = rsrs!INTERBEZ
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Text/de
            If Not IsNull(rsrs!BESCHREIB) Then
                cFeld = rsrs!BESCHREIB
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Keywords/de
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'IsNew
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'NewnessDate
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'IsAvailable
            cFeld = "1"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'AvailabilityDate
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
''            'PrepaymentType
''            cFeld = ""
''            cSatz = cSatz & cFeld & ";"
''            cSatz2 = cSatz2 & cFeld & ";"
            
            'AvailabilityComment
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
''            'PrepaymentValue
''            cFeld = ""
''            cSatz = cSatz & cFeld & ";"
''            cSatz2 = cSatz2 & cFeld & ";"
            
            'IsBundleProduct
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'SuperProduct
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'SelectedVariations
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'IsDefault
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'HasSubOwnPrices
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'Bilder
            If Not IsNull(rsrs!Bildgr) Then
                sQuelle = sQuellpfad & "\" & rsrs!Bildgr
                sZiel = sZielpfad & "\" & rsrs!Bildgr
                
                lRet = CopyFile(sQuelle, sZiel, lfail)
            End If
            
            'ImageSmall
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'ImageMedium
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'ImageLarge
            cFeld = "Import/" & rsrs!Bildgr
            cSatz = cSatz & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'ImageHotDeal
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            File3.Path = sQuellpfad
            File3.Pattern = Left(rsrs!Bildgr, Len(rsrs!Bildgr) - 4) & "*.jpg"
            File3.Refresh
            
            cFeld = Chr(34)
            If File3.ListCount > 1 Then
                For i = 0 To File3.ListCount - 1
                    sBildname = File3.list(i)
                    
                    sQuelle = sQuellpfad & "\" & sBildname
                    sZiel = sZielpfad & "\" & sBildname
                    
                    lRet = CopyFile(sQuelle, sZiel, lfail)
                    
                    cFeld = cFeld & "Import/" & sBildname & ";"
                    If i = File3.ListCount - 1 Then
                        cFeld = cFeld & "Import/" & sBildname
                    End If
                Next i
            End If
            cFeld = cFeld & Chr(34)
            
            'ImagesSlideShowString
'            cFeld = "Import/" & rsrs!Bildgr
            cSatz = cSatz & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'OwnStyle
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'SendDescription
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'LinksInNewWindow
            cFeld = "0"
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'WidgetLayout
            cFeld = "" 'Centered
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'WidgetLocale
            cFeld = "" 'de_DE
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'WidgetWidth
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'WidgetHeight
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            
            'WidgetText
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            'PurchasePrice
            cFeld = ""
            cSatz = cSatz & cFeld & ";"
            cSatz2 = cSatz2 & cFeld & ";"
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            cSatz2 = cSatz2 & Chr$(13) & Chr$(10)
                    
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz2
            
            '***Kat
            cSatz_Kat = cSatz_Kat & Chr$(13) & Chr$(10)
            
                    
            lPos_Kat = LOF(iFileNr_Kat)
            lPos_Kat = lPos_Kat + 1
            Put #iFileNr_Kat, lPos_Kat, cSatz_Kat
            
            If cSatz_Kat2 <> "" Then
                cSatz_Kat2 = cSatz_Kat2 & Chr$(13) & Chr$(10)
                
                lPos_Kat = LOF(iFileNr_Kat)
                lPos_Kat = lPos_Kat + 1
                Put #iFileNr_Kat, lPos_Kat, cSatz_Kat2
            End If
            
            
            
            '***Kat Ende
            
            rsrs.MoveNext
        Loop
    End If
        
    Close iFileNr_Kat
    Close iFileNr
        
    MsgBox "Die Exportdateien finden Sie unter: " & cPfad, vbInformation, "Winkiss Hinweis:"
    
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSV_Export_TonlineShop"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub CSV_Export_WooCommerceShop(lLinr As Long)
On Error GoTo LOKAL_ERROR


    Dim cFeld               As String
    Dim cSatz               As String
    Dim cSatz2              As String
    
    Dim cFeld_Kat           As String
    Dim cSatz_Kat           As String
    
    Dim cFeld_Kat2          As String
    Dim cSatz_Kat2          As String
    
    Dim iFileNr             As Integer
    Dim iFileNr_Kat         As Integer
    Dim rsrs                As Recordset
    Dim lPos                As Long
    Dim lPos_Kat            As Long
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    
    Dim cPfad               As String
    Dim cPfad1              As String
    Dim cdatei              As String
    Dim sAusgabedatname     As String
    
    Dim cdateiKAT           As String
    Dim sAusgabedatnameKAT  As String
    
    Dim sQuelle             As String
    Dim sZiel               As String
    
    Dim sQuellpfad          As String
    Dim sZielpfad           As String
    Dim lfail               As Long
    Dim lRet                As Long
    
    sQuellpfad = gcDBPfad
    sQuellpfad = ShortPath(sQuellpfad)
    If Right(sQuellpfad, 1) <> "\" Then
        sQuellpfad = sQuellpfad & "\"
    End If
    sQuellpfad = sQuellpfad & "PICTURE\ARTIKEL"
    
    Screen.MousePointer = 11
    
    sAusgabedatname = "wc-product.csv"
'    sAusgabedatnameKAT = "Kategorie-Produkt-Zuweisung.csv"
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdateiKAT = cPfad1 & "BOX\" & sAusgabedatnameKAT
    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"
    
    VerzVorhanden "Bilder", cPfad & "\"
    
    sZielpfad = gcDBPfad
    sZielpfad = ShortPath(sZielpfad)
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "BOX\Bilder"
    
    Kill sZielpfad & "\*.*"
    Kill cdatei
    Kill cdateiKAT
    

    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cFeld = "ID"
    cFeld = cFeld & ",Typ"
    cFeld = cFeld & ",Artikelnummer"
    cFeld = cFeld & ",Name"
    cFeld = cFeld & ",Veröffentlicht"
    cFeld = cFeld & "," & Chr(34) & "Ist hervorgehoben?" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Sichtbarkeit im Katalog" & Chr(34) & ""
    cFeld = cFeld & ",Kurzbeschreibung"
    cFeld = cFeld & ",Beschreibung"
    cFeld = cFeld & "," & Chr(34) & "Datum, an dem Angebotspreis beginnt" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Datum, an dem Angebotspreis endet" & Chr(34) & ""
    cFeld = cFeld & ",Steuerstatus"
    cFeld = cFeld & ",Steuerklasse"
    cFeld = cFeld & ",Vorrätig?"
    cFeld = cFeld & ",Lager"
    cFeld = cFeld & "," & Chr(34) & "Geringe Lagermenge" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Lieferrückstande erlaubt?" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Nur einzeln verkaufen?" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Gewicht (kg)" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Länge (cm)" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Breite (cm)" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Höhe (cm)" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Kundenbewertungen erlauben?" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Hinweis zum Kauf" & Chr(34) & ""
    cFeld = cFeld & ",Angebotspreis"
    cFeld = cFeld & "," & Chr(34) & "Regulärer Preis" & Chr(34) & ""
    cFeld = cFeld & ",Kategorien"
    cFeld = cFeld & ",Schlagwörter"
    cFeld = cFeld & ",Versandklasse"
    cFeld = cFeld & ",Bilder"
    cFeld = cFeld & ",Downloadlimit"
    cFeld = cFeld & "," & Chr(34) & "Ablauftage des Downloads" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Übergeordnetes Produkt" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Gruppierte Produkte" & Chr(34) & ""
    cFeld = cFeld & ",Zusatzverkäufe"
    cFeld = cFeld & "," & Chr(34) & "Cross-Sells (Querverkäufe)" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Externe URL" & Chr(34) & ""
    cFeld = cFeld & ",Button-Text"
    cFeld = cFeld & ",Position"
    cFeld = cFeld & "," & Chr(34) & "Meta: _yoast_wpseo_primary_product_cat" & Chr(34) & ""
    cFeld = cFeld & "," & Chr(34) & "Meta: _yoast_wpseo_content_score" & Chr(34) & ""
    cSatz = cFeld
    cSatz = cSatz & Chr$(13) & Chr$(10)
        
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "SHOPEX", gdBase
    CreateTableT2 "SHOPEX", gdBase
    
    cSQL = "Insert into SHOPEX Select "
    cSQL = cSQL & " i.ARTNR  "
    cSQL = cSQL & ", i.ARTBEZ  "
    cSQL = cSQL & ", i.INTERBEZ  "
    cSQL = cSQL & ", i.BESCHREIB  "
    cSQL = cSQL & ", i.BILDkl  "
    cSQL = cSQL & ", i.BILDmi  "
    cSQL = cSQL & ", i.BILDgr  "
    cSQL = cSQL & ", i.SHOPKVK  "
    cSQL = cSQL & ", i.KATEGORIE  "
    cSQL = cSQL & ", i.KATEGORIE2  "
    cSQL = cSQL & ", " & lLinr & " as LINR "
    cSQL = cSQL & ", 0 as Bestand "
    cSQL = cSQL & ", '' as EAN "
    cSQL = cSQL & " from Interart i inner join SHOPTEMP a on i.Artnr = a.Artnr "
    gdBase.Execute cSQL, dbFailOnError
    
   
    
    If Check3.Value = vbChecked Then 'Preis, regulärer
        cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
        cSQL = cSQL & "  Set  SHOPEX.SHOPKVK = Artikel.KVKPR1 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Check4.Value = vbChecked Then 'für alle Standardwarnhinweis
        cSQL = "Update SHOPEX set Beschreib = Beschreib  + ' " & Text3(6).Text & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
     cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
    cSQL = cSQL & " Set SHOPEX.Bestand = Artikel.Bestand "
    cSQL = cSQL & " , SHOPEX.EAN = Artikel.EAN "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update SHOPEX inner join LISRT on SHOPEX.LINR = LISRT.LINR "
    cSQL = cSQL & "  Set  SHOPEX.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    Set rsrs = gdBase.OpenRecordset("select * from SHOPEX ")
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cSatz = ""
        
            'ID = ARTNR
            
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cSatz = cSatz & cFeld & ","
            
            'Typ
            cFeld = ""
            cSatz = cSatz & "simple,"
            
            
            'Artikelnummer = EAN
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ","
    
            'Name
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cSatz = cSatz & Chr(34) & cFeld & Chr(34) & ","
            
            'veröffentlicht
            cFeld = ""
            cSatz = cSatz & "1,"
            
            'ist hervorgehoben
            cFeld = ""
            cSatz = cSatz & "0,"
            
            'sichtbarkeit im Katalog
            cFeld = ""
            cSatz = cSatz & "visible,"
            
            'Kurzbeschreibung
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cSatz = cSatz & Chr(34) & cFeld & Chr(34) & ","
            
            
            'Beschreibung
            If Not IsNull(rsrs!BESCHREIB) Then
                cFeld = rsrs!BESCHREIB
            Else
                cFeld = ""
            End If
            cSatz = cSatz & Chr(34) & cFeld & Chr(34) & ","
            
            
            'Datum
            cFeld = ""
            cSatz = cSatz & ","
            
            'Datum
            cFeld = ""
            cSatz = cSatz & ","
            
            
            
            'steuerstatus
            cFeld = ""
            cSatz = cSatz & "taxable,"
            
            'steuerklasse
            cFeld = ""
            cSatz = cSatz & ","
            
            'vorrätig?
            cFeld = ""
            cSatz = cSatz & "1,"
            
            'Lager = Bestand
            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ","
            
            'geringe Lagermenge?
            cFeld = ""
            cSatz = cSatz & ","
            
'            'geringe Lagermenge?
'            cFeld = ""
'            cSatz = cSatz & "2,"
            
            'geringe Lagermenge?
            cFeld = ""
            cSatz = cSatz & "0,"
            
            'geringe Lagermenge?
            cFeld = ""
            cSatz = cSatz & "0,"
            
            'steuerklasse
            cFeld = ""
            cSatz = cSatz & ","
            'steuerklasse
            cFeld = ""
            cSatz = cSatz & ","
            'steuerklasse
            cFeld = ""
            cSatz = cSatz & ","
            'steuerklasse
            cFeld = ""
            cSatz = cSatz & ","
            
            'steuerklasse
            cFeld = ""
            cSatz = cSatz & "1,"
            
            '
            cFeld = ""
            cSatz = cSatz & ","
            
            'angebotspreis
            
            cFeld = ""
            cSatz = cSatz & ","
            
             'regulärer Preis
            If Not IsNull(rsrs!SHOPKVK) Then
                cFeld = rsrs!SHOPKVK
            Else
                cFeld = ""
            End If
            
            
            If cFeld = "" Then cFeld = "0"
            cFeld = Format(cFeld, "####0.00")
            cSatz = cSatz & Chr(34) & cFeld & Chr(34) & ","
            
            
           
            'KATEGORIE
            If Not IsNull(rsrs!KATEGORIE) Then
                cFeld = rsrs!KATEGORIE
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & ","
            
            'Schlagwörter
            cFeld = ""
            cSatz = cSatz & Chr(34) & cFeld & Chr(34) & ","
            
            cFeld = ""
            cSatz = cSatz & ","
            
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            cFeld = ""
            cSatz = cSatz & ","
            
            cFeld = ""
            cSatz = cSatz & "0,"
            
            cFeld = ""
            cSatz = cSatz & "0,"
            
            
            'das Ende ohne Komma!
            cFeld = ""
            cSatz = cSatz & "0"
            
            
            
           
            
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
                    
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz2
            
            rsrs.MoveNext
        Loop
    End If
        
    Close iFileNr_Kat
    Close iFileNr
        
    MsgBox "Die Exportdateien finden Sie unter: " & cPfad, vbInformation, "Winkiss Hinweis:"
    
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSV_Export_WooCommerceShop"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub CSV_Export_Hitmeister(lLinr As Long)
On Error GoTo LOKAL_ERROR

    Dim cFeld               As String
    Dim cSatz               As String
    Dim iFileNr             As Integer
    Dim rsrs                As Recordset
    Dim lPos                As Long
    Dim cPfad               As String
    Dim cPfad1              As String
    Dim cdatei              As String
    Dim sAusgabedatname     As String
    Dim sQuelle             As String
    Dim sZiel               As String
    Dim sQuellpfad          As String
    Dim sZielpfad           As String
    Dim lfail               As Long
    Dim lRet                As Long
    
    sQuellpfad = gcDBPfad
    sQuellpfad = ShortPath(sQuellpfad)
    If Right(sQuellpfad, 1) <> "\" Then
        sQuellpfad = sQuellpfad & "\"
    End If
    sQuellpfad = sQuellpfad & "PICTURE\ARTIKEL"
    
    Screen.MousePointer = 11
    
    sAusgabedatname = "Produkte.csv"
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"
    
    VerzVorhanden "Bilder", cPfad & "\"
    
    sZielpfad = gcDBPfad
    sZielpfad = ShortPath(sZielpfad)
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "BOX\Bilder"
    
    Kill sZielpfad & "\*.*"
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cFeld = "" & Chr(34) & "Kategorie" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "ean" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "title" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "description" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "manufacturer" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "age_rating" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "is_porn" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bildnummer/ URL" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Füllmenge" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Einheit" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "UVP" & Chr(34) & ""
    cSatz = cFeld
    cSatz = cSatz & Chr$(13) & Chr$(10)
        
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "SHOPEX", gdBase
    CreateTableT2 "SHOPEX", gdBase
    
    cSQL = "Insert into SHOPEX Select "
    cSQL = cSQL & " i.ARTNR  "
    cSQL = cSQL & ", i.ARTBEZ  "
    cSQL = cSQL & ", i.INTERBEZ  "
    cSQL = cSQL & ", i.BESCHREIB  "
    cSQL = cSQL & ", i.BILDkl  "
    cSQL = cSQL & ", i.BILDmi  "
    cSQL = cSQL & ", i.BILDgr  "
    cSQL = cSQL & ", i.SHOPKVK  "
    cSQL = cSQL & ", i.KATEGORIE  "
    cSQL = cSQL & ", i.KATEGORIE2  "
    cSQL = cSQL & ", " & lLinr & " as LINR "
    cSQL = cSQL & ", 0 as Bestand "
    cSQL = cSQL & " from Interart i inner join SHOPTEMP a on i.Artnr = a.Artnr "
    gdBase.Execute cSQL, dbFailOnError
    
    If Check3.Value = vbChecked Then 'Preis, regulärer
        cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
        cSQL = cSQL & "  Set  SHOPEX.SHOPKVK = Artikel.KVKPR1 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Check4.Value = vbChecked Then 'für alle Standardwarnhinweis
        cSQL = "Update SHOPEX set Beschreib = Beschreib  + ' " & Text3(6).Text & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
    cSQL = cSQL & "  Set  SHOPEX.Bestand = Artikel.Bestand "
    cSQL = cSQL & ", SHOPEX.EAN = Artikel.EAN "
    cSQL = cSQL & ", SHOPEX.INHALT = Artikel.INHALT "
    cSQL = cSQL & ", SHOPEX.INHALTBEZ = Artikel.INHALTBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update SHOPEX inner join LISRT on SHOPEX.LINR = LISRT.LINR "
    cSQL = cSQL & "  Set  SHOPEX.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("select * from SHOPEX ")
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cFeld = ""
            cSatz = ""
        
            'KATEGORIE
            If Not IsNull(rsrs!KATEGORIE) Then
                cFeld = rsrs!KATEGORIE
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'EAN
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Title
            If Not IsNull(rsrs!INTERBEZ) Then
                cFeld = rsrs!INTERBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Description
            If Not IsNull(rsrs!BESCHREIB) Then
                cFeld = rsrs!BESCHREIB
            Else
                cFeld = ""
            End If
            
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Manufacturer
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = rsrs!LIEFBEZ
            Else
                cFeld = ""
            End If
            
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'age_rating
            
            cFeld = "0"
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'is_porn
            
            cFeld = "0"
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Bildnummer/ URL
            If Not IsNull(rsrs!Bildgr) Then
                sQuelle = sQuellpfad & "\" & rsrs!Bildgr
                sZiel = sZielpfad & "\" & rsrs!Bildgr
                
                lRet = CopyFile(sQuelle, sZiel, lfail)
            End If
            
            'Bild
            If Not IsNull(rsrs!Bildgr) Then
                cFeld = rsrs!Bildgr
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Füllmenge
            If Not IsNull(rsrs!INHALT) Then
                cFeld = rsrs!INHALT
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Einheit
            If Not IsNull(rsrs!INHALTBEZ) Then
                cFeld = rsrs!INHALTBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'UVP
            If Not IsNull(rsrs!SHOPKVK) Then
                cFeld = rsrs!SHOPKVK
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cFeld = Format(cFeld, "####0.00")
            cFeld = SwapStr(cFeld, ",", ".")
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            
        
            
            
            
        
            
            
            
            
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            rsrs.MoveNext
        Loop
    End If
        
    Close iFileNr
        
    MsgBox "Die Exportdateien finden Sie unter: " & cPfad, vbInformation, "Winkiss Hinweis:"
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSV_Export_Hitmeister"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub CSV_Export_xtCommerceShop(lLinr As Long)
On Error GoTo LOKAL_ERROR

    Dim cFeld               As String
    Dim cSatz               As String
    Dim iFileNr             As Integer
    Dim rsrs                As Recordset
    Dim lPos                As Long
    Dim cPfad               As String
    Dim cPfad1              As String
    Dim cdatei              As String
    Dim sAusgabedatname     As String
    Dim sQuelle             As String
    Dim sZiel               As String
    Dim sQuellpfad          As String
    Dim sZielpfad           As String
    Dim lfail               As Long
    Dim lRet                As Long
    
    sQuellpfad = gcDBPfad
    sQuellpfad = ShortPath(sQuellpfad)
    If Right(sQuellpfad, 1) <> "\" Then
        sQuellpfad = sQuellpfad & "\"
    End If
    sQuellpfad = sQuellpfad & "PICTURE\ARTIKEL"
    
    Screen.MousePointer = 11
    
    sAusgabedatname = "Produkte.csv"
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"
    
    VerzVorhanden "Bilder", cPfad & "\"
    
    sZielpfad = gcDBPfad
    sZielpfad = ShortPath(sZielpfad)
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "BOX\Bilder"
    
    Kill sZielpfad & "\*.*"
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cFeld = "" & Chr(34) & "1Art Bezeichnung" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "2Untertitel" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "3Artnr" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "4Bild" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "5VkPreis" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "6Kateg1" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "7Kateg2" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "8Kateg3" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "9Beschreibung" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "bil" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "leer" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "br" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Name" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "ArtNr" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "VK" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Bezeichnung2" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Inhalt1" & Chr(34) & ""
    cFeld = cFeld & ";" & Chr(34) & "Beschreibung2" & Chr(34) & ""
    cSatz = cFeld
    cSatz = cSatz & Chr$(13) & Chr$(10)
        
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "SHOPEX", gdBase
    CreateTableT2 "SHOPEX", gdBase
    
    cSQL = "Insert into SHOPEX Select "
    cSQL = cSQL & " i.ARTNR  "
    cSQL = cSQL & ", i.ARTBEZ  "
    cSQL = cSQL & ", i.INTERBEZ  "
    cSQL = cSQL & ", i.BESCHREIB  "
    cSQL = cSQL & ", i.BILDkl  "
    cSQL = cSQL & ", i.BILDmi  "
    cSQL = cSQL & ", i.BILDgr  "
    cSQL = cSQL & ", i.SHOPKVK  "
    cSQL = cSQL & ", i.KATEGORIE  "
    cSQL = cSQL & ", i.KATEGORIE2  "
    cSQL = cSQL & ", " & lLinr & " as LINR "
    cSQL = cSQL & ", 0 as Bestand "
    cSQL = cSQL & " from Interart i inner join SHOPTEMP a on i.Artnr = a.Artnr "
    gdBase.Execute cSQL, dbFailOnError
    
    If Check3.Value = vbChecked Then 'Preis, regulärer
        cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
        cSQL = cSQL & "  Set  SHOPEX.SHOPKVK = Artikel.KVKPR1 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Check4.Value = vbChecked Then 'für alle Standardwarnhinweis
        cSQL = "Update SHOPEX set Beschreib = Beschreib  + ' " & Text3(6).Text & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
    cSQL = cSQL & "  Set  SHOPEX.Bestand = Artikel.Bestand "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update SHOPEX inner join LISRT on SHOPEX.LINR = LISRT.LINR "
    cSQL = cSQL & "  Set  SHOPEX.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("select * from SHOPEX ")
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cFeld = ""
            cSatz = ""
        
            '1Art Bezeichnung
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            '2Untertitel
            If Not IsNull(rsrs!INTERBEZ) Then
                cFeld = rsrs!INTERBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            '3Artnr
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Bilder
            If Not IsNull(rsrs!Bildgr) Then
                sQuelle = sQuellpfad & "\" & rsrs!Bildgr
                sZiel = sZielpfad & "\" & rsrs!Bildgr
                
                lRet = CopyFile(sQuelle, sZiel, lfail)
            End If
            
            '4Bild
            If Not IsNull(rsrs!Bildgr) Then
                cFeld = rsrs!Bildgr
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            '5VkPreis
            If Not IsNull(rsrs!SHOPKVK) Then
                cFeld = rsrs!SHOPKVK
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cFeld = Format(cFeld, "####0.00")
            cFeld = SwapStr(cFeld, ",", ".")
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            '6Kateg1
            If Not IsNull(rsrs!KATEGORIE) Then
                cFeld = rsrs!KATEGORIE
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            '7Kateg2
            If Not IsNull(rsrs!KATEGORIE2) Then
                cFeld = rsrs!KATEGORIE2
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            '8Kateg3
            cFeld = ""
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
        
            '9Beschreibung
            If Not IsNull(rsrs!BESCHREIB) Then
                cFeld = rsrs!BESCHREIB
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'bil
            cFeld = ".jpg"
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'leer
            cFeld = ""
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'br
            cFeld = "<br>"
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Name
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Artnr
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Vk
            If Not IsNull(rsrs!SHOPKVK) Then
                cFeld = rsrs!SHOPKVK
            Else
                cFeld = "0"
            End If
            
            If cFeld = "" Then cFeld = "0"
            cFeld = Format(cFeld, "####0.00")
            cFeld = SwapStr(cFeld, ",", ".")
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Bezeichnung2
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = rsrs!ARTBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Inhalt1
            cFeld = ""
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            'Beschreibung2
            If Not IsNull(rsrs!INTERBEZ) Then
                cFeld = rsrs!INTERBEZ
            Else
                cFeld = ""
            End If
            cFeld = "" & Chr(34) & "" & cFeld & "" & Chr(34) & "": cSatz = cSatz & cFeld & ";"
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            rsrs.MoveNext
        Loop
    End If
        
    Close iFileNr
        
    MsgBox "Die Exportdateien finden Sie unter: " & cPfad, vbInformation, "Winkiss Hinweis:"
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSV_Export_xtCommerceShop"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub CSV_Export_xtCommerceShop_v3(lLinr As Long, ctZ As String)
On Error GoTo LOKAL_ERROR

    Dim cFeld               As String
    Dim cSatz               As String
    Dim iFileNr             As Integer
    Dim rsrs                As Recordset
    Dim lPos                As Long
    Dim cPfad               As String
    Dim cPfad1              As String
    Dim cdatei              As String
    Dim sAusgabedatname     As String
    Dim sQuelle             As String
    Dim sZiel               As String
    Dim sQuellpfad          As String
    Dim sZielpfad           As String
    Dim lfail               As Long
    Dim lRet                As Long
    
    sQuellpfad = gcDBPfad
    sQuellpfad = ShortPath(sQuellpfad)
    If Right(sQuellpfad, 1) <> "\" Then
        sQuellpfad = sQuellpfad & "\"
    End If
    sQuellpfad = sQuellpfad & "PICTURE\ARTIKEL"
    
    Screen.MousePointer = 11
    
    sAusgabedatname = "Produkte.csv"
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"
    
    VerzVorhanden "Bilder", cPfad & "\"
    
    sZielpfad = gcDBPfad
    sZielpfad = ShortPath(sZielpfad)
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "BOX\Bilder"
    
    Kill sZielpfad & "\*.*"
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cFeld = "XTSOL" & ctZ & "p_model" & ctZ & "p_stock" & ctZ & "p_sorting" & ctZ & "p_shipping" & ctZ & "p_tpl" & ctZ & "p_manufacturer" & ctZ & "p_fsk18" & ctZ & "p_priceNoTax" & ctZ & "p_priceNoTax1" & ctZ & "p_priceNoTax2"
    cFeld = cFeld & "" & ctZ & "p_priceNoTax3" & ctZ & "p_tax" & ctZ & "p_status" & ctZ & "p_weight" & ctZ & "p_ean" & ctZ & "p_disc" & ctZ & "p_opttpl" & ctZ & "p_vpe" & ctZ & "p_vpe_status" & ctZ & "p_vpe_value" & ctZ & "p_image" & ctZ & "p_name.de" & ctZ & "p_desc.de"
    cFeld = cFeld & "" & ctZ & "p_shortdesc.de" & ctZ & "p_meta_title.de" & ctZ & "p_meta_desc.de" & ctZ & "p_meta_key.de" & ctZ & "p_keywords.de" & ctZ & "p_url.de" & ctZ & "p_cat.0" & ctZ & "p_cat.1" & ctZ & "p_cat.2"
    cFeld = cFeld & "" & ctZ & "p_cat.3" & ctZ & "p_cat.4" & ctZ & "p_cat.5"
    cSatz = cFeld
    cSatz = cSatz & Chr$(13) & Chr$(10)
        
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "SHOPEX", gdBase
    CreateTableT2 "SHOPEX", gdBase
    
    cSQL = "Insert into SHOPEX Select "
    cSQL = cSQL & " i.ARTNR  "
    cSQL = cSQL & ", i.ARTBEZ  "
    cSQL = cSQL & ", i.INTERBEZ  "
    cSQL = cSQL & ", i.BESCHREIB  "
    cSQL = cSQL & ", i.BILDkl  "
    cSQL = cSQL & ", i.BILDmi  "
    cSQL = cSQL & ", i.BILDgr  "
    cSQL = cSQL & ", i.SHOPKVK  "
    cSQL = cSQL & ", i.KATEGORIE  "
    cSQL = cSQL & ", i.KATEGORIE2  "
    cSQL = cSQL & ", " & lLinr & " as LINR "
    cSQL = cSQL & ", 0 as Bestand "
    cSQL = cSQL & ", '' as EAN "
    cSQL = cSQL & " from Interart i inner join SHOPTEMP a on i.Artnr = a.Artnr "
    gdBase.Execute cSQL, dbFailOnError
    
    If Check3.Value = vbChecked Then 'Preis, regulärer
        cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
        cSQL = cSQL & "  Set  SHOPEX.SHOPKVK = Artikel.KVKPR1 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Check4.Value = vbChecked Then 'für alle Standardwarnhinweis
        cSQL = "Update SHOPEX set Beschreib = Beschreib  + ' " & Text3(6).Text & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Update SHOPEX inner join Artikel on SHOPEX.Artnr = Artikel.Artnr "
    cSQL = cSQL & " Set SHOPEX.Bestand = Artikel.Bestand "
    cSQL = cSQL & " , SHOPEX.EAN = Artikel.EAN "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update SHOPEX inner join LISRT on SHOPEX.LINR = LISRT.LINR "
    cSQL = cSQL & "  Set  SHOPEX.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("select * from SHOPEX ")
    
    
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            'XTSOL
            cFeld = "XTSOL"
            cSatz = cFeld
            cSatz = cSatz & ctZ
            
'''            'neu action
'''
'''            cSatz = cSatz & "insert;"
        
            'p_model
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = "0"
            End If
            If cFeld = "" Then cFeld = "0"
            cSatz = cSatz & cFeld & ctZ
            
            'p_stock - Bestand
            
            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = "0"
            End If
            If cFeld = "" Then cFeld = "0"
            cSatz = cSatz & cFeld & ctZ
            
            'p_sorting
            
            cFeld = "0"
            cSatz = cSatz & cFeld & ctZ
            
            'p_shipping
            
            cFeld = "1"
            cSatz = cSatz & cFeld & ctZ
            
            'p_tpl
            
            cFeld = "product_info_v2.html"
            cSatz = cSatz & cFeld & ctZ
            
            'p_manufacturer
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = SwapStr(rsrs!LIEFBEZ, "§", ",")
            Else
                cFeld = "Diverse"
            End If
            If cFeld = "" Then cFeld = "Diverse"
            cSatz = cSatz & cFeld & ctZ
            
            'p_fsk18
            cFeld = "0"
            cSatz = cSatz & cFeld & ctZ
            
            'p_priceNoTax
            
            If Not IsNull(rsrs!SHOPKVK) Then
                cFeld = CDbl(rsrs!SHOPKVK) * 100 / 119
            Else
                cFeld = "0"
            End If
            If cFeld = "" Then cFeld = "0"
            cFeld = Format(cFeld, "####0.00")
            cFeld = SwapStr(cFeld, ",", ".")
            cSatz = cSatz & cFeld & ctZ
            
            'p_priceNoTax.1
            
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            
            'p_priceNoTax.2
            
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            
            'p_priceNoTax.3
            
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            
            'p_tax
            
            cFeld = "1"
            cSatz = cSatz & cFeld & ctZ
            
            'p_status
            
            cFeld = "1"
            cSatz = cSatz & cFeld & ctZ
            
            'p_weight
            
            cFeld = "0.00"
            cSatz = cSatz & cFeld & ctZ
            
            'p_ean
            
            If Not IsNull(rsrs!EAN) Then
                cFeld = SwapStr(rsrs!EAN, "§", ",")
            Else
                cFeld = ""
            End If
           
            cSatz = cSatz & cFeld & ctZ
            
            'p_disc
            cFeld = "0.00"
            cSatz = cSatz & cFeld & ctZ
            
            'p_opttpl
            cFeld = "default"
            cSatz = cSatz & cFeld & ctZ
            
            'p_vpe
            cFeld = "0"
            cSatz = cSatz & cFeld & ctZ
            
            'p_vpe_status
            cFeld = "0"
            cSatz = cSatz & cFeld & ctZ
            'p_vpe_value
            cFeld = "0.0000"
            cSatz = cSatz & cFeld & ctZ
            
            
            
            
            
            
            '*************++Bilder
            If Not IsNull(rsrs!Bildgr) Then
                sQuelle = sQuellpfad & "\" & rsrs!Bildgr
                sZiel = sZielpfad & "\" & rsrs!Bildgr

                lRet = CopyFile(sQuelle, sZiel, lfail)
            End If

            'p_image
            If Not IsNull(rsrs!Bildgr) Then
                cFeld = rsrs!Bildgr
            Else
                cFeld = "keinbild.jpg"
            End If
            If cFeld = "" Then cFeld = "keinbild.jpg"
            cSatz = cSatz & cFeld & ctZ
            '*************Bilder Ende
            
            'p_name.de
            If Not IsNull(rsrs!ARTBEZ) Then
                cFeld = SwapStr(rsrs!ARTBEZ, "§", ",")
            Else
                cFeld = "keine Angabe"
            End If
            If cFeld = "" Then cFeld = "keine Angabe"
            cSatz = cSatz & cFeld & ctZ
            
            'p_desc.de
            If Not IsNull(rsrs!BESCHREIB) Then
                cFeld = SwapStr(rsrs!BESCHREIB, "§", ",")
            Else
                cFeld = "keine Angabe"
            End If
            If cFeld = "" Then cFeld = "keine Angabe"
            cSatz = cSatz & cFeld & ctZ
            
            
    
            'p_shortdesc.de
            
            If Not IsNull(rsrs!INTERBEZ) Then
                cFeld = SwapStr(rsrs!INTERBEZ, "§", ",")
            Else
                cFeld = "keine Angabe"
            End If
            If cFeld = "" Then cFeld = "keine Angabe"
            cSatz = cSatz & cFeld & ctZ
            
            'p_meta_title.de
            
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            
            'p_meta_desc.de
            
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            
            'p_meta_key.de
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            'p_keywords.de
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            'p_url.de
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            'p_cat.0
            
            
            If Not IsNull(rsrs!KATEGORIE) Then
                cFeld = rsrs!KATEGORIE
            Else
                cFeld = "Diverse"
            End If
            If cFeld = "" Then cFeld = "Diverse"
            cSatz = cSatz & cFeld & ctZ
            
            
            'p_cat.1
            
            If Not IsNull(rsrs!KATEGORIE2) Then
                cFeld = rsrs!KATEGORIE2
            Else
                cFeld = "Diverse"
            End If
            If cFeld = "" Then cFeld = "Diverse"
            
            cSatz = cSatz & cFeld & ctZ
            
            'p_cat.2
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            'p_cat.3
            cFeld = ""
            cSatz = cSatz & cFeld & ctZ
            'p_cat.4
            cFeld = ""
            cSatz = cSatz & cFeld & ""
            'p_cat.5
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
                    
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
            Loop
        End If
        
        Close iFileNr
    
        rsrs.Close
    
    MsgBox "Die Exportdateien finden Sie unter: " & cPfad, vbInformation, "Winkiss Hinweis:"
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSV_Export_xtCommerceShop_v3"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub print_Shop_Art()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String

    loeschNEW "PRINTSHOP", gdBase
    CreateTableT2 "PRINTSHOP", gdBase
    
    cSQL = "Insert into PRINTSHOP Select "
    cSQL = cSQL & " ARTNR  "
    cSQL = cSQL & ", ARTBEZ  "
    cSQL = cSQL & ", INTERBEZ  "
    cSQL = cSQL & ", BESCHREIB  "
    cSQL = cSQL & ", BILDkl  "
    cSQL = cSQL & ", BILDmi  "
    cSQL = cSQL & ", BILDgr  "
    cSQL = cSQL & ", SHOPKVK  "
    cSQL = cSQL & ", KATEGORIE  "
    cSQL = cSQL & " from Interart "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update PRINTSHOP inner join ARTIKEL on PRINTSHOP.ARTNR = ARTIKEL.ARTNR "
    cSQL = cSQL & "  Set  PRINTSHOP.LINR = ARTIKEL.LINR "
    cSQL = cSQL & "  ,  PRINTSHOP.KVKPR1 = ARTIKEL.KVKPR1 "
    cSQL = cSQL & "  ,  PRINTSHOP.BESTAND = ARTIKEL.BESTAND "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update PRINTSHOP inner join LISRT on PRINTSHOP.LINR = LISRT.LINR "
    cSQL = cSQL & "  Set  PRINTSHOP.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update PRINTSHOP "
    cSQL = cSQL & "  Set  PRINTSHOP.BILDgr = 'kein Bild' "
    cSQL = cSQL & " where PRINTSHOP.BILDgr = ''"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update PRINTSHOP "
    cSQL = cSQL & "  Set  PRINTSHOP.BILDgr = 'kein Bild' "
    cSQL = cSQL & " where PRINTSHOP.BILDgr is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update PRINTSHOP "
    cSQL = cSQL & "  Set  PRINTSHOP.KATEGORIE = 'keine Kategorie' "
    cSQL = cSQL & " where PRINTSHOP.KATEGORIE = ''"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update PRINTSHOP "
    cSQL = cSQL & "  Set  PRINTSHOP.KATEGORIE = 'keine Kategorie' "
    cSQL = cSQL & " where PRINTSHOP.KATEGORIE is null"
    gdBase.Execute cSQL, dbFailOnError
    
    If Option1(0).Value = True Then
    'nach Lieferanten
    
        reportbildschirm "", "aWKL163a"
    
    ElseIf Option1(1).Value = True Then
    'nach Kategorien
        reportbildschirm "", "aWKL163b"
    End If
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "print_Shop_Art"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Kat_Del(cbox As ComboBox, ckatSpalte As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    Dim iRet2 As Integer
    Dim cKat As String
    Dim lAnzDelArtproKat As Long
    
    
    
    If cbox.Text <> "" Then
        If cbox.Text <> "bitte wählen" Then
            cKat = Trim(cbox.Text)
            iRet = (MsgBox("Möchten Sie die Kategorie: " & cKat & " wirklich löschen?", vbQuestion + vbYesNo, "Winkiss Frage:"))
            If iRet = vbYes Then
                
                lAnzDelArtproKat = ermAnzahlKat(cKat, ckatSpalte)
                
                If lAnzDelArtproKat = 1 Then 'Artikel Artikeln
                    iRet2 = (MsgBox("Jetzt wird die Kategorie: " & cKat & " bei  " & lAnzDelArtproKat & " Artikel entfernt. Wirklich löschen?", vbQuestion + vbYesNo, "Winkiss Frage:"))
                Else
                    iRet2 = (MsgBox("Jetzt wird die Kategorie: " & cKat & " bei  " & lAnzDelArtproKat & " Artikeln entfernt. Wirklich löschen?", vbQuestion + vbYesNo, "Winkiss Frage:"))
                End If
                
                If iRet = vbYes Then
                
                    sSQL = "Update INTERART set " & ckatSpalte & " = '' where " & ckatSpalte & " = '" & cKat & "'"
                    gdBase.Execute sSQL, dbFailOnError
                        
                End If
                
            End If
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Kat_Del"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Art_Del()
On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    If Label5(1).Caption <> "" Then
        iRet = (MsgBox("Möchten Sie wirklich den Artikel: '" & Label5(7).Caption & "' als Shop-Artikel entfernen?", vbQuestion + vbYesNo, "Winkiss Frage:"))
        If iRet = vbYes Then
            
            delInterart Trim(Label5(1).Caption)
            
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Art_Del"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermAnzahlKat(cKat As String, ckatSpalte As String) As Long
On Error GoTo LOKAL_ERROR

    ermAnzahlKat = 0

    cSQL = "Select count(*) as maxi from INTERART "
    cSQL = cSQL & " Where '" & ckatSpalte & "' = '" & cKat & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermAnzahlKat = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermAnzahlKat"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Function ermAnzahlShopArt() As Long
On Error GoTo LOKAL_ERROR

    ermAnzahlShopArt = 0

    cSQL = "Select count(*) as maxi from INTERART "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermAnzahlShopArt = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermAnzahlShopArt"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub Pic_hinzu(sArtnr As String)
On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sBilddatei  As String
    
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    
    Dim sPfad As String
    
    cZiel = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If

    sPfad = sPfad & "PICTURE\ARTIKEL"

    sTitle = "Speichern der Bild/Logo - Datei"
    
'    cdlBild.Filter = "Alle Bilddateien |*.*|Bitmap-Dateien (*.bmp)|*.bmp|GIF (Graphics Interchange Format)" & _
'    "(*.gif)|*.gif|JPEG (File Interchange Format) (*.jpg;*.jpeg)|*.jpg;*.jpeg" & _
'    "|Alle Dateien|*.*|"
    
    sFilter = "JPEG (*.JPG)| *.JPG|GIF (*.GIF)| *.GIF|PNG (*.PNG)| *.PNG| Bitmapdateien (*.bmp)|*.bmp"
'    sFilter = "JPEG (*.JPG)| *.JPG| Bitmapdateien (*.bmp)|*.bmp"
'    sOldpfad = "C:"
    
    sBilddatei = pfadaendernplusDatname(sTitle, sFilter, sOldpfad)
    
    cQuelle = sBilddatei
    cQuelle = ShortPath(cQuelle)

    cZiel = gcDBPfad
    If Right(cZiel, 1) <> "\" Then
        cZiel = cZiel & "\"
    End If
    cZiel = ShortPath(cZiel)
    
    cZiel = cZiel & "PICTURE\ARTIKEL"
    
    'multibild
    File1.Path = cZiel
    File1.Pattern = sArtnr & "*.jpg"
    File1.Refresh
                
    If File1.ListCount > 0 Then
        cZiel = cZiel & "\" & sArtnr & "_" & File1.ListCount & ".jpg"
    Else
        cZiel = cZiel & "\" & sArtnr & ".jpg"
    End If
    
    
    
    
    
    
    

    lRet = CopyFile(cQuelle, cZiel, lfail)
    
    ZeigeBilder sArtnr, 1

    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Pic_hinzu"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ZeigeBilder(sArt As String, iStatus As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim sSpeicherpfad   As String
    Dim i               As Integer
    Dim sBildname       As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    For i = 0 To 7
        Picture2(i).Enabled = False
        Picture2(i).Visible = False
    Next i
    
    If FileExists(sPfad & "\" & sArt & ".jpg") Then
        Image1.Picture = LoadPicture(sPfad & "\" & sArt & ".jpg")
        Image2.Picture = LoadPicture(sPfad & "\" & sArt & ".jpg")
        sSpeicherpfad = sPfad & "\" & sArt & "_s.jpg"
        Label5(12).Caption = sArt & ".jpg"
        
        Picture1.Visible = True
        
        File1.Path = sPfad
        File1.Pattern = sArt & "*.jpg"
        File1.Refresh
                    
        If File1.ListCount > 1 Then
            'dann zeige weitere
            For i = 0 To File1.ListCount - 1
        
                If i > 7 Then
                    Exit For
                End If
                sBildname = File1.list(i)
                Picture2(i).Tag = sBildname
                Zeige_weitere_Bilder sBildname, 0, Picture2(i), Image4, 50
                
                If i > 0 Then
                    Picture2(i).Top = Picture2(0).Top
                    Picture2(i).Left = Picture2(i - 1).Left + 50 + Picture2(i - 1).Width
                End If
                
            Next i
            
            
        End If
        
        
        
        
    Else
        sSpeicherpfad = ""
        If FileExists(sPfad & "\" & "keinBild.jpg") Then
            Image1.Picture = LoadPicture(sPfad & "\" & "keinBild.jpg")
            Image2.Picture = LoadPicture(sPfad & "\" & "keinBild.jpg")
            
            Label5(12).Caption = "keinBild.jpg"
            
            Picture1.Visible = True
'            Picture2.Visible = True
        Else
            Picture1.Visible = False
'            Picture2.Visible = False
            
            Label5(12).Caption = ""
            
            Exit Sub
        End If
    End If
    
    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 350
        Case Is > 12000
            iDiv = 300
        Case Is > 11000
            iDiv = 250
        Case Is > 10000
            iDiv = 220
        Case Is > 8000
            iDiv = 200
    End Select
    
    Picture1.Tag = sArt & ".jpg"
    zeigImage_In_Picture Image1, Picture1, iDiv, "" 'sSpeicherpfad
'    zeigImage_In_Picture Image2, Picture2, 50, "" 'sSpeicherpfad
    
'
    
    
    
    Picture1.Refresh
'    Picture2.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 481 Then
        If iStatus = 1 Then
            MsgBox "Dieses Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        End If
        Kill sPfad & "\" & sArt & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ZeigeBilder"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
'    Resume Next
End Sub
Private Sub DelBild(sArt As String)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim i               As Integer
    Dim sBildname       As String
    
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim lRet            As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    
    
    Kill sPfad & "\" & sArt
    
    Dim sArtnr As String
    sArtnr = Left(sArt, Len(sArt) - 4)
    
    File1.Path = sPfad
    File1.Pattern = sArtnr & "*.jpg"
    File1.Refresh
                
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sBildname = File1.list(i)
            
            cQuelle = sPfad & "\"
            cQuelle = ShortPath(cQuelle)
            cQuelle = cQuelle & sBildname
    
            cZiel = sPfad & "\"
            cZiel = ShortPath(cZiel)
            If i > 0 Then
                cZiel = cZiel & "del_" & i & ".jpg"
            Else
                cZiel = cZiel & "del.jpg"
            End If
    
            lRet = CopyFile(cQuelle, cZiel, lfail)
        Next i
    End If
    
    Kill sPfad & "\" & sArtnr & "*.jpg"
    
    File1.Path = sPfad
    File1.Pattern = "del*.jpg"
    File1.Refresh
                
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sBildname = File1.list(i)
            
            cQuelle = sPfad & "\"
            cQuelle = ShortPath(cQuelle)
            cQuelle = cQuelle & sBildname
    
            cZiel = sPfad & "\"
            cZiel = ShortPath(cZiel)
            If i > 0 Then
                cZiel = cZiel & sArtnr & "_" & i & ".jpg"
            Else
                cZiel = cZiel & sArtnr & ".jpg"
            End If
    
            lRet = CopyFile(cQuelle, cZiel, lfail)
        Next i
    End If
    
    Kill sPfad & "\del*.jpg"
        
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DelBild"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Zeige_weitere_Bilder(sBild As String, iStatus As Integer, PicX As PictureBox, imx As Image, iDiv As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    
    If sBild = "" Then
        Exit Sub
    End If
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    If FileExists(sPfad & "\" & sBild) Then
        imx.Picture = LoadPicture(sPfad & "\" & sBild)
        
        PicX.Visible = True
        PicX.Enabled = True
    End If
    
    
    zeigImage_In_Picture imx, PicX, iDiv, ""
    PicX.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 481 Then
        If iStatus = 1 Then
            MsgBox "Dieses Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        End If
        Kill sPfad & "\" & sArt & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Zeige_weitere_Bilder"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub zeigImage_In_Picture(imgx As Image, PicX As PictureBox, iDiv As Integer, sSpeicherpfad As String)
    On Error GoTo LOKAL_ERROR

    Dim höhe As Integer
    Dim Breite As Integer
    Dim iTeiler As Integer

    If imgx.Width >= imgx.Height Then
        iTeiler = imgx.Width / iDiv
    Else
        iTeiler = imgx.Height / iDiv
    End If
    
    höhe = imgx.Height / iTeiler
    Breite = imgx.Width / iTeiler
    
    imgx.Height = höhe * Screen.TwipsPerPixelX
    imgx.Width = Breite * Screen.TwipsPerPixelY
    
    PicX.Picture = LoadPicture("")
    
    With PicX
        .BorderStyle = 0
        .Width = imgx.Width
        .Height = imgx.Height
        
        ' Wichtig: AutoRedraw = True
        .AutoRedraw = True
        
        ' Bild aus ImageBox übertragen
        .PaintPicture imgx.Picture, 0, 0, _
        imgx.Width, imgx.Height
      
        ' Bild abspeichern
        If sSpeicherpfad <> "" Then
            SavePicture .Image, sSpeicherpfad 'sPfad & "\" & gsARTNR & "kl.jpg"
        End If
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigImage_In_Picture"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeExportBild(sArt As String, iStatus As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim i               As Integer
    Dim sBildname       As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    If FileExists(sPfad & "\" & sArt & ".jpg") Then
        Image3.Picture = LoadPicture(sPfad & "\" & sArt & ".jpg")
        
        
        File1.Path = sPfad
        File1.Pattern = sArt & "*.jpg"
        File1.Refresh
                    
        If File1.ListCount > 1 Then
            'dann zeige weitere
            For i = 0 To File1.ListCount - 1
        
                If i > 7 Then
                    Exit For
                End If
                sBildname = File1.list(i)
                Picture4(i).Tag = sBildname
                Zeige_weitere_Bilder sBildname, 0, Picture4(i), Image4, 50
                
                If i > 0 Then
                    Picture4(i).Top = Picture4(0).Top
                    Picture4(i).Left = Picture4(i - 1).Left + 50 + Picture4(i - 1).Width
                End If
                
            Next i
            
            
        End If
        
        
        
        
    Else
        If FileExists(sPfad & "\" & "keinBild.jpg") Then
            Image3.Picture = LoadPicture(sPfad & "\" & "keinBild.jpg")
        Else
            Picture3.Visible = False
            Exit Sub
        End If
    End If
    
    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 350
        Case Is > 12000
            iDiv = 300
        Case Is > 11000
            iDiv = 250
        Case Is > 10000
            iDiv = 220
        Case Is > 8000
            iDiv = 200
    End Select
    
    zeigImage_In_Picture Image3, Picture3, iDiv, ""

Exit Sub
LOKAL_ERROR:
    If err.Number = 481 Then
''        MsgBox sArt & ".jpg wird gelöscht"
        If iStatus = 1 Then
            MsgBox "Dieses Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        End If
        Kill sPfad & "\" & sArt & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ZeigeExportBild"
        Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
'    Resume Next
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
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub List1_Click()
On Error GoTo LOKAL_ERROR

    If List1.ListIndex < 0 Then
    
    Else
    
        For i = 0 To 7
            Picture4(i).Visible = False
            Picture4(i).Tag = ""
            Picture4(i).Picture = LoadPicture("")
            
        Next i
    
        ZeigeExportBild Left(List1.list(List1.ListIndex), 6), 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Picture2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 350
        Case Is > 12000
            iDiv = 300
        Case Is > 11000
            iDiv = 250
        Case Is > 10000
            iDiv = 220
        Case Is > 8000
            iDiv = 200
    End Select
    
    Dim ctemp As String
    ctemp = Picture2(Index).Tag
    
    Zeige_weitere_Bilder Picture2(Index).Tag, 0, Picture1, Image4, iDiv
    
    Label5(12).Caption = ctemp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture2_Click"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 350
        Case Is > 12000
            iDiv = 300
        Case Is > 11000
            iDiv = 250
        Case Is > 10000
            iDiv = 220
        Case Is > 8000
            iDiv = 200
    End Select
    
    Dim ctemp As String
    ctemp = Picture2(Index).Tag
    
    Zeige_weitere_Bilder Picture2(Index).Tag, 0, Picture1, Image4, iDiv
    Label5(12).Caption = ctemp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture2_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Picture4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 350
        Case Is > 12000
            iDiv = 300
        Case Is > 11000
            iDiv = 250
        Case Is > 10000
            iDiv = 220
        Case Is > 8000
            iDiv = 200
    End Select
    
    Zeige_weitere_Bilder Picture4(Index).Tag, 0, Picture3, Image4, iDiv
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture4_Click"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Picture4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 350
        Case Is > 12000
            iDiv = 300
        Case Is > 11000
            iDiv = 250
        Case Is > 10000
            iDiv = 220
        Case Is > 8000
            iDiv = 200
    End Select
    
    Zeige_weitere_Bilder Picture4(Index).Tag, 0, Picture3, Image4, iDiv
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture4_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_Change(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 2
            anzeige "normal", "noch " & 100 - Len(Text3(2).Text) & " Zeichen", Label5(3)
        Case 0
            anzeige "normal", "noch " & 255 - Len(Text3(0).Text) & " Zeichen", Label5(4)
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_Change"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = glSelBack1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    Select Case Index
        Case Is = 0, 1, 2
    
'            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
'            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
'            cValid = cValid & "+äÄÜüÖöß"
'            cValid = cValid & vbKeyControl
'
'            cZeichen = Chr$(KeyAscii)
'
'            If InStr(cValid, cZeichen) = 0 Then
'                KeyAscii = 0
'            End If
        Case 3
            cValid = gcNUM & "," & Chr$(8)
        
            cZeichen = Chr$(KeyAscii)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 4
            cValid = gcNUM & Chr$(8)
        
            cZeichen = Chr$(KeyAscii)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        Command4_Click 2
    End If
    
    If Index = 4 Then
        If KeyCode = vbKeyF2 Then
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            
            If gF2Prompt.cWahl <> "" Then
                Text3(4).Text = gF2Prompt.cWahl
                
            End If
            Text3(4).SetFocus
        
        End If
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

