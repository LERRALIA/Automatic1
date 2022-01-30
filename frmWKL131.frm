VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWKL131 
   Caption         =   "Einstellungen an der Kasse"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL131.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      Caption         =   "Erläuterung"
      Height          =   735
      Left            =   8400
      TabIndex        =   226
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   18
         Left            =   9480
         TabIndex        =   228
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lblErlaeuterung 
         Caption         =   "Erläuterung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   240
         TabIndex        =   227
         Top             =   360
         Width           =   11175
      End
   End
   Begin VB.Frame fraAuszahlungsgrund 
      Caption         =   "Auszahlungsgründe"
      Height          =   255
      Left            =   7680
      TabIndex        =   146
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   9
         Left            =   600
         MaxLength       =   50
         TabIndex        =   159
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   8
         Left            =   600
         MaxLength       =   50
         TabIndex        =   158
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   7
         Left            =   600
         MaxLength       =   50
         TabIndex        =   157
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   6
         Left            =   600
         MaxLength       =   50
         TabIndex        =   156
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   5
         Left            =   600
         MaxLength       =   50
         TabIndex        =   155
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   4
         Left            =   600
         MaxLength       =   50
         TabIndex        =   154
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   3
         Left            =   600
         MaxLength       =   50
         TabIndex        =   153
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   2
         Left            =   600
         MaxLength       =   50
         TabIndex        =   152
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   1
         Left            =   600
         MaxLength       =   50
         TabIndex        =   151
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   50
         TabIndex        =   150
         Top             =   840
         Width           =   2415
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   10
         Left            =   9480
         TabIndex        =   147
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "Auszahlungsgründe:"
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
         Index           =   24
         Left            =   120
         TabIndex        =   148
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame fraKassenabschluss 
      Caption         =   "Kassenabschluss / Bargeldeingabe"
      Height          =   255
      Left            =   7680
      TabIndex        =   140
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "vollständige Bargeldeingabe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   145
         Top             =   840
         Value           =   -1  'True
         Width           =   5535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bargeldeingabe mit Münzen und Scheinen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   144
         Top             =   1320
         Width           =   5535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "minimierte Bargeldeingabe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   143
         Top             =   1800
         Width           =   5415
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   9
         Left            =   9480
         TabIndex        =   141
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "Darstellungsart der Bargeldeingabe"
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
         Index           =   27
         Left            =   120
         TabIndex        =   142
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame fraSpez 
      Caption         =   "spezielle Kasseneinstellungen"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   39
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   264
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   38
         Left            =   120
         MaxLength       =   6
         TabIndex        =   263
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   37
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   259
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   36
         Left            =   120
         MaxLength       =   6
         TabIndex        =   258
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   35
         Left            =   120
         MaxLength       =   6
         TabIndex        =   255
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   34
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   254
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   33
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   251
         Top             =   4080
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   32
         Left            =   120
         MaxLength       =   6
         TabIndex        =   250
         Top             =   4080
         Width           =   855
      End
      Begin VB.CheckBox Check47 
         Caption         =   "immer minimieren (leerer Warenkorb)"
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
         Left            =   4680
         TabIndex        =   246
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   8520
         MaxLength       =   1
         TabIndex        =   245
         ToolTipText     =   "Tage vor dem Termin"
         Top             =   6840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check44 
         Caption         =   "Termin Erinnerung per SMS"
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
         Left            =   4440
         TabIndex        =   240
         Top             =   6840
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   237
         Top             =   6840
         Width           =   975
      End
      Begin VB.CheckBox Check40 
         Caption         =   "auch bei Terminpreisen anwenden"
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
         Left            =   360
         TabIndex        =   221
         Top             =   4950
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   10200
         MaxLength       =   6
         TabIndex        =   205
         Top             =   5760
         Width           =   855
      End
      Begin VB.CheckBox Check37 
         Caption         =   "Staffelpreise anwenden"
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
         Left            =   8040
         TabIndex        =   203
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   201
         ToolTipText     =   "Diese Zielfiliale von der Automatik ausschliessen"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   8040
         MaxLength       =   6
         TabIndex        =   195
         ToolTipText     =   "Tragen Sie hier die Artikelnummer ein. (die ersten 6 Stellen des EAN, mit 222 am Anfang)"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   6240
         MaxLength       =   6
         TabIndex        =   193
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check36 
         Caption         =   "Parkvorgänge netto"
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
         Left            =   8040
         TabIndex        =   192
         Top             =   6360
         Width           =   2415
      End
      Begin VB.CheckBox Check33 
         Caption         =   "Nachfragen bei Warengruppen ohne Preis"
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
         Left            =   8040
         TabIndex        =   185
         Top             =   6000
         Width           =   3495
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
         Height          =   1770
         Index           =   1
         Left            =   8040
         TabIndex        =   180
         Top             =   2640
         Width           =   3015
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   24
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   199
            Top             =   1440
            Width           =   975
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
            Index           =   23
            Left            =   1920
            MaxLength       =   9
            TabIndex        =   197
            Text            =   "Text1"
            ToolTipText     =   "Bei Erreichen der Bonusgrenze Gutschein anbieten."
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox Check41 
            Caption         =   "Bonus beim nächsten Besuch"
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
            Left            =   120
            TabIndex        =   183
            Top             =   800
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
            Index           =   20
            Left            =   1800
            MaxLength       =   9
            TabIndex        =   182
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check38 
            Caption         =   "Bonusauszahlung Stammfiliale"
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
            TabIndex        =   181
            Top             =   560
            Width           =   2655
         End
         Begin VB.Label lbl6 
            Caption         =   "Bonusabzug (Artnr)"
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
            TabIndex        =   200
            ToolTipText     =   "Bei Erreichen der Bonusgrenze Gutschein anbieten."
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label lbl6 
            Caption         =   "Gutscheinwert:"
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
            Index           =   36
            Left            =   120
            TabIndex        =   198
            ToolTipText     =   "Bei Erreichen der Bonusgrenze Gutschein anbieten."
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lbl6 
            Caption         =   "Bonusgrenze:"
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
            Index           =   33
            Left            =   120
            TabIndex        =   184
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   178
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox Check28 
         Caption         =   "Sonderpreis rabattierfähig"
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
         Left            =   8040
         TabIndex        =   122
         Top             =   720
         Width           =   2895
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
         Height          =   855
         Index           =   5
         Left            =   8040
         TabIndex        =   119
         Top             =   4320
         Width           =   3015
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
            Index           =   27
            Left            =   1800
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   208
            Top             =   480
            Width           =   1095
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
            Index           =   16
            Left            =   1800
            MaxLength       =   9
            PasswordChar    =   "*"
            TabIndex        =   120
            Top             =   140
            Width           =   1095
         End
         Begin VB.Label lbl6 
            Caption         =   "Startpasswort:"
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
            Index           =   40
            Left            =   120
            TabIndex        =   207
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lbl6 
            Caption         =   "Kassenpasswort:"
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
            TabIndex        =   121
            Top             =   195
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   10200
         MaxLength       =   6
         TabIndex        =   117
         Top             =   5400
         Width           =   855
      End
      Begin VB.CheckBox Check26 
         Caption         =   "einfache Gutscheinerstellung statt Geld auszahlen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8040
         TabIndex        =   114
         Top             =   960
         Width           =   2895
      End
      Begin VB.CheckBox Check25 
         Caption         =   "mit Couponüberprüfung"
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
         Left            =   8040
         TabIndex        =   113
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   4440
         MaxLength       =   150
         TabIndex        =   111
         Top             =   6480
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   6240
         MaxLength       =   6
         TabIndex        =   109
         Top             =   5400
         Width           =   855
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Münzen und Scheine anzeigen"
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
         Left            =   4440
         TabIndex        =   96
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CheckBox Check109 
         Caption         =   "bei EC Lastschrift, nachträgliche Kundenbindung zulassen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   4440
         TabIndex        =   95
         Top             =   2880
         Width           =   3135
      End
      Begin VB.CheckBox Check51 
         Caption         =   "KB nur mit Kundenbindung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   94
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CheckBox Check66 
         Caption         =   "Bedienernummer leeren"
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
         Left            =   4440
         TabIndex        =   93
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox Check105 
         Caption         =   "nur Angemeldete an der Kasse zulassen"
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
         Left            =   4440
         TabIndex        =   92
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7200
         MaxLength       =   6
         TabIndex        =   91
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7200
         MaxLength       =   6
         TabIndex        =   90
         Top             =   4440
         Width           =   615
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Display mittels Zweitmonitor"
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
         Left            =   4440
         TabIndex        =   89
         Top             =   3360
         Width           =   3255
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Handelsspanne ausblenden"
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
         Left            =   4440
         TabIndex        =   88
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CheckBox Check22 
         Caption         =   "einfache Zollerstattung (Barzahlung)"
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
         Left            =   4440
         TabIndex        =   87
         Top             =   4800
         Width           =   3495
      End
      Begin VB.CheckBox Check24 
         Caption         =   "'alter Gutschein' ausblenden"
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
         Left            =   4440
         TabIndex        =   86
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   6240
         MaxLength       =   6
         TabIndex        =   85
         Top             =   960
         Width           =   855
      End
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
         Height          =   255
         Left            =   6240
         TabIndex        =   84
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   6240
         MaxLength       =   6
         TabIndex        =   83
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   6240
         MaxLength       =   6
         TabIndex        =   82
         Top             =   5040
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         Caption         =   "kein Kundenbonus erhöhen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   58
         Top             =   6120
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   6
         TabIndex        =   57
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   2160
         TabIndex        =   56
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   54
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   53
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   52
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   51
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   120
         MaxLength       =   20
         TabIndex        =   50
         ToolTipText     =   "die ersten 15 Zeichen werden auf dem Kassenbon gedruckt"
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Kundenrabatte deaktivieren"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   49
         Top             =   5880
         Width           =   3015
      End
      Begin VB.CheckBox Check18 
         Caption         =   "kein Kundenbonus erhöhen, wenn KVK um x% < LVK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   48
         Top             =   6360
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   195
         Index           =   5
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   47
         Top             =   6360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check19 
         Caption         =   "kein Kundenbonus erhöhen, bei Gewährung von Artikel- bzw. Gesamtrabatt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   46
         Top             =   6600
         Width           =   4335
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   6
         Left            =   9480
         TabIndex        =   14
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   55
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
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
            Size            =   12
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
         Caption         =   ">>"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPickerSMS 
         Height          =   255
         Left            =   7200
         TabIndex        =   241
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   120455170
         CurrentDate     =   43936.4166666667
      End
      Begin VB.Label lbl6 
         Caption         =   "ab 2 Artikel des Lieferanten = 10% Artikelrabatt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   53
         Left            =   120
         TabIndex        =   262
         Top             =   5160
         Width           =   4095
      End
      Begin VB.Label lbl6 
         Caption         =   "in %"
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
         Index           =   52
         Left            =   1680
         TabIndex        =   261
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lbl6 
         Caption         =   "ab x Euro Warenkorbwert = x Gesamtrabatt (nur auf rabattierfähige Artikel)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   170
         Index           =   51
         Left            =   120
         TabIndex        =   260
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   2
         X1              =   120
         X2              =   4080
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lbl6 
         Caption         =   "in %"
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
         Index           =   50
         Left            =   1680
         TabIndex        =   257
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label lbl6 
         Caption         =   "20% auf das Lieblingsprodukt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   170
         Index           =   49
         Left            =   120
         TabIndex        =   256
         Top             =   3120
         Width           =   4095
      End
      Begin VB.Label lbl6 
         Caption         =   "Beim Kassieren: Artikel mit Artikelrabatt anwenden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   48
         Left            =   120
         TabIndex        =   253
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Label lbl6 
         Caption         =   "in %"
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
         Left            =   1680
         TabIndex        =   252
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lbl6 
         Caption         =   "Bonusauszahlung über Kundenhistorie (Artnr)"
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
         TabIndex        =   238
         ToolTipText     =   "Bei Erreichen der Bonusgrenze Gutschein anbieten."
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         Caption         =   "EC Auszahlung:"
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
         Index           =   39
         Left            =   8760
         TabIndex        =   206
         ToolTipText     =   $"frmWKL131.frx":0442
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label lbl6 
         Caption         =   "AF"
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
         Index           =   38
         Left            =   9360
         TabIndex        =   202
         ToolTipText     =   "Diese Zielfiliale von der Automatik ausschliessen"
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lbl6 
         Caption         =   "Lottoauszahlung:"
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
         Index           =   35
         Left            =   8040
         TabIndex        =   196
         ToolTipText     =   "Tragen Sie hier die Artikelnummer ein. (die ersten 6 Stellen des EAN, mit 222 am Anfang)"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label lbl6 
         Caption         =   "Lieferant Zeitungen"
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
         Left            =   4440
         TabIndex        =   194
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   $"frmWKL131.frx":04ED
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   32
         Left            =   8040
         TabIndex        =   179
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         Caption         =   "spez Fotoartikel:"
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
         Left            =   9720
         TabIndex        =   118
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label lbl6 
         Caption         =   "Abverkaufsliste (spez L unter 2.Bon) Geben Sie die beteiligten ArtNr getrennt mit Komma ein!"
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
         Index           =   2
         Left            =   4440
         TabIndex        =   112
         Top             =   5760
         Width           =   3375
      End
      Begin VB.Label lbl6 
         Caption         =   "Paketlieferant:"
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
         Left            =   4440
         TabIndex        =   110
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   "verbindlicher Gesamtrabatt"
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
         Left            =   4440
         TabIndex        =   104
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label lbl6 
         Caption         =   "Warnung bei einem Preis kleiner"
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
         Left            =   4440
         TabIndex        =   103
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label lbl6 
         Caption         =   "Geschenkset Artnr"
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
         Left            =   4440
         TabIndex        =   102
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   "EUR"
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
         Index           =   60
         Left            =   7200
         TabIndex        =   101
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl6 
         Caption         =   "Schwellenwert bei Kartenzahlung"
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
         Index           =   61
         Left            =   4440
         TabIndex        =   100
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lbl6 
         Caption         =   "Spanne/Zeitungen"
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
         Left            =   4440
         TabIndex        =   99
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   "%"
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
         Left            =   7200
         TabIndex        =   98
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl6 
         Caption         =   "primärer Lieferant:"
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
         Left            =   4440
         TabIndex        =   97
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   "AboPlus-Karte auf folgende Artikel anwenden:"
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
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lbl6 
         Caption         =   "Partnerfirmennummer"
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
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lbl6 
         Caption         =   "Bei Erreichen der Bonusgrenze diesen Warengruppen - Artikel abziehen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label lbl6 
         Caption         =   "Bei Benutzung der  folgenden Warengruppe (Artnr) Erfassung der Bonusnummer einblenden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   62
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   4080
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   3
         X1              =   120
         X2              =   4080
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label lbl6 
         Caption         =   "Jubiläumsrabatt, Eröffnungsrabatt "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   8
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "die ersten 15 Zeichen werden auf dem Kassenbon gedruckt"
         Top             =   4440
         Width           =   3015
      End
      Begin VB.Label lbl6 
         Caption         =   "in %"
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
         Left            =   2520
         TabIndex        =   60
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label lbl6 
         Caption         =   "x in %"
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
         Index           =   11
         Left            =   3720
         TabIndex        =   59
         Top             =   6360
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame fraAllg 
      Caption         =   "allgemeine Kasseneinstellungen"
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox Check48 
         Caption         =   "keine Bestandsveränderungen bei Warengruppen"
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
         TabIndex        =   249
         Top             =   5720
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   6360
         MaxLength       =   100
         TabIndex        =   242
         Top             =   6120
         Width           =   2295
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
         Height          =   810
         Index           =   0
         Left            =   4440
         TabIndex        =   223
         Top             =   6360
         Width           =   4185
         Begin VB.CheckBox Check45 
            Caption         =   "ab sofort anwenden"
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
            TabIndex        =   224
            Top             =   480
            Width           =   1935
         End
         Begin sevCommand3.Command Command5 
            Height          =   255
            Index           =   17
            Left            =   3120
            TabIndex        =   229
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
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
            Caption         =   "Info"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   255
            Index           =   19
            Left            =   2640
            TabIndex        =   230
            Top             =   480
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
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
            Caption         =   "x"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label lbl6 
            Caption         =   "Gutscheine beim Verkauf versteuern"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   44
            Left            =   120
            TabIndex        =   225
            Top             =   120
            Width           =   3975
         End
      End
      Begin VB.CheckBox Check39 
         Caption         =   "Kassieren mit Hinweis auf Kundenwahl"
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
         TabIndex        =   204
         Top             =   840
         Width           =   3975
      End
      Begin VB.CheckBox Check35 
         Caption         =   "auch bei Kundenauswahl"
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
         Left            =   4680
         TabIndex        =   187
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox Check32 
         Caption         =   "bei nicht rabattierfähigen Artikeln im Warenkorb nicht runden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9240
         TabIndex        =   177
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   9240
         MaxLength       =   6
         TabIndex        =   174
         Top             =   6240
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Runden"
         Height          =   1095
         Left            =   9240
         TabIndex        =   169
         Top             =   4800
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "sinnvoll runden"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   173
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "aufrunden"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   171
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "abrunden"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   170
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   14
            Left            =   9480
            TabIndex        =   172
            Top             =   6720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
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
            Caption         =   "Weiter"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin VB.TextBox Text1 
         Height          =   885
         Index           =   17
         Left            =   9240
         Locked          =   -1  'True
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   162
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CheckBox Check31 
         Caption         =   "bei Kundenwahl, gesperrte = Zeile ""rot"" darstellen"
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
         Left            =   4440
         TabIndex        =   138
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   6360
         MaxLength       =   100
         TabIndex        =   136
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CheckBox Check30 
         Caption         =   "bei Kundenwahl (Kunden mit Farbmerkmal) immer Farbbeschreibung anzeigen "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4440
         TabIndex        =   135
         Top             =   5160
         Width           =   4335
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
         Height          =   1290
         Index           =   3
         Left            =   120
         TabIndex        =   123
         Top             =   5880
         Width           =   4185
         Begin VB.CheckBox Check43 
            Caption         =   "Nr Komplett"
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
            Left            =   2760
            TabIndex        =   222
            Top             =   960
            Width           =   1215
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
            Height          =   255
            Left            =   2760
            TabIndex        =   166
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton opt1 
            Caption         =   "von Hand"
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
            TabIndex        =   127
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "auto"
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
            Index           =   5
            Left            =   1440
            TabIndex        =   126
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox Check42 
            Caption         =   "ohne Vorschlag"
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
            TabIndex        =   125
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check46 
            Caption         =   "Restgutschein = Original"
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
            Left            =   120
            TabIndex        =   124
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lbl6 
            Caption         =   "Restgutscheine erzeugen ab:"
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
            Left            =   2760
            TabIndex        =   168
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbl6 
            Caption         =   "EUR"
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
            Left            =   3600
            TabIndex        =   167
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbl6 
            Caption         =   "GutscheinNr vergeben:"
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
            Left            =   120
            TabIndex        =   128
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.CheckBox Check109 
         Caption         =   "Kassieren nur mit Bestand"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   116
         Top             =   4920
         Width           =   2535
      End
      Begin VB.CheckBox Check27 
         Caption         =   "passendes Bargeld bestätigen"
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
         Left            =   4440
         TabIndex        =   115
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   885
         Index           =   11
         Left            =   4680
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   108
         Top             =   3720
         Width           =   3855
      End
      Begin VB.CheckBox Check107 
         Caption         =   "MB Blocken? Diese Frage stellen"
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
         Left            =   4440
         TabIndex        =   107
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   480
         MaxLength       =   100
         TabIndex        =   105
         Top             =   5400
         Width           =   3615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Kassenschublade öffnen bei Retoure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CheckBox Check63 
         Caption         =   "Ges Rabatt und Art Rabatt verschleiern"
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
         TabIndex        =   43
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CheckBox Check29 
         Caption         =   "Schublade bei Kreditkartenzahlungen nicht öffnen"
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
         TabIndex        =   42
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CheckBox Check71 
         Caption         =   "nur 'geführte' Artikel an der Kasse zulassen"
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
         Left            =   4440
         TabIndex        =   41
         Top             =   2880
         Width           =   3975
      End
      Begin VB.CheckBox Check70 
         Caption         =   "Bon Ja/Nein an der Kasse einblenden"
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
         TabIndex        =   40
         Top             =   2280
         Width           =   4095
      End
      Begin VB.CheckBox Check69 
         Caption         =   "Bon drucken aus"
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
         TabIndex        =   39
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox Check62 
         Caption         =   "Umsatzanzeige aktivieren (2. Bon)"
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
         TabIndex        =   38
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox Check60 
         Caption         =   "Geburtstagsrabatte aktivieren"
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
         TabIndex        =   37
         Top             =   3720
         Width           =   3615
      End
      Begin VB.CheckBox Check59 
         Caption         =   "'Kunde neu' mit Duplikatssuche"
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
         TabIndex        =   36
         Top             =   3480
         Width           =   3615
      End
      Begin VB.CheckBox Check53 
         Caption         =   "bei Zahlung immer DIN A4"
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
         TabIndex        =   35
         Top             =   4200
         Width           =   2775
      End
      Begin VB.CheckBox Check61 
         Caption         =   "ohne Bestandsprotokollierung"
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
         TabIndex        =   34
         Top             =   4680
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         Caption         =   "DIN A4 mit Firmenangaben im Rechnungsfuß"
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
         TabIndex        =   33
         Top             =   4440
         Width           =   4215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "bei Kundenwahl(Farbe ""rot"") mit PopUp Meldung "
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
         Left            =   4440
         TabIndex        =   32
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "bei Kundenwahl(ohne Email) mit PopUp Meldung "
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
         Left            =   4440
         TabIndex        =   31
         Top             =   960
         Width           =   4335
      End
      Begin VB.CheckBox Check86 
         Caption         =   "Stornobestätigung mit 2. Bediener"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   1920
         Width           =   3015
      End
      Begin VB.CheckBox Check65 
         Caption         =   "bei Gutscheineinlösung mit Details"
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
         Left            =   4440
         TabIndex        =   29
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CheckBox Check89 
         Caption         =   "Kasse mit Leiste 2 starten (nur dieser Computer)"
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
         Left            =   4440
         TabIndex        =   28
         Top             =   2160
         Width           =   4575
      End
      Begin VB.CheckBox Check78 
         Caption         =   "keine Schublade"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CheckBox Check75 
         Caption         =   "'Artikelrabatt halten' Funktion"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Schublade bei Kollegenverkäufen nicht öffnen"
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
         TabIndex        =   25
         Top             =   1800
         Width           =   4335
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Schublade beim ""Bargeld zählen"" öffnen"
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
         TabIndex        =   24
         Top             =   2040
         Width           =   4335
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Schublade bei Kundenbestellungen öffnen"
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
         TabIndex        =   23
         Top             =   1320
         Width           =   4335
      End
      Begin VB.CheckBox Check10 
         Caption         =   "nur Artikel mit Preis an der Kasse zulassen"
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
         Left            =   4440
         TabIndex        =   22
         Top             =   3120
         Width           =   3975
      End
      Begin VB.CheckBox Check11 
         Caption         =   "einfache Änderung der Kassennummer "
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
         TabIndex        =   21
         Top             =   3000
         Width           =   3855
      End
      Begin VB.CheckBox Check12 
         Caption         =   "automatische Artikelsuche an der Kasse zulassen"
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
         TabIndex        =   20
         Top             =   3240
         Width           =   4335
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Postleitzahl erfassen (mein Einzugsgebiet)"
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
         Left            =   4440
         TabIndex        =   19
         Top             =   360
         Width           =   4335
      End
      Begin VB.CheckBox Check16 
         Caption         =   "bei Zahlung DIN A4 zulassen"
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
         TabIndex        =   18
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Bargeldhöhe anzeigen (Bargeld zählen)"
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
         TabIndex        =   17
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   9240
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Bestandsdateien"
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
         TabIndex        =   15
         Top             =   4920
         Width           =   1815
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   5
         Left            =   9480
         TabIndex        =   12
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   11
         Left            =   10320
         TabIndex        =   161
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
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
         Caption         =   "S"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   12
         Left            =   9240
         TabIndex        =   164
         Top             =   3720
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "Hinzufügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   13
         Left            =   10560
         TabIndex        =   165
         Top             =   3720
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "bei erfolgter NV:"
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
         Index           =   45
         Left            =   4440
         TabIndex        =   244
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   "bei unbek. Artikeln:"
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
         Index           =   43
         Left            =   4440
         TabIndex        =   243
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label lbl6 
         Caption         =   "runden"
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
         Left            =   10200
         TabIndex        =   176
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label lbl6 
         Caption         =   "erst ab Warenkorbwert:"
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
         Left            =   9240
         TabIndex        =   175
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label lbl6 
         Caption         =   "Sind diese Artikel im Warenkorb enthalten, so wird der Endbetrag nicht gerundet."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   26
         Left            =   9240
         TabIndex        =   163
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lbl6 
         Caption         =   "Abrunden (keine 1 und 2 Cent mehr)"
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
         Index           =   25
         Left            =   9240
         TabIndex        =   160
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lbl6 
         Caption         =   "Email an diese Adresse"
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
         Left            =   4440
         TabIndex        =   137
         Top             =   5640
         Width           =   4455
      End
      Begin VB.Label lbl6 
         Caption         =   "Pfadangabe (Bestandlive.exe)"
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
         Left            =   480
         TabIndex        =   106
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label lbl6 
         Caption         =   "Rundungsrabatt bei Barverkäufen mit dieser ArtikelNr"
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
         Index           =   12
         Left            =   9240
         TabIndex        =   45
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame fraTasten 
      Caption         =   "Tasten an der Kasse deaktivieren / Sounds"
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   7680
         MaxLength       =   20
         TabIndex        =   247
         ToolTipText     =   "die ersten 15 Zeichen werden auf dem Kassenbon gedruckt"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check108 
         Caption         =   "Kasse/Artikelsuche: Artikelpositionen färben"
         Height          =   255
         Left            =   6000
         TabIndex        =   239
         Tag             =   "4"
         Top             =   3120
         Width           =   4095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "YabandPay"
         Height          =   255
         Index           =   22
         Left            =   3360
         TabIndex        =   236
         Tag             =   "17"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "PayPal"
         Height          =   255
         Index           =   21
         Left            =   3360
         TabIndex        =   235
         Tag             =   "17"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Google Pay"
         Height          =   255
         Index           =   20
         Left            =   3360
         TabIndex        =   234
         Tag             =   "17"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Apple Pay"
         Height          =   255
         Index           =   19
         Left            =   3360
         TabIndex        =   233
         Tag             =   "17"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "AliPay"
         Height          =   255
         Index           =   18
         Left            =   3360
         TabIndex        =   232
         Tag             =   "17"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Diners Club"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   231
         Tag             =   "17"
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Frame Frame4 
         Caption         =   "für die EC-Karte dieses Bild verwenden"
         Height          =   2535
         Left            =   6000
         TabIndex        =   215
         Top             =   3720
         Width           =   4455
         Begin VB.OptionButton Option1 
            Caption         =   "Maestro"
            Height          =   255
            Index           =   11
            Left            =   1080
            TabIndex        =   217
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "EC"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   216
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   16
            Left            =   9480
            TabIndex        =   218
            Top             =   6720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
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
            Caption         =   "Weiter"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command SSCommand8 
            Height          =   1095
            Index           =   0
            Left            =   2400
            TabIndex        =   219
            ToolTipText     =   "EC-Karte"
            Top             =   1200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1931
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
            ToolTipTitle    =   "EC-Karte"
            ButtonStyle     =   2
            Caption         =   "EC-Karte"
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command SSCommand8 
            Height          =   1095
            Index           =   4
            Left            =   120
            TabIndex        =   220
            ToolTipText     =   "EC-Karte"
            Top             =   1200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1931
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
            ToolTipTitle    =   "EC-Karte"
            ButtonStyle     =   2
            Caption         =   "EC-Karte"
            PictureAlign    =   3
            Version3        =   -1  'True
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sonstige"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   214
         Tag             =   "17"
         Top             =   6000
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "EC-Karte"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   213
         Tag             =   "17"
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "American Express"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   211
         Tag             =   "17"
         Top             =   5280
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Eurocard / Mastercard"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   210
         Tag             =   "17"
         Top             =   5040
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visa"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   209
         Tag             =   "17"
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         Caption         =   "Farbschema Warenkorb"
         Height          =   855
         Left            =   6000
         TabIndex        =   188
         Top             =   2160
         Width           =   4455
         Begin VB.OptionButton Option1 
            Caption         =   "Artikelpositionen färben aufgrund der Artikelfarbe"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   190
            Top             =   240
            Value           =   -1  'True
            Width           =   4215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Artikelpositionen färben aufgrund der Bedienerfarbe"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   189
            Top             =   480
            Width           =   4215
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   15
            Left            =   9480
            TabIndex        =   191
            Top             =   6720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
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
            Caption         =   "Weiter"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin VB.CheckBox Check34 
         Caption         =   "Restgutschein auch als Barauszahlung per Button 'Bar' zulassen"
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
         TabIndex        =   186
         Top             =   3480
         Width           =   5535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "EC Last"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   80
         Tag             =   "0"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Zahlung Gutsch"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   79
         Tag             =   "15"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Verkauf Gutsch"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   78
         Tag             =   "14"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Kartenverkauf"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   77
         Tag             =   "17"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sond Preis"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   76
         Tag             =   "2"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ges Rabatt"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   75
         Tag             =   "3"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Art Rabatt"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   74
         Tag             =   "4"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Kredit VK"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   73
         Tag             =   "5"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Schublade"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   72
         Tag             =   "33"
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Prov."
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   71
         Tag             =   "13"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2. Bon"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   70
         Tag             =   "34"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox Check57 
         Caption         =   "EC Lastschrift"
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
         TabIndex        =   69
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CheckBox Check56 
         Caption         =   "Scheck"
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
         Left            =   1560
         TabIndex        =   68
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox Check55 
         Caption         =   "Dukaten"
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
         TabIndex        =   67
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Unterbrechen/Fortsetzen"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   66
         Tag             =   "9"
         Top             =   2400
         Width           =   3135
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   4
         Left            =   9480
         TabIndex        =   10
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   7
         Left            =   6840
         TabIndex        =   131
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         Caption         =   "Test Ton"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   8
         Left            =   6360
         TabIndex        =   133
         Top             =   1320
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl6 
         Caption         =   "statt GZ"
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
         Index           =   46
         Left            =   7680
         TabIndex        =   248
         ToolTipText     =   "die ersten 15 Zeichen werden auf dem Kassenbon gedruckt"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lbl6 
         Caption         =   "Diese Kreditkarten aktivieren"
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
         Index           =   41
         Left            =   240
         TabIndex        =   212
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "bitte Farbe wählen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   32
         Left            =   4200
         TabIndex        =   134
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lbl6 
         Caption         =   "Testanzeige"
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
         Left            =   4200
         TabIndex        =   132
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lbl6 
         Caption         =   "Diesen Warnton abspielen bei Artikeln, die mit diesem Farbmerkmal versehen sind."
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
         Index           =   22
         Left            =   4200
         TabIndex        =   130
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lbl6 
         Caption         =   "Jugendschutz (Zigaretten/Alkohol)"
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
         Left            =   4200
         TabIndex        =   129
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lbl6 
         Caption         =   "bei gemischter Zahlung ausblenden:"
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
         Index           =   62
         Left            =   240
         TabIndex        =   81
         Top             =   2640
         Width           =   3375
      End
   End
   Begin VB.Frame fraWeiche 
      Caption         =   "Auswahl"
      Height          =   3015
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   6015
      Begin VB.OptionButton Option1 
         Caption         =   "Auszahlungsgründe"
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
         Index           =   4
         Left            =   240
         TabIndex        =   149
         Top             =   2280
         Width           =   5535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Kassenabschluss / Bargeldeingabe"
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
         Index           =   3
         Left            =   240
         TabIndex        =   139
         Top             =   1800
         Width           =   5535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "spezielle Kasseneinstellungen"
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
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "allgemeine Kasseneinstellungen"
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Value           =   -1  'True
         Width           =   5535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tasten an der Kasse deaktivieren / Sounds"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5535
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   3
         Left            =   9480
         TabIndex        =   5
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   6615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Einstellungen an der Kasse"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmWKL131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check107_Click()
On Error GoTo LOKAL_ERROR

    If Check107.value = vbChecked Then 'aktiviere
        Text1(11).Visible = True
    Else
        Text1(11).Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check107_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Check13_Click()
On Error GoTo LOKAL_ERROR

    If Check13.value = vbChecked Then
        Check35.Visible = True
    Else
        Check35.Visible = False
        Check35.value = vbUnchecked
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check13_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check15_Click()
On Error GoTo LOKAL_ERROR

    If Check15.value = vbChecked Then
        Check47.Visible = True
    Else
        Check47.Visible = False
        Check47.value = vbUnchecked
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check15_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check16_Click()
On Error GoTo LOKAL_ERROR

    If Check16.value = vbChecked Then
        Check53.Visible = True
    Else
        Check53.Visible = False
        Check53.value = vbUnchecked
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check16_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check18_Click()
On Error GoTo LOKAL_ERROR

    If Check18.value = vbChecked Then 'aktiviere
        Text1(5).Visible = True
        lbl6(11).Visible = True
    Else
        Text1(5).Visible = False
        lbl6(11).Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check18_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Check23_Click()
On Error GoTo LOKAL_ERROR

    If Check23.value = vbChecked Then 'aktiviere
        Text1(10).Visible = True
        lbl6(0).Visible = True
    Else
        Text1(10).Visible = False
        lbl6(0).Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check23_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check44_Click()
On Error GoTo LOKAL_ERROR

    If Check44.value = vbChecked Then
        DTPickerSMS.Visible = True
        Text1(30).Visible = True
    ElseIf Check44.value = vbUnchecked Then
        DTPickerSMS.Visible = False
        Text1(30).Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check44_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check45_Click()
On Error GoTo LOKAL_ERROR

    If Check45.value = vbChecked Then
        Dim sMessText As String
        Dim iRet As Integer
        sMessText = "Achtung: diese Einstellung lässt sich nicht rückgängig machen. Lesen Sie sich bitte über den Button 'Info' die Erläuterungen durch." & vbCrLf & vbCrLf
        sMessText = sMessText & "Möchten Sie jetzt wirklich diese Einstellung dauerhaft für Ihr Unternehmen einstellen?"
        
        iRet = MsgBox(sMessText, vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
        
        Else
            Check45.value = vbUnchecked
        End If
        
    ElseIf Check45.value = vbUnchecked Then
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check45_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim sSQL As String
    Dim byNr As Byte
    Dim sButext As String
    
    Select Case index
        Case 0
            Unload frmWKL131
        Case 1
            Kill App.Path & "\BUTTONS.CFG"
            
            loeschNEW "BUTTON", gdBase
            CreateTable "BUTTON", gdBase
            
            For i = 0 To 11
                If Check1(i).value = vbChecked Then
                    byNr = Check1(i).Tag
                    sButext = Check1(i).Caption
                    sSQL = "Insert into Button (indexnr,buttonnr,buttontext) values (" & i & "," & byNr & ",'" & sButext & "')"
                    gdBase.Execute sSQL, dbFailOnError
                End If
            Next i
            
            speicherKKaktiv
            speicherAUSBLEND
            speicherBonusgrenze
            speicherGeld
            speicherParknetto
            speicherJugendschutz
            speicherKKSCHUB
            speicherBARDINA4
            SpeicherAboPlus
            SpeicherBonusNrEingabe
            SpeicherBonusAutoAbzug
            speicherStornoCheckMit2Bed
            speicherBARDINA_41
            speicherJubi
            SpeicherRundRabArtnr
            SpeicherGeschenkSetArtnr
            SpeicherPrimLinr
            SpeicherZeitungsSpanne

            speicherCoupon
            speicherGutscheinnummernvergabe
            
            speicherBargeldEingabe
            
            speicherAuszahlungsgrund
            speicherRESTGU
            
        Case 2
            If Text1(0).Text <> "" Then
                List1.AddItem (Text1(0).Text)
                Text1(0).Text = ""
            End If
        Case 3
            'Vorauswahl
            If Option1(0).value = True Then
                fraWeiche.Visible = False
                fraTasten.Visible = True
                lbl6(23).Caption = ""
            ElseIf Option1(1).value = True Then
                fraWeiche.Visible = False
                fraAllg.Visible = True
            ElseIf Option1(2).value = True Then
                fraWeiche.Visible = False
                fraSpez.Visible = True
            ElseIf Option1(3).value = True Then
                fraWeiche.Visible = False
                fraKassenabschluss.Visible = True
            ElseIf Option1(4).value = True Then
                fraWeiche.Visible = False
                fraAuszahlungsgrund.Visible = True
            End If
        Case 4
            fraWeiche.Visible = True
            fraTasten.Visible = False
        Case 5
            fraWeiche.Visible = True
            fraAllg.Visible = False
        Case 6
            fraWeiche.Visible = True
            fraSpez.Visible = False
        Case 9
            fraWeiche.Visible = True
            fraKassenabschluss.Visible = False
        Case 10
            fraWeiche.Visible = True
            fraAuszahlungsgrund.Visible = False
        Case 7
        
            Dim yearAngabe As Integer
            yearAngabe = Year(Now) - 18
            anzeige "JUGENDSCHUTZ", "Volljährig bei Geburt vor " & Format(DateValue(Now) + 1, "dd.mm.") & yearAngabe, lbl6(23)
        Case 8
            gsBackcolor = Label4(32).BackColor
            gsForecolor = Label4(32).ForeColor
            gsArtikelFarbe = Label4(32).Tag
            
            frmWKL49.Show 1
            
            Label4(32).BackColor = gsBackcolor
            Label4(32).ForeColor = gsForecolor
            Label4(32).Tag = gsArtikelFarbe
            
            If gsArtikelFarbe <> "" Then
                Label4(32).Caption = "Farbauswahl"
            Else
                Label4(32).Caption = "bitte Farbe wählen"
            End If
        Case 11
            gcSuch = ""
            gsARTNR = ""
            frmWKL70.Show 1
            Me.Refresh
            If gsARTNR <> "" Then
                Text1(6).Text = gsARTNR
                gsARTNR = ""
            End If
        Case 12
            gcSuch = ""
            gsARTNR = ""
            frmWKL70.Show 1
            Me.Refresh
            If gsARTNR <> "" Then
                If Text1(17).Text = "" Then
                    Text1(17).Text = gsARTNR
                Else
                    Text1(17).Text = Text1(17).Text & vbCrLf & gsARTNR
                End If
                gsARTNR = ""
            End If
        Case 13
            Text1(17).Text = ""
        Case 17
            
            
            lblErlaeuterung.Caption = "Verehrte K.I.S.S.-Kunden," & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "der Gesetzgeber hat im Oktober 2018 beschlossen, dass sich ab dem 01. Januar 2019 unter bestimmten Umständen die Besteuerung von verkauften Gutscheinen ändert." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Bisher galt: Der Verkauf eines Gutscheines ist nicht umsatzsteuerpflichtig, erst bei der Einlösung des Gutscheines wird der Teil von dem Gutschein umsatzsteuerpflichtig, der eingelöst wurde." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Ab 2019 unterscheidet der Gesetzgeber zukünftig zwischen sogenannten Einzweck- und Mehrzweck-Gutscheinen." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Einzweck bedeutet hierbei: In dem Geschäft werden ausschließlich Waren und Dienstleistungen verkauft, die entweder dem vollen, dem ermäßigten oder gar keinem Mehrwertsteuersatz unterliegen." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Mehrzweck bedeutet: Es werden sowohl Produkte und Dienstleitungen verkauft, die dem ermäßigten, als auch dem vollem oder keinem Mehrwertsteuersatz unterliegen (Bsp.: Drogeriemarkt)." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Bei den 'Mehrzwecklern' unter Ihnen ändert sich durch die neue Regelung gar nichts. Sie müssen ihre Gutscheine nach der alten, bisherigen Regelung weiter versteuern. Von der neuen Regelung sind Sie nicht betroffen." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Bei den 'Einzwecklern' dreht sich der Zeitpunkt der Besteuerung um. Künftig müssen Sie die Mehrwertsteuer nach dem Verkauf des Gutscheines abführen. Wird der Gutschein ganz oder teilweise eingelöst, brauchen sie" & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "das vereinnahmte Geld dafür dann nicht mehr zu versteuern. Das Verfahren dreht sich also genau um. Nach wie vor gibt es natürlich keine Doppelbesteuerung." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Eine Parfümerie handelt gewöhnlich nur mit Artikeln, die zu 19% versteuert werden, es sei denn, es werden dort auch beispielsweise Fachbücher oder Fachzeitschriften verkauft. Dann gilt die Einzweckregelung nicht." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Was ändert sich künftig für die 'Einzweckler' in unserer Software? Für die gilt: Alles, was vor dem 01.01.2019 verkauft wurde, unterliegt der alten Regelung. Alles was danach verkauft wird, unterliegt der neuen Regelung." & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Dies ist nicht nur für uns ein enormer Aufwand, sondern auch für Ihre Buchhaltung bzw. Ihre Steuerberater." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Da wir nicht wissen können, ob Sie zu diesem Personenkreis gehören, müssen Sie selbst entscheiden, ob Sie am 01.01.2019 den Schalter auf 'Einzweck' umlegen müssen oder nicht. Sie werden ab dem 01.01.2019 bei jedem Programmstart" & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "danach gefragt, wie Sie sich entschieden haben, bis Sie sich entschieden haben.. Rückgängig machen können Sie die Entscheidung danach nicht mehr, dass können nur wir für Sie veranlassen. Also bitte überlegen Sie vorher gut, ob sich nicht" & vbCrLf
            
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "vielleicht doch das eine oder andere Produkt in Ihrem Sortiment befindet, was nur mit 7% MwSt. behaftet ist, oder mit dem Sie vielleicht ab 2019 handeln möchten. Besprechen Sie die neuen Regelungen am besten mit Ihrem Steuerberater." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Bei den Filialisten unter Ihnen können wir im Vorfeld den Schalter umlegen, so dass die Entscheidung in Ihren Filialen nicht von Ihren Mitarbeitern getroffen werden muss. Dazu sprechen Sie uns bitte an." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Wir werden die Programmänderung bis Ende Dezember an alle WINKISS und KISSLIVE Kunden verteilen. Eine Änderung in unserem Programm ZENTRALE werden wir nicht mehr vornehmen." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Bei Fragen hierzu kontaktieren Sie mich bitte, ich stehe Ihnen gerne zur Verfügung." & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Mit freundlichen Grüßen" & vbCrLf & vbCrLf
            lblErlaeuterung.Caption = lblErlaeuterung.Caption & "Carsten Schröder"
            
            Frame5.Visible = True
            lblErlaeuterung.Refresh
                        
        Case 18
            Frame5.Visible = False
        Case 19
            
            'Stichtag
            
            
            loeschNEW "STICHTAG", gdBase
            
            Check45.Visible = True
            Check45.Enabled = True
            Check45.value = vbUnchecked
            
            sSQL = "Update DBEINSTE Set GutscheinBeiVKversteuern = False"
            gdBase.Execute sSQL, dbFailOnError
            gbGutscheinBeiVKversteuern = False
            
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub speicherRESTGU()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update KASSEIN Set RESTGU = 0 "
    gdBase.Execute sSQL, dbFailOnError
    gdRESTGU = 0
    
    If Text6.Text <> "" Then
        If IsNumeric(Text6.Text) Then
            sSQL = "Update KASSEIN Set RESTGU = '" & Text6.Text & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            gdRESTGU = CDbl(Text6.Text)
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherRESTGU"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherStornoCheckMit2Bed()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check86.value = vbChecked Then
        sSQL = "Update DBEINSTE Set STORNOcheck2Bed = True "
        gdBase.Execute sSQL, dbFailOnError
        
        gbSTORNOcheck2Bed = True
    Else
        sSQL = "Update DBEINSTE Set STORNOcheck2Bed = False "
        gdBase.Execute sSQL, dbFailOnError
        
        gbSTORNOcheck2Bed = False
        
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherStornoCheckMit2Bed"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub speicherBARDINA_41()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    
    If Check108.value = vbChecked Then
        sSQL = "Update KASSEIN Set ArtsucheArtFarb = true"
        gdBase.Execute sSQL, dbFailOnError
        gbArtsucheArtFarb = True
    ElseIf Check108.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set ArtsucheArtFarb = False"
        gdBase.Execute sSQL, dbFailOnError
        gbArtsucheArtFarb = False
    End If
    

    
    If Check51.value = vbChecked Then
        sSQL = "Update KASSEIN Set KBmBI = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKBmBI = True
    ElseIf Check51.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KBmBI = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKBmBI = False
    End If
    
    If Check109(3).value = vbChecked Then
        sSQL = "Update KASSEIN Set NachKBbeiEC = true"
        gdBase.Execute sSQL, dbFailOnError
        gbNachKBbeiEC = True
    ElseIf Check109(3).value = vbUnchecked Then
        sSQL = "Update KASSEIN Set NachKBbeiEC = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNachKBbeiEC = False
    End If
    
    If Check65.value = vbChecked Then
        sSQL = "Update KASSEIN Set MGDETAILS = true"
        gdBase.Execute sSQL, dbFailOnError
        gbmGDetails = True
    ElseIf Check65.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set MGDETAILS = False"
        gdBase.Execute sSQL, dbFailOnError
        gbmGDetails = False
    End If
    
    If Check75.value = vbChecked Then
        sSQL = "Update KASSEIN Set ARTRABH = true"
        gdBase.Execute sSQL, dbFailOnError
        gbArtrabhalten = True
    ElseIf Check75.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set ARTRABH = False"
        gdBase.Execute sSQL, dbFailOnError
        gbArtrabhalten = False
    End If
    
    If Check66.value = vbChecked Then
        sSQL = "Update KASSEIN Set BEDLEER = true"
        gdBase.Execute sSQL, dbFailOnError
        gbBEDLEER = True
    ElseIf Check66.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set BEDLEER = False"
        gdBase.Execute sSQL, dbFailOnError
        gbBEDLEER = False
    End If
    
    
    
    
    'nicht mehr global
    If Check89.value = vbChecked Then
        sSQL = "Update WKEINSTE Set LEISTE2START = true"
        gdApp.Execute sSQL, dbFailOnError
        gbLeiste2Start = True
    ElseIf Check89.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set LEISTE2START = False"
        gdApp.Execute sSQL, dbFailOnError
        gbLeiste2Start = False
    End If
    
    If Check105.value = vbChecked Then
        sSQL = "Update KASSEIN Set IDENTUSER = true"
        gdBase.Execute sSQL, dbFailOnError
        gbIdentUser = True
    ElseIf Check105.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set IDENTUSER = False"
        gdBase.Execute sSQL, dbFailOnError
        gbIdentUser = False
    End If
    
    If Check78.value = vbChecked Then
        sSQL = "Update WKEINSTE Set NOBONDRUCKER = true"
        gdApp.Execute sSQL, dbFailOnError
        gbNOBONDRUCKER = True
    ElseIf Check78.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set NOBONDRUCKER = False"
        gdApp.Execute sSQL, dbFailOnError
        gbNOBONDRUCKER = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBARDINA_41"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub speicherBARDINA4()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check3.value = vbChecked Then
        sSQL = "Update WKEINSTE Set DINA4RECHFU = true"
        gdApp.Execute sSQL, dbFailOnError
        gbDINA4RECHFU = True
    ElseIf Check3.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set DINA4RECHFU = False"
        gdApp.Execute sSQL, dbFailOnError
        gbDINA4RECHFU = False
    End If
    
    If Check16.value = vbChecked Then
        sSQL = "Update WKEINSTE Set DINA4VIS = true"
        gdApp.Execute sSQL, dbFailOnError
        gbDINA4VIS = True
    ElseIf Check16.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set DINA4VIS = False"
        gdApp.Execute sSQL, dbFailOnError
        gbDINA4VIS = False
    End If
    
    If Check53.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BARDINA4 = true"
        gdApp.Execute sSQL, dbFailOnError
        gbBARDINA4 = True
    ElseIf Check53.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BARDINA4 = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBARDINA4 = False
    End If
    
    If Check15.value = vbChecked Then
        sSQL = "Update WKEINSTE Set ZWEITMONI = true"
        gdApp.Execute sSQL, dbFailOnError
        gbZweitMoni = True
    ElseIf Check15.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set ZWEITMONI = False"
        gdApp.Execute sSQL, dbFailOnError
        gbZweitMoni = False
    End If
    
    If Check47.value = vbChecked Then
        sSQL = "Update WKEINSTE Set ZWEITMONIMINI = true"
        gdApp.Execute sSQL, dbFailOnError
        gbZweitMoniMinimieren = True
    ElseIf Check47.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set ZWEITMONIMINI = False"
        gdApp.Execute sSQL, dbFailOnError
        gbZweitMoniMinimieren = False
    End If
    
    If Check59.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUDU = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKUDU = True
    ElseIf Check59.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUDU = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUDU = False
    End If
    
    If Check60.value = vbChecked Then
        sSQL = "Update KASSEIN Set GEBRABK = true"
        gdBase.Execute sSQL, dbFailOnError
        gbGEBRABK = True
    ElseIf Check60.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set GEBRABK = False"
        gdBase.Execute sSQL, dbFailOnError
        gbGEBRABK = False
    End If
    
    If Check61.value = vbChecked Then
        sSQL = "Update KASSEIN Set OhnebestProt = true"
        gdBase.Execute sSQL, dbFailOnError
        gbOhnebestProt = True
    ElseIf Check61.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set OhnebestProt = False"
        gdBase.Execute sSQL, dbFailOnError
        gbOhnebestProt = False
    End If
    
    If Check48.value = vbChecked Then
        sSQL = "Update KASSEIN Set KeineBestVerWarengru = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKeineBestVerWarengru = True
    ElseIf Check48.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KeineBestVerWarengru = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKeineBestVerWarengru = False
    End If
    
    If Check23.value = vbChecked Then
        sSQL = "Update KASSEIN Set BestDateien = true"
        gdBase.Execute sSQL, dbFailOnError
        gbBestDateien = True
    ElseIf Check23.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set BestDateien = False"
        gdBase.Execute sSQL, dbFailOnError
        gbBestDateien = False
    End If
    
    If Check20.value = vbChecked Then
        sSQL = "Update KASSEIN Set BarAnz = true"
        gdBase.Execute sSQL, dbFailOnError
        gbBarAnz = True
    ElseIf Check20.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set BarAnz = False"
        gdBase.Execute sSQL, dbFailOnError
        gbBarAnz = False
    End If
    
    If Check22.value = vbChecked Then
        sSQL = "Update KASSEIN Set EinfacheZollErstattung = true"
        gdBase.Execute sSQL, dbFailOnError
        gbEinfacheZollErstattung = True
    ElseIf Check22.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set EinfacheZollErstattung = False"
        gdBase.Execute sSQL, dbFailOnError
        gbEinfacheZollErstattung = False
    End If
    
    If Check62.value = vbChecked Then
        sSQL = "Update WKEINSTE Set UmsAnz = true"
        gdApp.Execute sSQL, dbFailOnError
        gbUmsAnz = True
    ElseIf Check62.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set UmsAnz = False"
        gdApp.Execute sSQL, dbFailOnError
        gbUmsAnz = False
    End If

    If Check11.value = vbChecked Then
        sSQL = "Update WKEINSTE Set EDITKASSNR = true"
        gdApp.Execute sSQL, dbFailOnError
        gbEDITKASSNR = True
    ElseIf Check11.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set EDITKASSNR = False"
        gdApp.Execute sSQL, dbFailOnError
        gbEDITKASSNR = False
    End If
    
    If Check12.value = vbChecked Then
        sSQL = "Update KASSEIN Set AUTOSEEK = true"
        gdBase.Execute sSQL, dbFailOnError
        gbArtikelTextSuche = True
    ElseIf Check12.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set AUTOSEEK = False"
        gdBase.Execute sSQL, dbFailOnError
        gbArtikelTextSuche = False
    End If
    
    If Check39.value = vbChecked Then
        sSQL = "Update KASSEIN Set MitKundeWahlHinweis = true"
        gdBase.Execute sSQL, dbFailOnError
        gbMitKundeWahlHinweis = True
    ElseIf Check39.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set MitKundeWahlHinweis = False"
        gdBase.Execute sSQL, dbFailOnError
        gbMitKundeWahlHinweis = False
    End If
    
    If Check13.value = vbChecked Then
        sSQL = "Update KASSEIN Set PLZGEBIET = true"
        gdBase.Execute sSQL, dbFailOnError
        gbPLZGEBIET = True
        
        If NewTableSuchenDBKombi("PLZGEBIET", gdBase) = False Then
            loeschNEW "PLZGEBIET", gdBase
            CreateTableT2 "PLZGEBIET", gdBase
        End If
    
    ElseIf Check13.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set PLZGEBIET = False"
        gdBase.Execute sSQL, dbFailOnError
        gbPLZGEBIET = False
    End If
    
    If Check35.value = vbChecked Then
        sSQL = "Update KASSEIN Set PLZGEBIET_AuchBeiKUWAHL = true"
        gdBase.Execute sSQL, dbFailOnError
        gbPLZGEBIET_AuchBeiKUWAHL = True
    ElseIf Check35.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set PLZGEBIET_AuchBeiKUWAHL = False"
        gdBase.Execute sSQL, dbFailOnError
        gbPLZGEBIET_AuchBeiKUWAHL = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBARDINA4"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherGutscheinnummernvergabe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If opt1(4).value = True Then
        sSQL = "Update DBEINSTE Set HaGuNr = true "
        gdBase.Execute sSQL, dbFailOnError
        gbGutsch = True
    ElseIf opt1(5).value = True Then
        sSQL = "Update DBEINSTE Set HaGuNr = False "
        gdBase.Execute sSQL, dbFailOnError
        gbGutsch = False
    End If
    
    If Check42.value = vbChecked Then
        sSQL = "Update DBEINSTE Set OGV = true"
        gdBase.Execute sSQL, dbFailOnError
        gbOGV = True
    ElseIf Check42.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set OGV = False"
        gdBase.Execute sSQL, dbFailOnError
        gbOGV = False
    End If
    
    If Check43.value = vbChecked Then
        sSQL = "Update DBEINSTE Set GutschnrKomplett = true"
        gdBase.Execute sSQL, dbFailOnError
        gbGutschnrKomplett = True
    ElseIf Check43.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set GutschnrKomplett = False"
        gdBase.Execute sSQL, dbFailOnError
        gbGutschnrKomplett = False
    End If
    
    If Check45.Visible = True Then
    
        If Check45.value = vbChecked Then
        
            sSQL = "Update DBEINSTE Set GutscheinBeiVKversteuern = true"
            gdBase.Execute sSQL, dbFailOnError
            gbGutscheinBeiVKversteuern = True
            
            'Datum
            'Tabelle: Stichtag
            'Spalte: Datum
            
            loeschNEW "STICHTAG", gdBase
            CreateTableT3 "STICHTAG", gdBase
            
            Dim dateTag As Date
            dateTag = DateValue(Now)
            
            sSQL = "Insert into STICHTAG (Datum) values  "
            sSQL = sSQL & " ( '" & dateTag & "') "
            gdBase.Execute sSQL, dbFailOnError
            
            Check45.Enabled = False
    
        ElseIf Check45.value = vbUnchecked Then
        
            sSQL = "Update DBEINSTE Set GutscheinBeiVKversteuern = False"
            gdBase.Execute sSQL, dbFailOnError
            gbGutscheinBeiVKversteuern = False
        End If
        
    End If
    
    
    
    
    
    
    
    
    
    If Check46.value = vbChecked Then
        sSQL = "Update DBEINSTE Set RGO = true"
        gdBase.Execute sSQL, dbFailOnError
        gbRGO = True
    ElseIf Check46.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set RGO = False"
        gdBase.Execute sSQL, dbFailOnError
        gbRGO = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherGutscheinnummernvergabe"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherCoupon()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Check25.value = vbChecked Then
        sSQL = "Update KASSEIN Set Coupon = true"
        gdBase.Execute sSQL, dbFailOnError
        gbCoupon = True

    ElseIf Check25.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set Coupon = False"
        gdBase.Execute sSQL, dbFailOnError
        gbCoupon = False
    End If
    
    If Check26.value = vbChecked Then
        sSQL = "Update KASSEIN Set GuStattBar = true"
        gdBase.Execute sSQL, dbFailOnError
        gbGuStattBar = True

    ElseIf Check26.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set GuStattBar = False"
        gdBase.Execute sSQL, dbFailOnError
        gbGuStattBar = False
    End If
    
    If Check37.value = vbChecked Then
        sSQL = "Update KASSEIN Set MitStaffelPreis = true"
        gdBase.Execute sSQL, dbFailOnError
        gbMitStaffelPreis = True

    ElseIf Check37.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set MitStaffelPreis = False"
        gdBase.Execute sSQL, dbFailOnError
        gbMitStaffelPreis = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherCoupon"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speicherAUSBLEND()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Datendrin("KASSEIN", gdBase) = False Then
    
        sSQL = "Insert into KASSEIN  (AUSBLDU) values (False)"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    If Check6.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUBONUS = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUBONUS = True
    ElseIf Check6.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUBONUS = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUBONUS = False
    End If
    
    If Check107.value = vbChecked Then
        sSQL = "Update KASSEIN Set MBBLOCKFrage = True"
        gdBase.Execute sSQL, dbFailOnError
        gbMBBLOCKFrage = True
    ElseIf Check107.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set MBBLOCKFrage = False"
        gdBase.Execute sSQL, dbFailOnError
        gbMBBLOCKFrage = False
    End If
    
    If Text1(11).Text <> "" Then
        gsSperrFrage = Text1(11).Text
    End If
    
    If gsSperrFrage <> "" Then
        sSQL = "Update KASSEIN Set SperrFrage = '" & gsSperrFrage & "'"
        gdBase.Execute sSQL, dbFailOnError
       
    Else
        sSQL = "Update KASSEIN Set SperrFrage = '' "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Check19.value = vbChecked Then
        sSQL = "Update KASSEIN Set NOKUBONUS_AGRAB = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNoKUBONUS_wenn_Art_and_Ges_rab = True
    ElseIf Check19.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set NOKUBONUS_AGRAB = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNoKUBONUS_wenn_Art_and_Ges_rab = False
    End If
    
    If Check18.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUBONUS_WENN = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUBONUS_WENN = True
    ElseIf Check18.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUBONUS_WENN = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUBONUS_WENN = False
    End If
    
    gsiKUBONUS_SCHWELLE = 0
    If Text1(5).Text <> "" Then
        If IsNumeric(Text1(5).Text) Then
            gsiKUBONUS_SCHWELLE = CSng(Text1(5).Text)
        End If
    End If
    
    If gsiKUBONUS_SCHWELLE > 0 Then
        sSQL = "Update KASSEIN Set KUBONUS_SCHWELLE = '" & gsiKUBONUS_SCHWELLE & "'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update KASSEIN Set KUBONUS_SCHWELLE = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Check5.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUWAHLMAIL = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLMAIL = True
    ElseIf Check5.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUWAHLMAIL = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLMAIL = False
    End If
    
    If Check4.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUWAHLROT = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLROT = True
    ElseIf Check4.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUWAHLROT = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLROT = False
    End If
    
    
    If Check31.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUWAHLGESPERRTROT = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLGESPERRTROT = True
    ElseIf Check31.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUWAHLGESPERRTROT = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLGESPERRTROT = False
    End If
    
    
    
    
    
    
    If Check30.value = vbChecked Then
        sSQL = "Update KASSEIN Set KUWAHLfbimmer = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLfbimmer = True
    ElseIf Check30.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KUWAHLfbimmer = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKUWAHLfbimmer = False
    End If
    
    If Check55.value = vbChecked Then
        sSQL = "Update KASSEIN Set AUSBLDU = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAUSBLDU = True
        
    ElseIf Check55.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set AUSBLDU = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAUSBLDU = False
        
    End If
    
    If Check56.value = vbChecked Then
        sSQL = "Update KASSEIN Set AUSBLsh = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAUSBLSH = True
        
    ElseIf Check56.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set AUSBLsh = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAUSBLSH = False
        
    End If
    
    If Check57.value = vbChecked Then
        sSQL = "Update KASSEIN Set AUSBLls = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAUSBLLS = True
        
    ElseIf Check57.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set AUSBLls = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAUSBLLS = False
        
    End If
    
    If Check69.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONNEIN = true"
        gdApp.Execute sSQL, dbFailOnError
        gbBONNEIN = True
    ElseIf Check69.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BONNEIN = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBONNEIN = False
    End If
    
    If Check70.value = vbChecked Then
        sSQL = "Update KASSEIN Set BONWAHL = true"
        gdBase.Execute sSQL, dbFailOnError
        gbBONWAHL = True
    ElseIf Check70.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set BONWAHL = False"
        gdBase.Execute sSQL, dbFailOnError
        gbBONWAHL = False
    End If
    
    If Check10.value = vbChecked Then
        sSQL = "Update KASSEIN Set MITPREIS = true"
        gdBase.Execute sSQL, dbFailOnError
        gbMitPreis = True
    ElseIf Check10.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set MITPREIS = False"
        gdBase.Execute sSQL, dbFailOnError
        gbMitPreis = False
    End If
    
    If Check71.value = vbChecked Then
        sSQL = "Update KASSEIN Set kassgefuehrt = true"
        gdBase.Execute sSQL, dbFailOnError
        gbkassgefuehrt = True
    ElseIf Check71.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set kassgefuehrt = False"
        gdBase.Execute sSQL, dbFailOnError
        gbkassgefuehrt = False
    End If
    
    If Check63.value = vbChecked Then
        sSQL = "Update KASSEIN Set RABVS = true"
        gdBase.Execute sSQL, dbFailOnError
        gbRabVs = True
    ElseIf Check63.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set RABVS = False"
        gdBase.Execute sSQL, dbFailOnError
        gbRabVs = False
    End If
    
    
    
    If Check2.value = vbChecked Then
        sSQL = "Update KASSEIN Set OpenSchubRetoure = True"
        gdBase.Execute sSQL, dbFailOnError
        gbOpenSchubRetoure = True
    ElseIf Check2.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set OpenSchubRetoure = False"
        gdBase.Execute sSQL, dbFailOnError
        gbOpenSchubRetoure = False
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
Private Sub speicherBargeldEingabe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Option2(0).value = True Then 'frmwk21b
        giBARGELDART = 0
    ElseIf Option2(1).value = True Then 'frmwk21s
        giBARGELDART = 1
    ElseIf Option2(2).value = True Then 'frmwk21t
        giBARGELDART = 2
    End If
    
    sSQL = "Update KASSEIN Set BARGELDART = " & giBARGELDART & " "
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBargeldEingabe"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub speicherAuszahlungsgrund()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim i       As Integer
    Dim sGrund  As String
    
    loeschNEW "AUSZAHLUNGSGRUND", gdBase
    CreateTableT2 "AUSZAHLUNGSGRUND", gdBase
    
    For i = 0 To 9
        sGrund = Text5(i).Text
        
        If sGrund <> "" Then
        
            sSQL = "Insert into Auszahlungsgrund (GRUNDNR,Auszahlungsgrund,Filiale) values (0,'" & sGrund & "','" & gcFilNr & "')"
            gdBase.Execute sSQL, dbFailOnError
        End If
    
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAuszahlungsgrund"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub leseAuszahlungsgrund()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim i       As Integer
    Dim sGrund  As String
    Dim rsrs As DAO.Recordset
    
    If NewTableSuchenDBKombi("AUSZAHLUNGSGRUND", gdBase) = True Then

        sSQL = "Select * from AUSZAHLUNGSGRUND"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
        
                sGrund = ""
                If Not IsNull(rsrs!AUSZAHLUNGSGRUND) Then
                    sGrund = Trim(rsrs!AUSZAHLUNGSGRUND)
                End If
                
                If sGrund <> "" Then
                    For i = 0 To 9
                        If Text5(i).Text = "" Then
                            Text5(i).Text = sGrund
                            Exit For
                        End If
                    Next i
                End If
                rsrs.MoveNext
            Loop
        End If

        rsrs.Close: Set rsrs = Nothing
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leseAuszahlungsgrund"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub
Private Sub SpeicherAboPlus()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim i As Integer
    
    loeschNEW "ABOPLUS", gdBase
    
    
    
    If List1.ListCount > 0 Then
        
        If Text2.Text <> "" Then
            CreateTableT2 "ABOPLUS", gdBase
            
            For i = 0 To List1.ListCount - 1
            
                sSQL = "Delete from  ABOPLUS where ARTNR = " & CLng(Trim(List1.list(i))) & " "
                gdBase.Execute sSQL, dbFailOnError
            
                sSQL = "Insert into ABOPLUS (ARTNR) values (" & CLng(Trim(List1.list(i))) & ")"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "UPDATE ABOPLUS set PFNR = " & Text2.Text & " "
                gdBase.Execute sSQL, dbFailOnError
            Next i
        Else
            'Partnerfirmennummer eintragen
            anzeige "rot", "Partnerfirmennummer eintragen", Label1(4)
            Text2.SetFocus
        End If
    
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAboPlus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherBonusNrEingabe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "BONUSNRE", gdBase
    
    If Text1(2).Text <> "" Then
        CreateTableT2 "BONUSNRE", gdBase
        
        sSQL = "Insert into BONUSNRE (ARTNR) values (" & CLng(Trim(Text1(2).Text)) & ")"
        gdBase.Execute sSQL, dbFailOnError
            
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherBonusNrEingabe"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherRundRabArtnr()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    
    glAutoAusSchFiliale = 0
    If Text1(25).Text <> "" Then
        If IsNumeric(Text1(25).Text) Then
    
            sSQL = "Update KASSEIN Set AutoAusSchFiliale = " & Trim(Text1(25).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glAutoAusSchFiliale = Trim(Text1(25).Text)
        
        End If
    Else
    
        sSQL = "Update KASSEIN Set AutoAusSchFiliale = 0"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    
    glAutoKundnrforKundBest = 0
    If Text1(19).Text <> "" Then
        If IsNumeric(Text1(19).Text) Then
    
            sSQL = "Update KASSEIN Set AutoKundnrforKundBest = " & Trim(Text1(19).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glAutoKundnrforKundBest = Trim(Text1(19).Text)
        
        End If
    Else
    
        sSQL = "Update KASSEIN Set AutoKundnrforKundBest = 0"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    
    
    
    
    
    
   
    glRRArtnr = 0
    If Text1(6).Text <> "" Then
        If IsNumeric(Text1(6).Text) Then
    
            sSQL = "Update KASSEIN Set RRArtnr = " & Trim(Text1(6).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glRRArtnr = Trim(Text1(6).Text)
        
        End If
    Else
    
        sSQL = "Update KASSEIN Set RRArtnr = 0"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    

    glBonusGrenzeArtnr = 0
    If Text1(24).Text <> "" Then
        If IsNumeric(Text1(24).Text) Then
    
            sSQL = "Update KASSEIN Set BonusGrenzeArtnr = " & Trim(Text1(24).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glBonusGrenzeArtnr = Trim(Text1(24).Text)
        
        End If
    Else
    
        sSQL = "Update KASSEIN Set BonusGrenzeArtnr = 0"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    
    
    
    glBonusAuszahlungArtnr = 0
    If Text1(28).Text <> "" Then
        If IsNumeric(Text1(28).Text) Then
    
            sSQL = "Update KASSEIN Set BonusAuszahlungArtnr = " & Trim(Text1(28).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glBonusAuszahlungArtnr = Trim(Text1(28).Text)
        
        End If
    Else
    
        sSQL = "Update KASSEIN Set BonusAuszahlungArtnr = 0"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    

    
    
    
    
    
    
    
    
    
    If Check32.value = vbChecked Then
        sSQL = "Update KASSEIN Set NurBonusfRunden = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNurBonusfRunden = True
    ElseIf Check32.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set NurBonusfRunden = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNurBonusfRunden = False
    End If
    
    If Check33.value = vbChecked Then
        sSQL = "Update KASSEIN Set NachfragenbeiWGNohnePreis = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNachfragenbeiWGNohnePreis = True
    ElseIf Check33.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set NachfragenbeiWGNohnePreis = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNachfragenbeiWGNohnePreis = False
    End If
    
    
    If Option1(10).value = True Then
        gsFARBKASSE = "1"
        sSQL = "Update KASSEIN Set FARBKASSE = 1"
        gdBase.Execute sSQL, dbFailOnError
    ElseIf Option1(7).value = True Then
        gsFARBKASSE = "2"
        sSQL = "Update KASSEIN Set FARBKASSE = 2"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Option1(6).value = True Then
        gsECBILD = "1"
        sSQL = "Update KASSEIN Set ECBILD = 1"
        gdBase.Execute sSQL, dbFailOnError
    ElseIf Option1(11).value = True Then
        gsECBILD = "2"
        sSQL = "Update KASSEIN Set ECBILD = 2"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    If Option1(8).value = True Then
        gsAbrunden = "1"
        sSQL = "Update KASSEIN Set RUNDEN = 1"
        gdBase.Execute sSQL, dbFailOnError
    ElseIf Option1(9).value = True Then
        gsAbrunden = "2"
        sSQL = "Update KASSEIN Set RUNDEN = 2"
        gdBase.Execute sSQL, dbFailOnError
    ElseIf Option1(5).value = True Then
        gsAbrunden = "3"
        sSQL = "Update KASSEIN Set RUNDEN = 3"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Update KASSEIN Set SCHWELLEWK = 0 "
    gdBase.Execute sSQL, dbFailOnError
    gdSCHWELLEWK = 0
    
    If Text1(18).Text <> "" Then
        If IsNumeric(Text1(18).Text) Then
            sSQL = "Update KASSEIN Set SCHWELLEWK = '" & Text1(18).Text & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            gdSCHWELLEWK = CDbl(Text1(18).Text)
        End If
    End If
    
    If Check44.value = vbChecked Then
        sSQL = "Update DBEINSTE Set TerminReminderSMS = true"
        gdBase.Execute sSQL, dbFailOnError
        gbTerminReminderSMS = True
        
        
        'zu welcher Uhrzeit
        
        
        Dim sSMSstart As String
        sSMSstart = Right(DTPickerSMS.value, 8)
        
        If sSMSstart <> "" Then
            gsTerminReminderstart = sSMSstart
            
            sSQL = "Update WKEINSTE set TerminReminderstart  = '" & sSMSstart & "'"
            gdApp.Execute sSQL, dbFailOnError
            
        Else
            gsTerminReminderstart = ""
            
            sSQL = "Update WKEINSTE set TerminReminderstart  = ''"
            gdApp.Execute sSQL, dbFailOnError
        End If
        
        
        
        'tage vorher
        glTageVorTermin = 0
        If Text1(30).Text <> "" Then
            If IsNumeric(Text1(30).Text) Then
        
                sSQL = "Update WKEINSTE Set TageVorTermin = " & Trim(Text1(30).Text)
                gdApp.Execute sSQL, dbFailOnError
                
                glTageVorTermin = Trim(Text1(30).Text)
            
            End If
        Else
        
            sSQL = "Update WKEINSTE Set TageVorTermin = 2"
            gdApp.Execute sSQL, dbFailOnError
        
        End If
        
        
        
        
        
    ElseIf Check44.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set TerminReminderSMS = False"
        gdBase.Execute sSQL, dbFailOnError
        gbTerminReminderSMS = False
        
        sSQL = "Update WKEINSTE set TerminReminderstart  = ''"
        gdApp.Execute sSQL, dbFailOnError
        
        sSQL = "Update WKEINSTE Set TageVorTermin = 2"
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherRundRabArtnr"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherZeitungsSpanne()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    gdZeitungsSpanne = 0
    If Text1(8).Text <> "" Then
        If IsNumeric(Text1(8).Text) Then
    
            sSQL = "Update KASSEIN Set ZeitungsSpanne = '" & Trim(Text1(8).Text) & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            gdZeitungsSpanne = Trim(Text1(8).Text)
        
        End If
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherZeitungsSpanne"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherGeschenkSetArtnr()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    glGSArtnr = 0
    If Text1(7).Text <> "" Then
        If IsNumeric(Text1(7).Text) Then
    
            sSQL = "Update KASSEIN Set GSArtnr = " & Trim(Text1(7).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glGSArtnr = Trim(Text1(7).Text)
        
        End If
    End If
    
    
   
    glZehnProzLinr = 0
    If Text1(38).Text <> "" Then
        If IsNumeric(Text1(38).Text) Then
    
            sSQL = "Update KASSEIN Set ZehnProzLinr = " & Trim(Text1(38).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glZehnProzLinr = Trim(Text1(38).Text)
        
        End If
    End If
    
    glZehnProzArtnr = 0
    If Text1(39).Text <> "" Then
        If IsNumeric(Text1(39).Text) Then
    
            sSQL = "Update KASSEIN Set ZehnProzArtnr = " & Trim(Text1(39).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glZehnProzArtnr = Trim(Text1(39).Text)
        
        End If
    End If
    
    
    
    glBaganzArtnr = 0
    If Text1(32).Text <> "" Then
        If IsNumeric(Text1(32).Text) Then
    
            sSQL = "Update KASSEIN Set BaganzArtnr = " & Trim(Text1(32).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glBaganzArtnr = Trim(Text1(32).Text)
        
        End If
    End If
    
    
    glBaganzAR = 0
    If Text1(33).Text <> "" Then
        If IsNumeric(Text1(33).Text) Then
    
            sSQL = "Update KASSEIN Set BaganzAr = " & Trim(Text1(33).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glBaganzAR = Trim(Text1(33).Text)
        
        End If
    End If
    
    'ab hier WarenkorbWertRabatt
    
    
    
    gdWarenkorbWert = 0
    If Text1(36).Text <> "" Then
        If IsNumeric(Text1(36).Text) Then
        
            gdWarenkorbWert = Trim(Text1(36).Text)
    
            sSQL = "Update KASSEIN Set WarenkorbWert = '" & gdWarenkorbWert & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
        
        End If
    End If
    
    
    gdWarenkorbGR = 0
    If Text1(37).Text <> "" Then
        If IsNumeric(Text1(37).Text) Then
        
            gdWarenkorbGR = Trim(Text1(37).Text)
    
            sSQL = "Update KASSEIN Set WarenkorbGR = '" & gdWarenkorbGR & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
        
        End If
    End If
    
    
    
    
    
    
    'ab hier Liebling
    
    
    
    glLieblingArtnr = 0
    If Text1(35).Text <> "" Then
        If IsNumeric(Text1(35).Text) Then
    
            sSQL = "Update KASSEIN Set LieblingArtnr = " & Trim(Text1(35).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glLieblingArtnr = Trim(Text1(35).Text)
        
        End If
    End If
    
    
    glLieblingAR = 0
    If Text1(34).Text <> "" Then
        If IsNumeric(Text1(34).Text) Then
    
            sSQL = "Update KASSEIN Set LieblingAR = " & Trim(Text1(34).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glLieblingAR = Trim(Text1(34).Text)
        
        End If
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherGeschenkSetArtnr"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherPrimLinr()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    glPrimLinr = 0
    If Text1(9).Text <> "" Then
        If IsNumeric(Text1(9).Text) Then
    
            sSQL = "Update KASSEIN Set PrimLinr = " & Trim(Text1(9).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glPrimLinr = Trim(Text1(9).Text)
        
        End If
    End If
    
    glZeitungsLinr = 0
    If Text1(21).Text <> "" Then
        If IsNumeric(Text1(21).Text) Then
    
            sSQL = "Update KASSEIN Set ZeitungsLinr = " & Trim(Text1(21).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glZeitungsLinr = Trim(Text1(21).Text)
        
        End If
    End If
    
    
    
    glPaketLinr = 0
    If Text1(12).Text <> "" Then
        If IsNumeric(Text1(12).Text) Then
    
            sSQL = "Update KASSEIN Set PaketLinr = " & Trim(Text1(12).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glPaketLinr = Trim(Text1(12).Text)
        
            If NewTableSuchenDBKombi("PAKET", gdBase) = False Then
                sSQL = "Create Table Paket (Preis double, sendok Bit)"
                gdBase.Execute sSQL, dbFailOnError
            End If
        
        End If
    End If
    
    glECAuszahlArtnr = 0
    If Text1(26).Text <> "" Then
        If IsNumeric(Text1(26).Text) Then
    
            sSQL = "Update KASSEIN Set ECAuszahlArtnr = " & Trim(Text1(26).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glECAuszahlArtnr = Trim(Text1(26).Text)
        End If
    End If
    
    
    
    glSpezLottoauszahlartikel = 0
    If Text1(22).Text <> "" Then
        If IsNumeric(Text1(22).Text) Then
    
            sSQL = "Update KASSEIN Set SpezLottoauszahlartikel = " & Trim(Text1(22).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glSpezLottoauszahlartikel = Trim(Text1(22).Text)
        End If
    End If
    
    glSpezFotoartikel = 0
    If Text1(14).Text <> "" Then
        If IsNumeric(Text1(14).Text) Then
    
            sSQL = "Update KASSEIN Set spezFotoartikel = " & Trim(Text1(14).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            glSpezFotoartikel = Trim(Text1(14).Text)
        End If
    End If
    
    
    
    
    
    gsSpezArtikel = ""
    If Text1(13).Text <> "" Then
    
        If Right(Text1(13).Text, 1) <> "," Then
            Text1(13).Text = Text1(13).Text & ","
        End If
    
        sSQL = "Update KASSEIN Set SPEZARTIKEL = '" & Trim(SwapStr(Text1(13).Text, ",", "$")) & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        gsSpezArtikel = Trim(Text1(13).Text)
    End If
    
    gsRabattAusnahmeArtikel = ""
    If Text1(17).Text <> "" Then
    
        sSQL = "Update KASSEIN Set RabattAusnahmeArtikel = '" & Trim(SwapStr(Text1(17).Text, vbCrLf, "$")) & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        gsRabattAusnahmeArtikel = Trim(Text1(17).Text)
    End If
    
    

    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherPrimLinr"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherBonusAutoAbzug()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "BONUSAA", gdBase
    
    If Text1(1).Text <> "" Then
        CreateTableT2 "BONUSAA", gdBase
        
        sSQL = "Insert into BONUSAA (ARTNR) values (" & CLng(Trim(Text1(1).Text)) & ")"
        gdBase.Execute sSQL, dbFailOnError
            
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherBonusAutoAbzug"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim i       As Integer
    Dim rsrs    As Recordset

    Positionieren131
    Modul6.Skalieren_Kasse Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    anzeige "normal", "", Label1(4)
    
    If NewTableSuchenDBKombi("BUTTON", gdBase) Then
        lesebutton
    End If
    
    Dim cPfad   As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Set SSCommand8(4).Picture = LoadPicture(cPfad & "Picture\System\" & "EC.jpg")
    Set SSCommand8(0).Picture = LoadPicture(cPfad & "Picture\System\" & "Maestro.jpg")

    
    
    leseAuszahlungsgrund
    
    If gbKK_Visa = True Then
        Check1(12).value = vbChecked
    Else
        Check1(12).value = vbUnchecked
    End If
    
    If gbKK_EurocardMastercard = True Then
        Check1(13).value = vbChecked
    Else
        Check1(13).value = vbUnchecked
    End If
    
    If gbKK_AmericanExpress = True Then
        Check1(14).value = vbChecked
    Else
        Check1(14).value = vbUnchecked
    End If
    
    If gbKK_DinersClub = True Then
        Check1(15).value = vbChecked
    Else
        Check1(15).value = vbUnchecked
    End If
    
    If gbKK_ECKarte = True Then
        Check1(16).value = vbChecked
    Else
        Check1(16).value = vbUnchecked
    End If
    
    If gbKK_Sonstige = True Then
        Check1(17).value = vbChecked
    Else
        Check1(17).value = vbUnchecked
    End If
    
    
    
    
    If gbKK_AliPay = True Then
        Check1(18).value = vbChecked
    Else
        Check1(18).value = vbUnchecked
    End If
    
    If gbKK_ApplePay = True Then
        Check1(19).value = vbChecked
    Else
        Check1(19).value = vbUnchecked
    End If
    
    If gbKK_GooglePay = True Then
        Check1(20).value = vbChecked
    Else
        Check1(20).value = vbUnchecked
    End If
    
    If gbKK_PayPal = True Then
        Check1(21).value = vbChecked
    Else
        Check1(21).value = vbUnchecked
    End If
    
    If gbKK_YabandPay = True Then
        Check1(22).value = vbChecked
    Else
        Check1(22).value = vbUnchecked
    End If
    
    
    
    
    
    
    
    
    
    
    
       
      
    
    If gbRESTinBAR = True Then
        Check34.value = vbChecked
    Else
        Check34.value = vbUnchecked
    End If
    
    If gbKASSMBEST = True Then
        Check109(1).value = vbChecked
    Else
        Check109(1).value = vbUnchecked
    End If
    
    If gbMitStaffelPreis = True Then
        Check37.value = vbChecked
    Else
        Check37.value = vbUnchecked
    End If
    
    If gbPBARGeld = True Then
        Check27.value = vbChecked
    Else
        Check27.value = vbUnchecked
    End If
    
    If gbKundRabattDeaktiv Then
        Check14.value = vbChecked
    Else
        Check14.value = vbUnchecked
    End If
    
    If gbJBTART Then
        Check40.value = vbChecked
    Else
        Check40.value = vbUnchecked
    End If
    
    If gbKUBONUS = True Then
        Check6.value = vbUnchecked
    Else
        Check6.value = vbChecked
    End If
    
    If gbGuStattBar = True Then
        Check26.value = vbChecked
    Else
        Check26.value = vbUnchecked
    End If
    
    If gbCoupon = True Then
        Check25.value = vbChecked
    Else
        Check25.value = vbUnchecked
    End If
    
    If gbNoKUBONUS_wenn_Art_and_Ges_rab = True Then
        Check19.value = vbChecked
    Else
        Check19.value = vbUnchecked
    End If
    
    If gbKUBONUS_WENN = True Then
        Check18.value = vbUnchecked
    Else
        Check18.value = vbChecked
    End If
    Text1(5).Text = gsiKUBONUS_SCHWELLE
    
    If gbKUWAHLMAIL = True Then
        Check5.value = vbChecked
    Else
        Check5.value = vbUnchecked
    End If
    
    If gbKUWAHLROT = True Then
        Check4.value = vbChecked
    Else
        Check4.value = vbUnchecked
    End If
    
    If gbKUWAHLGESPERRTROT = True Then
        Check31.value = vbChecked
    Else
        Check31.value = vbUnchecked
    End If
    
    If gbKUWAHLfbimmer = True Then
        Check30.value = vbChecked
    Else
        Check30.value = vbUnchecked
    End If
    
    If gbAUSBLDU = True Then
        Check55.value = vbChecked
    Else
        Check55.value = vbUnchecked
    End If
    
    If gbAUSBLSH = True Then
        Check56.value = vbChecked
    Else
        Check56.value = vbUnchecked
    End If
    
    If gbAUSBLLS = True Then
        Check57.value = vbChecked
    Else
        Check57.value = vbUnchecked
    End If
    
    '*----Neu 19.10.2010
    
    If gbBarAnz = True Then
        Check20.value = vbChecked
    Else
        Check20.value = vbUnchecked
    End If
    
    If gbEinfacheZollErstattung = True Then
        Check22.value = vbChecked
    Else
        Check22.value = vbUnchecked
    End If
    
    If gbUmsAnz = True Then
        Check62.value = vbChecked
    Else
        Check62.value = vbUnchecked
    End If
    
    If gbEDITKASSNR = True Then
        Check11.value = vbChecked
    Else
        Check11.value = vbUnchecked
    End If
    
    If gbArtikelTextSuche = True Then
        Check12.value = vbChecked
    Else
        Check12.value = vbUnchecked
    End If
    
    
    
    If gbGEBRABK = True Then
        Check60.value = vbChecked
    Else
        Check60.value = vbUnchecked
    End If
    
    If gbKUDU = True Then
        Check59.value = vbChecked
    Else
        Check59.value = vbUnchecked
    End If
    
    If gbZweitMoni = True Then
        Check15.value = vbChecked
        Check47.Visible = True
    Else
        Check15.value = vbUnchecked
        Check47.Visible = False
        Check47.value = vbUnchecked
    End If
    
    If gbZweitMoniMinimieren = True Then
        Check47.value = vbChecked
    Else
        Check47.value = vbUnchecked
    End If
    
    If giBARGELDART = 0 Then
        Option2(0).value = True
    ElseIf giBARGELDART = 1 Then
        Option2(1).value = True
    ElseIf giBARGELDART = 2 Then
        Option2(2).value = True
    End If
    
    
    If gbOpenSchubRetoure = True Then '55
        Check2.value = vbChecked
    Else
        Check2.value = vbUnchecked
    End If
    
    If gbBONNEIN = True Then
        Check69.value = vbChecked
    Else
        Check69.value = vbUnchecked
    End If
    
    If gbBONWAHL = True Then
        Check70.value = vbChecked
    Else
        Check70.value = vbUnchecked
    End If
    
    If gbMitPreis = True Then
        Check10.value = vbChecked
    Else
        Check10.value = vbUnchecked
    End If
    
    If gbkassgefuehrt = True Then
        Check71.value = vbChecked
    Else
        Check71.value = vbUnchecked
    End If
    
    If gbKKSCHUB = True Then
        Check29.value = vbChecked
    Else
        Check29.value = vbUnchecked
    End If
    
    If gbKOLSCHUB = True Then
        Check7.value = vbChecked
    Else
        Check7.value = vbUnchecked
    End If
    
    If gbBARZSCHUB = True Then
        Check8.value = vbChecked
    Else
        Check8.value = vbUnchecked
    End If
    
    If gbKBSCHUB = True Then
        Check9.value = vbChecked
    Else
        Check9.value = vbUnchecked
    End If
    
    
    If gbGeld = True Then
        Check21.value = vbChecked
    Else
        Check21.value = vbUnchecked
    End If
    
    If gbHandelsspanne_Ausblenden = True Then
        Check17.value = vbChecked
    Else
        Check17.value = vbUnchecked
    End If
    
    If gbAlterGutschein_Ausblenden = True Then
        Check24.value = vbChecked
    Else
        Check24.value = vbUnchecked
    End If
    
    Text1(3).Text = gsiGESRAB
    Text1(4).Text = gsGESRABBEZ
    
    Text1(31).Text = gsGZBez
    
    Text3.Text = gdVerBGesrabatt
    
    Text4.Text = gdCheckPreis
    
    If gbRabVs = True Then
        Check63.value = vbChecked
    Else
        Check63.value = vbUnchecked
    End If
    
    If gbBARDINA4 = True Then
        Check53.value = vbChecked
    Else
        Check53.value = vbUnchecked
    End If
    
    If gbDINA4VIS = True Then
        Check16.value = vbChecked
    Else
        Check16.value = vbUnchecked
    End If
    
    
    
    If gbDINA4RECHFU = True Then
        Check3.value = vbChecked
    Else
        Check3.value = vbUnchecked
    End If
    
    If gbKeineBestVerWarengru = True Then
        Check48.value = vbChecked
    Else
        Check48.value = vbUnchecked
    End If
    
    If gbOhnebestProt = True Then
        Check61.value = vbChecked
    Else
        Check61.value = vbUnchecked
    End If
    
    If gbBestDateien = True Then
        Check23.value = vbChecked
        Text1(10).Text = gsPfadBestandlive
        Text1(10).Visible = True
        lbl6(0).Visible = True
    Else
        Check23.value = vbUnchecked
        Text1(10).Text = ""
        Text1(10).Visible = False
        lbl6(0).Visible = False
    End If
    
    If gbSTORNOcheck2Bed = True Then
        Check86.value = vbChecked
    Else
        Check86.value = vbUnchecked
    End If
    
    'Warengruppen für Aboplus auslesen
    
    If NewTableSuchenDBKombi("ABOPLUS", gdBase) = True Then
        For i = 0 To 19
            glWGTaste(i) = 0
        Next i
        
        Set rsrs = gdBase.OpenRecordset("ABOPLUS")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            i = 0
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    glWGTaste(i) = rsrs!artnr
                    i = i + 1
                End If
                
                If Not IsNull(rsrs!PFNR) Then
                    Text2.Text = rsrs!PFNR
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close
        
        List1.Clear
        For i = 0 To 19
            If glWGTaste(i) > 0 Then
                List1.AddItem glWGTaste(i)
            End If
        Next i
    End If
    
    'Warengruppen für Bonus Auto Abzug auslesen
    If NewTableSuchenDBKombi("BONUSAA", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("BONUSAA")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!artnr) Then
                Text1(1).Text = rsrs!artnr
            End If
        End If
        rsrs.Close
    End If
    
    'Warengruppen für Bonus NR Eingabe auslesen
    If NewTableSuchenDBKombi("BONUSNRE", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("BONUSNRE")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!artnr) Then
                Text1(2).Text = rsrs!artnr
            End If
        End If
        rsrs.Close
    End If
    
    If gbNachKBbeiEC = True Then
        Check109(3).value = vbChecked
    Else
        Check109(3).value = vbUnchecked
    End If
    
    If gbKBmBI = True Then
        Check51.value = vbChecked
    Else
        Check51.value = vbUnchecked
    End If
    
    If gbmGDetails = True Then
        Check65.value = vbChecked
    Else
        Check65.value = vbUnchecked
    End If
    
    If gbArtrabhalten = True Then
        Check75.value = vbChecked
    Else
        Check75.value = vbUnchecked
    End If
    
    If gbBEDLEER = True Then
        Check66.value = vbChecked
    Else
        Check66.value = vbUnchecked
    End If

    
    
    
    
    If gbLeiste2Start = True Then
        Check89.value = vbChecked
    Else
        Check89.value = vbUnchecked
    End If
    
    If gbNOBONDRUCKER = True Then
        Check78.value = vbChecked
    Else
        Check78.value = vbUnchecked
    End If
            
    If gbIdentUser = True Then
        Check105.value = vbChecked
    Else
        Check105.value = vbUnchecked
    End If
    
    If gbPLZGEBIET = True Then
        Check13.value = vbChecked
    Else
        Check13.value = vbUnchecked
    End If
    
    If gbMitKundeWahlHinweis = True Then
        Check39.value = vbChecked
    Else
        Check39.value = vbUnchecked
    End If
    
    If gbPLZGEBIET_AuchBeiKUWAHL = True Then
        Check35.value = vbChecked
    Else
        Check35.value = vbUnchecked
    End If
    
    Text1(6).Text = glRRArtnr
    Text1(7).Text = glGSArtnr
    Text1(8).Text = gdZeitungsSpanne
    Text1(9).Text = glPrimLinr
    Text1(21).Text = glZeitungsLinr
    
    Text1(12).Text = glPaketLinr
    Text1(14).Text = glSpezFotoartikel
    Text1(22).Text = glSpezLottoauszahlartikel
    Text1(24).Text = glBonusGrenzeArtnr
    Text1(28).Text = glBonusAuszahlungArtnr
    Text1(26).Text = glECAuszahlArtnr
    
    gsSpezArtikel = SwapStr(gsSpezArtikel, "$", ",")
    Text1(13).Text = gsSpezArtikel
    
    gsRabattAusnahmeArtikel = SwapStr(gsRabattAusnahmeArtikel, "$", vbCrLf)
    Text1(17).Text = gsRabattAusnahmeArtikel
    
    Text1(32).Text = glBaganzArtnr
    Text1(33).Text = glBaganzAR
    
    Text1(35).Text = glLieblingArtnr
    Text1(34).Text = glLieblingAR
    
    Text1(36).Text = gdWarenkorbWert
    Text1(37).Text = gdWarenkorbGR
    
    Text1(38).Text = glZehnProzLinr
    Text1(39).Text = glZehnProzArtnr
    
    Text8.Text = gdKartenschwellenwert
    
    If gbMBBLOCKFrage = True Then
        Check107.value = vbChecked
        Text1(11).Text = gsSperrFrage
        Text1(11).Visible = True
    Else
        Check107.value = vbUnchecked
        Text1(11).Text = ""
        Text1(11).Visible = False
    End If
    
    Text1(16).Text = gsKassPass
    
    Text1(27).Text = gskPW
    
    Label4(32).Tag = gsJUGENDSCHUTZFARBE
    Dim ctmp As String
    ctmp = gsJUGENDSCHUTZFARBE
    
    With Label4(32)
        If ctmp = "1" Then
            .BackColor = glfarbe(1)
            .ForeColor = vbBlack
        ElseIf ctmp = "2" Then
            .BackColor = glfarbe(2)
            .ForeColor = vbBlack
        ElseIf ctmp = "3" Then
            .BackColor = glfarbe(3)
            .ForeColor = vbBlack
        ElseIf ctmp = "4" Then
            .BackColor = glfarbe(4)
            .ForeColor = vbBlack
        ElseIf ctmp = "5" Then
            .BackColor = glfarbe(5)
            .ForeColor = vbBlack
        ElseIf ctmp = "6" Then
            .BackColor = glfarbe(6)
            .ForeColor = vbBlack
        ElseIf ctmp = "7" Then
            .BackColor = glfarbe(7)
            .ForeColor = vbBlack
        ElseIf ctmp = "8" Then
            .BackColor = glfarbe(8)
            .ForeColor = vbBlack
        ElseIf ctmp = "9" Then
            .BackColor = glfarbe(9)
            .ForeColor = vbBlack
        ElseIf ctmp = "11" Then
            .BackColor = glfarbe2(1)
            .ForeColor = vbBlack
        ElseIf ctmp = "12" Then
            .BackColor = glfarbe2(2)
            .ForeColor = vbBlack
        ElseIf ctmp = "13" Then
            .BackColor = glfarbe2(3)
            .ForeColor = vbBlack
        ElseIf ctmp = "14" Then
            .BackColor = glfarbe2(4)
            .ForeColor = vbBlack
        ElseIf ctmp = "15" Then
            .BackColor = glfarbe2(5)
            .ForeColor = vbBlack
        ElseIf ctmp = "16" Then
            .BackColor = glfarbe2(6)
            .ForeColor = vbBlack
        ElseIf ctmp = "17" Then
            .BackColor = glfarbe2(7)
            .ForeColor = vbBlack
        ElseIf ctmp = "18" Then
            .BackColor = glfarbe2(8)
            .ForeColor = vbBlack
        ElseIf ctmp = "19" Then
            .BackColor = glfarbe2(9)
            .ForeColor = vbBlack
        End If
    End With
            
            
    If gsJUGENDSCHUTZFARBE <> "" Then
        Label4(32).Caption = "Farbauswahl"
    Else
        Label4(32).Caption = "bitte Farbe wählen"
    End If
    
    If gbSondRab Then
        Check28.value = vbChecked
    Else
        Check28.value = vbUnchecked
    End If
    
    If gbGutsch = True Then
        opt1(4).value = True
        opt1(5).value = False
    Else
        opt1(4).value = False
        opt1(5).value = True
    End If
    
    If gbOGV = True Then
        Check42.value = vbChecked
    Else
        Check42.value = vbUnchecked
    End If
    
    If gbGutscheinBeiVKversteuern = True Then
    
        Check45.Visible = False

        lbl6(44).Caption = "Gutscheine beim Verkauf versteuern = aktiviert"
        
        Dim dateStichtag As Date
        dateStichtag = ermStichtag
        
        lbl6(44).Caption = lbl6(44).Caption & ", Stichtag: " & dateStichtag
    Else
        Check45.Visible = True
        Check45.value = vbUnchecked
        lbl6(44).Caption = "Gutscheine beim Verkauf versteuern"
    End If
    
    
    If gbGutschnrKomplett = True Then
        Check43.value = vbChecked
    Else
        Check43.value = vbUnchecked
    End If
    
    If gbRGO = True Then
        Check46.value = vbChecked
    Else
        Check46.value = vbUnchecked
    End If
    
    
    
    If gsUnbekanntStrichMail <> "" Then
        Text1(15).Text = gsUnbekanntStrichMail
    Else
        Text1(15).Text = ""
    End If
    
    If gsNachtVerarbeitungMail <> "" Then
        Text1(29).Text = gsNachtVerarbeitungMail
    Else
        Text1(29).Text = ""
    End If
    
    
    
    Text6.Text = gdRESTGU
    
    Text1(18).Text = gdSCHWELLEWK
    
    If gsFARBKASSE = "1" Then
        Option1(10).value = True
    ElseIf gsFARBKASSE = "2" Then
        Option1(7).value = True
    End If
    
    
    If gsECBILD = "1" Then
        Option1(6).value = True
    ElseIf gsECBILD = "2" Then
        Option1(11).value = True
    
    End If
    
    If gsAbrunden = "1" Then
        Option1(8).value = True
    ElseIf gsAbrunden = "2" Then
        Option1(9).value = True
    ElseIf gsAbrunden = "3" Then
        Option1(5).value = True
    End If
    
    If gbNurBonusfRunden = True Then
        Check32.value = vbChecked
    Else
        Check32.value = vbUnchecked
    End If
    
    Text1(19).Text = glAutoKundnrforKundBest
    
    Text1(25).Text = glAutoAusSchFiliale
    
    
    
    
    LeseBonusGrenze
    
    Text1(23).Text = gdBonusGutscheinBeiGrenze
    
    If gbFILBONI = True Then
        Check38.value = vbChecked
    Else
        Check38.value = vbUnchecked
    End If
    
    If gbBonusBNB = True Then
        Check41.value = vbChecked
    Else
        Check41.value = vbUnchecked
    End If
    
    If gbNachfragenbeiWGNohnePreis = True Then
        Check33.value = vbChecked
    Else
        Check33.value = vbUnchecked
    End If
    
    
    If gbParknetto = True Then
        Check36.value = vbChecked
    Else
        Check36.value = vbUnchecked
    End If
    
    If gbArtsucheArtFarb = True Then
        Check108.value = vbChecked
    Else
        Check108.value = vbUnchecked
    End If
    
    If gbTerminReminderSMS = True Then
        Check44.value = vbChecked
        
        If gsTerminReminderstart <> "" Then
            DTPickerSMS.value = gsTerminReminderstart
        Else
            DTPickerSMS.value = "10:00:00"
        End If
        
        Text1(30).Text = glTageVorTermin
    Else
        Check44.value = vbUnchecked
        Text1(30).Text = 2
    End If


Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub LeseBonusGrenze()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim cSQL As String
    Dim dWert As Double
    Dim iFileNr As Integer
    
    cSQL = "Select * from DBEINSTE"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BONUSGRENZ) Then
            dWert = rsrs!BONUSGRENZ
        Else
            dWert = 0
        End If
    Else
        dWert = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    gdBonusGrenze = dWert
    Text1(20).Text = gdBonusGrenze
    
    
    
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseBonusGrenze"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Positionieren131()
On Error GoTo LOKAL_ERROR
    
    fraWeiche.Top = 600
    fraWeiche.Left = 120
    fraWeiche.Width = 11655
    fraWeiche.Height = 7215
    
    Frame5.Top = 600
    Frame5.Left = 120
    Frame5.Width = 11655
    Frame5.Height = 7215
    
    fraAllg.Top = 600
    fraAllg.Left = 120
    fraAllg.Width = 11655
    fraAllg.Height = 7215
    
    fraTasten.Top = 600
    fraTasten.Left = 120
    fraTasten.Width = 11655
    fraTasten.Height = 7215
    
    fraSpez.Top = 600
    fraSpez.Left = 120
    fraSpez.Width = 11655
    fraSpez.Height = 7215
    
    fraKassenabschluss.Top = 600
    fraKassenabschluss.Left = 120
    fraKassenabschluss.Width = 11655
    fraKassenabschluss.Height = 7215
    
    fraAuszahlungsgrund.Top = 600
    fraAuszahlungsgrund.Left = 120
    fraAuszahlungsgrund.Width = 11655
    fraAuszahlungsgrund.Height = 7215
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren131"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherBonusgrenze()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim dWert As Double
    
    dWert = 0
    
    If Text1(20).Text <> "" Then
        If IsNumeric(Text1(20).Text) Then
            dWert = Format(Text1(20).Text, "######0.00")
        End If
    End If
    
    sSQL = "Update DBEINSTE Set BONUSGRENZ = '" & dWert & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    gdBonusGrenze = dWert
    
    dWert = 0
    
    If Text1(23).Text <> "" Then
        If IsNumeric(Text1(23).Text) Then
            dWert = Format(Text1(23).Text, "######0.00")
        End If
    End If
    
    sSQL = "Update DBEINSTE Set BonusGutscheinBeiGrenze = '" & dWert & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    gdBonusGutscheinBeiGrenze = dWert
    
    
    
    
    If Check38.value = vbChecked Then
        sSQL = "Update DBEINSTE Set FILBONI = true "
        gdBase.Execute sSQL, dbFailOnError
        gbFILBONI = True
        
    ElseIf Check38.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set FILBONI = false "
        gdBase.Execute sSQL, dbFailOnError

        gbFILBONI = False
    End If
    
    If Check41.value = vbChecked Then
        sSQL = "Update DBEINSTE Set BonusBNB = true "
        gdBase.Execute sSQL, dbFailOnError
        gbBonusBNB = True
        
    ElseIf Check41.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set BonusBNB = false "
        gdBase.Execute sSQL, dbFailOnError

        gbBonusBNB = False
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBonusgrenze"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherGeld()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    If Check21.value = vbChecked Then
        sSQL = "Update WKEINSTE Set Geld = true"
        gdApp.Execute sSQL, dbFailOnError
        gbGeld = True
    ElseIf Check21.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set Geld = False"
        gdApp.Execute sSQL, dbFailOnError
        gbGeld = False
    End If
    
    If Check17.value = vbChecked Then
        sSQL = "Update KASSEIN Set Handelsspanne_Ausblenden = True"
        gdBase.Execute sSQL, dbFailOnError
        gbHandelsspanne_Ausblenden = True
    ElseIf Check17.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set Handelsspanne_Ausblenden = False"
        gdBase.Execute sSQL, dbFailOnError
        gbHandelsspanne_Ausblenden = False
    End If
    
    If Check24.value = vbChecked Then
        sSQL = "Update KASSEIN Set AlterGutschein_Ausblenden = True"
        gdBase.Execute sSQL, dbFailOnError
        gbAlterGutschein_Ausblenden = True
    ElseIf Check24.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set AlterGutschein_Ausblenden = False"
        gdBase.Execute sSQL, dbFailOnError
        gbAlterGutschein_Ausblenden = False
    End If
    
    If Check109(1).value = vbChecked Then
        sSQL = "Update KASSEIN Set KASSMBEST = True"
        gdBase.Execute sSQL, dbFailOnError
        gbKASSMBEST = True
    ElseIf Check109(1).value = vbUnchecked Then
        sSQL = "Update KASSEIN Set KASSMBEST = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKASSMBEST = False
    End If
    
    If Check34.value = vbChecked Then
        sSQL = "Update KASSEIN Set RESTinBAR = True"
        gdBase.Execute sSQL, dbFailOnError
        gbRESTinBAR = True
    ElseIf Check34.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set RESTinBAR = False"
        gdBase.Execute sSQL, dbFailOnError
        gbRESTinBAR = False
    End If
    
    gdVerBGesrabatt = 0
    
    If Text3.Text <> "" Then
        If IsNumeric(Text3.Text) Then
            gdVerBGesrabatt = CDbl(Text3.Text)
        End If
    End If
    
    If gdVerBGesrabatt > 0 Then
        sSQL = "Update WKEINSTE Set VerBGesrabatt = '" & gdVerBGesrabatt & "'"
        gdApp.Execute sSQL, dbFailOnError
       
    Else
        sSQL = "Update WKEINSTE Set VerBGesrabatt = 0 "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    gdCheckPreis = 0
    
    If Text4.Text <> "" Then
        If IsNumeric(Text4.Text) Then
            gdCheckPreis = CDbl(Text4.Text)
        End If
    End If
    
    If gdCheckPreis > 0 Then
        sSQL = "Update WKEINSTE Set CHECKPREIS = '" & gdCheckPreis & "'"
        gdApp.Execute sSQL, dbFailOnError
       
    Else
        sSQL = "Update WKEINSTE Set CHECKPREIS = 0 "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    gdKartenschwellenwert = 0
    
    If Text8.Text <> "" Then
        If IsNumeric(Text8.Text) Then
            gdKartenschwellenwert = CDbl(Text8.Text)
        End If
    End If
    
    If gdKartenschwellenwert > 0 Then
        sSQL = "Update WKEINSTE Set Kartenschwellenwert = '" & gdKartenschwellenwert & "'"
        gdApp.Execute sSQL, dbFailOnError
       
    Else
        sSQL = "Update WKEINSTE Set Kartenschwellenwert = 0 "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    gsPfadBestandlive = ""
    
    If Text1(10).Text <> "" Then
        gsPfadBestandlive = Text1(10).Text
    End If
    
    If gsPfadBestandlive <> "" Then
        sSQL = "Update WKEINSTE Set PfadBestandlive = '" & gsPfadBestandlive & "'"
        gdApp.Execute sSQL, dbFailOnError
       
    Else
        sSQL = "Update WKEINSTE Set PfadBestandlive = '' "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherGeld"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherParknetto()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "ParkNetto", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "ParkNetto", "BIT", gdApp
    End If
    
    If Check36.value = vbChecked Then
        sSQL = "Update WKEINSTE Set PARKNetto = true"
        gdApp.Execute sSQL, dbFailOnError
        gbParknetto = True
    ElseIf Check36.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set PARKNetto = False"
        gdApp.Execute sSQL, dbFailOnError
        gbParknetto = False
    End If

    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherParknetto"
        Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub speicherJugendschutz()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim ctmp As String
    
    gsJUGENDSCHUTZFARBE = ""
    
    ctmp = Trim$(Label4(32).Tag)
    If ctmp <> "" Then
    
        gsJUGENDSCHUTZFARBE = ctmp
        
    End If
    
    
    sSQL = "Update DBEINSTE Set JUGENDSCHUTZFARBE = '" & gsJUGENDSCHUTZFARBE & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    If Text1(15).Text <> "" Then
        If InStr(1, Text1(15).Text, "@") = 0 Or InStr(1, Text1(15).Text, ".") = 0 Then
            anzeige "rot2", "Bitte geben Sie die Emailadresse richtig ein!", Label1(4)
            
            fraTasten.Visible = False
            fraWeiche.Visible = False
            fraAllg.Visible = True
            fraSpez.Visible = False
            
            Exit Sub

        End If
        
        
        
        
        If ermFirmenMail = "" Then
        
            MsgBox "Bitte auch in den Unternehmensdaten eine Emaildadresse als Absendermailadresse hinterlegen (Service/Einstellungen/Unternehmens-Daten)", vbInformation, "Winkiss Hinweis:"
            
        End If
        
        
    End If
    
    
    gsUnbekanntStrichMail = Text1(15).Text
    
    sSQL = "Update DBEINSTE Set UnbekanntStrichMail = '" & gsUnbekanntStrichMail & "' "
    gdBase.Execute sSQL, dbFailOnError

    
    
    
    
    If Text1(29).Text <> "" Then
        If InStr(1, Text1(29).Text, "@") = 0 Or InStr(1, Text1(29).Text, ".") = 0 Then
            anzeige "rot2", "Bitte geben Sie die Emailadresse richtig ein!", Label1(4)
            
            fraTasten.Visible = False
            fraWeiche.Visible = False
            fraAllg.Visible = True
            fraSpez.Visible = False
            
            Exit Sub

        End If
        
        
        
        
        If ermFirmenMail = "" Then
        
            MsgBox "Bitte auch in den Unternehmensdaten eine Emaildadresse als Absendermailadresse hinterlegen (Service/Einstellungen/Unternehmens-Daten)", vbInformation, "Winkiss Hinweis:"
            
        End If
        
        
    End If
    
    
    gsNachtVerarbeitungMail = Text1(29).Text
    
    sSQL = "Update DBEINSTE Set NachtVerarbeitungMail = '" & gsNachtVerarbeitungMail & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    
   
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherJugendschutz"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub speicherJubi()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check28.value = vbChecked Then
        sSQL = "Update DBEINSTE Set SondRab = True "
        gdBase.Execute sSQL, dbFailOnError
        gbSondRab = True
    Else
        sSQL = "Update DBEINSTE Set SondRab = False "
        gdBase.Execute sSQL, dbFailOnError
        gbSondRab = False
        
    End If
    
    sSQL = "Update WKEINSTE Set KassPass = '" & Text1(16).Text & "' "
    gdApp.Execute sSQL, dbFailOnError

    gsKassPass = Text1(16).Text
    
    
    sSQL = "Update KASSEIN Set kPW = '" & Text1(27).Text & "' "
    gdBase.Execute sSQL, dbFailOnError

    gskPW = Text1(27).Text
    
    
    
    
    gsiGESRAB = 0
    
    If Text1(3).Text <> "" Then
        If IsNumeric(Text1(3).Text) Then
            gsiGESRAB = CSng(Text1(3).Text)
        End If
    End If
    
    If gsiGESRAB > 0 Then
        sSQL = "Update DBEINSTE Set GESRAB = '" & gsiGESRAB & "'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update DBEINSTE Set GESRAB = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    gsGESRABBEZ = "Jubiläumsrabatt:"
    
    If Text1(4).Text <> "" Then
        gsGESRABBEZ = Text1(4).Text
    End If
    
    If gsGESRABBEZ <> "" Then
        sSQL = "Update DBEINSTE Set GESRABBEZ = '" & gsGESRABBEZ & "'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update DBEINSTE Set GESRABBEZ = 'Jubiläumsrabatt' "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    gsGZBez = ""
    
    If Text1(31).Text <> "" Then
        gsGZBez = Text1(31).Text
    End If
    
    If gsGZBez <> "" Then
        sSQL = "Update DBEINSTE Set GZBEZ = '" & gsGZBez & "'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update DBEINSTE Set GZBez = '' "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    
    If Check14.value = vbChecked Then
        sSQL = "Update DBEINSTE Set KundRabattDeaktiv = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKundRabattDeaktiv = True
    ElseIf Check14.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set KundRabattDeaktiv = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKundRabattDeaktiv = False
    End If
    
    If Check40.value = vbChecked Then
        sSQL = "Update DBEINSTE Set JBTART = true"
        gdBase.Execute sSQL, dbFailOnError
        gbJBTART = True
    ElseIf Check40.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set JBTART = False"
        gdBase.Execute sSQL, dbFailOnError
        gbJBTART = False
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherJubi"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub speicherKKSCHUB()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check27.value = vbChecked Then
        sSQL = "Update WKEINSTE Set PBARGeld = true"
        gdApp.Execute sSQL, dbFailOnError
        gbPBARGeld = True
        
    ElseIf Check27.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set PBARGeld = False"
        gdApp.Execute sSQL, dbFailOnError
        gbPBARGeld = False
        
    End If
    
    If Check29.value = vbChecked Then
        sSQL = "Update WKEINSTE Set KKSCHUB = true"
        gdApp.Execute sSQL, dbFailOnError
        gbKKSCHUB = True
        
    ElseIf Check29.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set KKSCHUB = False"
        gdApp.Execute sSQL, dbFailOnError
        gbKKSCHUB = False
        
    End If
    
    If Check7.value = vbChecked Then
        sSQL = "Update WKEINSTE Set KOLSCHUB = true"
        gdApp.Execute sSQL, dbFailOnError
        gbKOLSCHUB = True
        
    ElseIf Check7.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set KOLSCHUB = False"
        gdApp.Execute sSQL, dbFailOnError
        gbKOLSCHUB = False
        
    End If
    
    If Check8.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BARZSCHUB = true"
        gdApp.Execute sSQL, dbFailOnError
        gbBARZSCHUB = True
    ElseIf Check8.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BARZSCHUB = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBARZSCHUB = False
    End If
    
    If Check9.value = vbChecked Then
        sSQL = "Update WKEINSTE Set KBSCHUB = true"
        gdApp.Execute sSQL, dbFailOnError
        gbKBSCHUB = True
    ElseIf Check9.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set KBSCHUB = False"
        gdApp.Execute sSQL, dbFailOnError
        gbKBSCHUB = False
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherKKSCHUB"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherKKaktiv()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    'Visa
    If Check1(12).value = vbChecked Then
        sSQL = "Update kassein Set KK_Visa = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_Visa = True
    ElseIf Check1(12).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_Visa = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_Visa = False
    End If
    
    'Eurocard/Mastercard
    If Check1(13).value = vbChecked Then
        sSQL = "Update kassein Set KK_EurocardMastercard = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_EurocardMastercard = True
    ElseIf Check1(13).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_EurocardMastercard = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_EurocardMastercard = False
    End If
    
    'American Express
    If Check1(14).value = vbChecked Then
        sSQL = "Update kassein Set KK_AmericanExpress = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_AmericanExpress = True
    ElseIf Check1(14).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_AmericanExpress = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_AmericanExpress = False
    End If
    
    'Diners Club
    If Check1(15).value = vbChecked Then
        sSQL = "Update kassein Set KK_DinersClub = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_DinersClub = True
    ElseIf Check1(15).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_DinersClub = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_DinersClub = False
    End If
    
    'EC Karte
    If Check1(16).value = vbChecked Then
        sSQL = "Update kassein Set KK_ECKarte = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_ECKarte = True
    ElseIf Check1(16).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_ECKarte = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_ECKarte = False
    End If
    
    'Sonstige
    If Check1(17).value = vbChecked Then
        sSQL = "Update kassein Set KK_Sonstige = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_Sonstige = True
    ElseIf Check1(17).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_Sonstige = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_Sonstige = False
    End If
    
    'AliPay
    If Check1(18).value = vbChecked Then
        sSQL = "Update kassein Set KK_AliPay = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_AliPay = True
    ElseIf Check1(18).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_AliPay = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_AliPay = False
    End If
    
    'ApplePay
    If Check1(19).value = vbChecked Then
        sSQL = "Update kassein Set KK_ApplePay = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_ApplePay = True
    ElseIf Check1(19).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_ApplePay = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_ApplePay = False
    End If
    
    'GooglePay
    If Check1(20).value = vbChecked Then
        sSQL = "Update kassein Set KK_GooglePay = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_GooglePay = True
    ElseIf Check1(20).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_GooglePay = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_GooglePay = False
    End If
    
    'PayPal
    If Check1(21).value = vbChecked Then
        sSQL = "Update kassein Set KK_PayPal = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_PayPal = True
    ElseIf Check1(21).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_PayPal = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_PayPal = False
    End If
    
    'YabandPay
    If Check1(22).value = vbChecked Then
        sSQL = "Update kassein Set KK_YabandPay = true"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_YabandPay = True
    ElseIf Check1(22).value = vbUnchecked Then
        sSQL = "Update kassein Set KK_YabandPay = False"
        gdBase.Execute sSQL, dbFailOnError
        gbKK_YabandPay = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherKKaktiv"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check21_Click()
On Error GoTo LOKAL_ERROR
    
    If Check21.value = vbChecked Then
        If checkpic = False Then
            MsgBox "Es sind nicht alle Münzen/Scheine auf Ihrem Rechner vorhanden, wenden Sie sich an unsere Hotline(0511/955910)!", vbInformation, "Winkiss Hinweis:"
            Check21.value = vbUnchecked
        End If
    Else
    
    End If


    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check21_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub lesebutton()
On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    Set rsrs = gdBase.OpenRecordset("Button")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!indexnr) Then
                Check1(rsrs!indexnr).value = vbChecked
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesebutton"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check70_Click()
On Error GoTo LOKAL_ERROR
    
    If Check70.value = vbChecked Then
        Check69.Enabled = False
    Else
        Check69.Enabled = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check70_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo LOKAL_ERROR
'
'    Label1(0).ForeColor = glS1
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Form_MouseMove"
'    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
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

Private Sub Label4_dblClick(index As Integer)
On Error GoTo LOKAL_ERROR

If index = 32 Then
    Label4(index).Caption = "alle Farben"
    Label4(index).Tag = ""
    Label4(index).BackColor = lbl6(22).BackColor
    Label4(index).ForeColor = lbl6(22).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub opt1_Click(index As Integer)
On Error GoTo LOKAL_ERROR
    
    If opt1(4).value = True Then
        Check42.Visible = True
    ElseIf opt1(4).value = False Then
        Check42.Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "opt1_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Select Case index
        Case 4, 11, 31 'JubiBezeich
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%()"
        Case 3, 5, 8, 18, 20, 23, 33, 34, 36, 37 ' Prozente + Bonusgrenze
            cValid = "1234567890," & Chr$(8)
        Case 0, 1, 2, 6, 7, 9, 12, 14, 19, 21, 22, 24, 25, 26, 28, 30, 32, 33, 34, 35, 38, 39 ' Nr
            cValid = "1234567890" & Chr$(8)
        Case 10, 16, 27
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46)  '& - .
            cValid = cValid & "+äÄÜüÖöß/:\%()"
        Case 15, 29
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46)  '& - .
            cValid = cValid & "+äÄÜüÖöß/:\%().@"
        Case 13, 17 ' Nr
            cValid = "1234567890," & Chr$(8)
    End Select

    
    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(index As Integer)
On Error GoTo LOKAL_ERROR
    Text1(index).BackColor = glSelBack1
    Text1(index).SelStart = 0
    Text1(index).SelLength = Len(Text1(index).Text)
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus()
On Error GoTo LOKAL_ERROR
    Text3.BackColor = glSelBack1
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus()
On Error GoTo LOKAL_ERROR
    Text3.BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
On Error GoTo LOKAL_ERROR
    Text2.BackColor = glSelBack1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus()
On Error GoTo LOKAL_ERROR
    Text2.BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    cValid = gcNUM & Chr$(8)
    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_GotFocus(index As Integer)
On Error GoTo LOKAL_ERROR
    Text5(index).BackColor = glSelBack1
    Text5(index).SelStart = 0
    Text5(index).SelLength = Len(Text5(index).Text)
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Select Case index
        
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46)  '& - .
            cValid = cValid & "+äÄÜüÖöß/:\%()"
    End Select

    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text5_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR
    Text5(index).BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
