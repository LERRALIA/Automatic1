VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL94 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kosmetik"
   ClientHeight    =   8610
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   11910
   Icon            =   "frmWKL94.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   7800
      Width           =   2295
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   11
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   10
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   9
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   4200
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   7
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   4560
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   89
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Raucher:"
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   68
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kaffeegenuss:"
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   67
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sonnenbank:"
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   66
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Wasser:"
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   65
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sport:"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   64
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kundentyp"
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
         Index           =   13
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Typ Make-up:"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Duftrichtungen:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Lieblingsfarben:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Gesichtsform:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Teint:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Augen:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Haarfarbe:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   3000
      TabIndex        =   307
      Top             =   1800
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Chem. Peeling: Aufklärungsbogen und KD-Info"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   333
         Top             =   4200
         Width           =   5055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Ultraschall: Aufklärungsbogen"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   332
         Top             =   3960
         Width           =   5055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "MDA: Aufklärungsbogen und KD-Info"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   331
         Top             =   3720
         Width           =   5055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Micro-Needling: Aufklärungsbogen und KD-Info"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   330
         Top             =   3480
         Width           =   5175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Enthaarung: KD-Info, Anamnese-/Beratungsbogen"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   329
         Top             =   3000
         Width           =   5175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Einwilligung Minderjährige"
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   328
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "EV Hydrafacial"
         Height          =   255
         Index           =   28
         Left            =   5880
         TabIndex        =   327
         Top             =   3360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Aufklärungsbogen LashLifting"
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   326
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kundendatenblatt/Einverständniserklärung"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   325
         Top             =   1560
         Width           =   5055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kundenberatungsbogen Wimpern"
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   324
         Top             =   2280
         Width           =   4815
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "EV Lashes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   5760
         TabIndex        =   323
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "EV Säurepeeling"
         Height          =   255
         Index           =   23
         Left            =   5640
         TabIndex        =   322
         Top             =   3000
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "EV Micro Needling"
         Height          =   255
         Index           =   22
         Left            =   5880
         TabIndex        =   321
         Top             =   3840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "EV Permanent Make-up"
         Height          =   255
         Index           =   21
         Left            =   5520
         TabIndex        =   320
         Top             =   4320
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Pflegehinweis Wimpern"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   5640
         TabIndex        =   319
         Top             =   2160
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kundendaten/Anamnese/Einwilligungserklärung (Monteil)"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   315
         Top             =   2040
         Width           =   4935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "alle auswählen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   314
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Anamnesebogen  JetPeel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5640
         TabIndex        =   313
         Top             =   1920
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Mesoporation"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   312
         Top             =   1800
         Width           =   5055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "EV JetPeel"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   311
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":074C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   310
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Einwilligungen"
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
         Index           =   53
         Left            =   120
         TabIndex        =   309
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":0A56
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   308
         Top             =   5880
         Width           =   1935
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   120
      Top             =   7920
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox Text4 
         Height          =   405
         Index           =   6
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   278
         Top             =   5760
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Knickfuß"
         Height          =   255
         Index           =   14
         Left            =   5400
         TabIndex        =   252
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   16
         Left            =   7440
         TabIndex        =   251
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   15
         Left            =   8280
         TabIndex        =   250
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   245
         Top             =   5040
         Width           =   5055
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Index           =   1
         Left            =   5400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   243
         Top             =   4320
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         Caption         =   "nach Maß"
         Height          =   195
         Index           =   1
         Left            =   5880
         TabIndex        =   242
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Konfektion"
         Height          =   195
         Index           =   0
         Left            =   5880
         TabIndex        =   241
         Top             =   5280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Einlagen"
         Height          =   255
         Index           =   14
         Left            =   5400
         TabIndex        =   240
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   13
         Left            =   8280
         TabIndex        =   239
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   12
         Left            =   7440
         TabIndex        =   238
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   11
         Left            =   8280
         TabIndex        =   237
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   10
         Left            =   7440
         TabIndex        =   236
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   9
         Left            =   8280
         TabIndex        =   235
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   8
         Left            =   7440
         TabIndex        =   234
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   7
         Left            =   8280
         TabIndex        =   233
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   6
         Left            =   7440
         TabIndex        =   232
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   5
         Left            =   8280
         TabIndex        =   231
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   4
         Left            =   7440
         TabIndex        =   230
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   3
         Left            =   8280
         TabIndex        =   229
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   228
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   1
         Left            =   8280
         TabIndex        =   227
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Index           =   0
         Left            =   7440
         TabIndex        =   226
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   7560
         TabIndex        =   225
         Top             =   5400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Plattfuß"
         Height          =   255
         Index           =   13
         Left            =   5400
         TabIndex        =   220
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hohlfuß"
         Height          =   255
         Index           =   12
         Left            =   5400
         TabIndex        =   219
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Knickfuß"
         Height          =   255
         Index           =   11
         Left            =   5400
         TabIndex        =   218
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Spreizfuß"
         Height          =   255
         Index           =   10
         Left            =   5400
         TabIndex        =   217
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Senkfuß"
         Height          =   255
         Index           =   9
         Left            =   5400
         TabIndex        =   214
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Krampfadern"
         Height          =   255
         Index           =   8
         Left            =   5400
         TabIndex        =   213
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   79
         Left            =   4920
         TabIndex        =   212
         Top             =   4275
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   78
         Left            =   4680
         TabIndex        =   211
         Top             =   4200
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   77
         Left            =   4440
         TabIndex        =   210
         Top             =   4125
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   76
         Left            =   4200
         TabIndex        =   209
         Top             =   4035
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   75
         Left            =   3960
         TabIndex        =   208
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   74
         Left            =   3480
         TabIndex        =   207
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   73
         Left            =   3240
         TabIndex        =   206
         Top             =   4035
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   72
         Left            =   3000
         TabIndex        =   205
         Top             =   4125
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   71
         Left            =   2760
         TabIndex        =   204
         Top             =   4200
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   70
         Left            =   2520
         TabIndex        =   203
         Top             =   4275
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   69
         Left            =   4920
         TabIndex        =   202
         Top             =   3795
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   68
         Left            =   4680
         TabIndex        =   201
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   67
         Left            =   4440
         TabIndex        =   200
         Top             =   3645
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   66
         Left            =   4200
         TabIndex        =   199
         Top             =   3555
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   65
         Left            =   3960
         TabIndex        =   198
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   64
         Left            =   3480
         TabIndex        =   197
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   63
         Left            =   3240
         TabIndex        =   196
         Top             =   3555
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   62
         Left            =   3000
         TabIndex        =   195
         Top             =   3645
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   61
         Left            =   2760
         TabIndex        =   194
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   60
         Left            =   2520
         TabIndex        =   193
         Top             =   3795
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   59
         Left            =   4920
         TabIndex        =   192
         Top             =   3315
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   58
         Left            =   4680
         TabIndex        =   191
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   57
         Left            =   4440
         TabIndex        =   190
         Top             =   3165
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   56
         Left            =   4200
         TabIndex        =   189
         Top             =   3075
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   55
         Left            =   3960
         TabIndex        =   188
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   54
         Left            =   3480
         TabIndex        =   187
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   53
         Left            =   3240
         TabIndex        =   186
         Top             =   3075
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   52
         Left            =   3000
         TabIndex        =   185
         Top             =   3165
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   51
         Left            =   2760
         TabIndex        =   184
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   50
         Left            =   2520
         TabIndex        =   183
         Top             =   3315
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   49
         Left            =   4920
         TabIndex        =   182
         Top             =   2835
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   48
         Left            =   4680
         TabIndex        =   181
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   47
         Left            =   4440
         TabIndex        =   180
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   46
         Left            =   4200
         TabIndex        =   179
         Top             =   2595
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   45
         Left            =   3960
         TabIndex        =   178
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   44
         Left            =   3480
         TabIndex        =   177
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   43
         Left            =   3240
         TabIndex        =   176
         Top             =   2595
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   42
         Left            =   3000
         TabIndex        =   175
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   41
         Left            =   2760
         TabIndex        =   174
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   40
         Left            =   2520
         TabIndex        =   173
         Top             =   2835
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   39
         Left            =   4920
         TabIndex        =   172
         Top             =   2355
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   38
         Left            =   4680
         TabIndex        =   171
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   37
         Left            =   4440
         TabIndex        =   170
         Top             =   2205
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   36
         Left            =   4200
         TabIndex        =   169
         Top             =   2115
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   35
         Left            =   3960
         TabIndex        =   168
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   34
         Left            =   3480
         TabIndex        =   167
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   33
         Left            =   3240
         TabIndex        =   166
         Top             =   2115
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   32
         Left            =   3000
         TabIndex        =   165
         Top             =   2205
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   31
         Left            =   2760
         TabIndex        =   164
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   30
         Left            =   2520
         TabIndex        =   163
         Top             =   2355
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   29
         Left            =   4920
         TabIndex        =   162
         Top             =   1875
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   28
         Left            =   4680
         TabIndex        =   161
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   27
         Left            =   4440
         TabIndex        =   160
         Top             =   1725
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   26
         Left            =   4200
         TabIndex        =   159
         Top             =   1635
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   25
         Left            =   3960
         TabIndex        =   158
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   24
         Left            =   3480
         TabIndex        =   157
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   23
         Left            =   3240
         TabIndex        =   156
         Top             =   1635
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   22
         Left            =   3000
         TabIndex        =   155
         Top             =   1725
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   21
         Left            =   2760
         TabIndex        =   154
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   153
         Top             =   1875
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   19
         Left            =   4920
         TabIndex        =   152
         Top             =   1395
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   18
         Left            =   4680
         TabIndex        =   151
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   150
         Top             =   1245
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   16
         Left            =   4200
         TabIndex        =   149
         Top             =   1155
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   15
         Left            =   3960
         TabIndex        =   148
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   147
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   13
         Left            =   3240
         TabIndex        =   146
         Top             =   1155
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   145
         Top             =   1245
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   144
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   143
         Top             =   1395
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   142
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   141
         Top             =   880
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   140
         Top             =   800
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   139
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   138
         Top             =   640
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   137
         Top             =   640
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   136
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   135
         Top             =   800
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   134
         Top             =   880
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   133
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "eingewachsene Nägel"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   132
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Nagelpilz"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   131
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hammerzehen"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   130
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hautpilz"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   129
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hornhaut"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   124
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hühneraugen"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   123
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hallux Valgus"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   122
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Warzen"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   121
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "verwendete Produkte"
         Height          =   255
         Index           =   52
         Left            =   120
         TabIndex        =   279
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "nach außen"
         Height          =   255
         Index           =   46
         Left            =   5640
         TabIndex        =   254
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "nach innen"
         Height          =   255
         Index           =   45
         Left            =   5640
         TabIndex        =   253
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Zustand der Nägel:"
         Height          =   255
         Index           =   43
         Left            =   120
         TabIndex        =   244
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Schuhgröße:"
         Height          =   255
         Index           =   42
         Left            =   7560
         TabIndex        =   224
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "andere Fußinformationen:"
         Height          =   255
         Index           =   41
         Left            =   5400
         TabIndex        =   223
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Unterschenkel"
         Height          =   255
         Index           =   40
         Left            =   5640
         TabIndex        =   222
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Oberschenkel"
         Height          =   255
         Index           =   39
         Left            =   5640
         TabIndex        =   221
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF00&
         Caption         =   "links"
         Height          =   255
         Index           =   38
         Left            =   7200
         TabIndex        =   216
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF00&
         Caption         =   "rechts"
         Height          =   255
         Index           =   37
         Left            =   8040
         TabIndex        =   215
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF00&
         Caption         =   "rechts"
         Height          =   255
         Index           =   36
         Left            =   3960
         TabIndex        =   128
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF00&
         Caption         =   "links"
         Height          =   255
         Index           =   35
         Left            =   2520
         TabIndex        =   127
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":0D60
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   126
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Füße"
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
         Index           =   31
         Left            =   120
         TabIndex        =   125
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   104
      Top             =   0
      Width           =   11655
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   4
         Left            =   3840
         TabIndex        =   109
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   9000
         TabIndex        =   272
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "Beruf:"
         Height          =   255
         Index           =   49
         Left            =   8040
         TabIndex        =   271
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   119
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "Vorname:"
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   118
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "geb. am:"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   117
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "Tel. priv.:"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   116
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "Mobil:"
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   115
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         Caption         =   "Adresse:"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   114
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   113
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   6720
         TabIndex        =   112
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   111
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   6720
         TabIndex        =   108
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   107
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   6720
         TabIndex        =   106
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Kundendaten"
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
         Index           =   15
         Left            =   120
         TabIndex        =   105
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   110
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   94
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox Text4 
         Height          =   765
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   276
         Top             =   4440
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         Height          =   765
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   257
         Top             =   2760
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         Height          =   765
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   255
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "verwendete Produkte"
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   277
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "linke Hand"
         Height          =   255
         Index           =   11
         Left            =   7200
         TabIndex        =   270
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "rechte Hand"
         Height          =   255
         Index           =   10
         Left            =   7200
         TabIndex        =   269
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "L5"
         Height          =   255
         Index           =   9
         Left            =   6720
         TabIndex        =   268
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "L4"
         Height          =   255
         Index           =   8
         Left            =   6360
         TabIndex        =   267
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "L3"
         Height          =   255
         Index           =   7
         Left            =   6000
         TabIndex        =   266
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "L2"
         Height          =   255
         Index           =   6
         Left            =   5640
         TabIndex        =   265
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "L1"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   264
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "R5"
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   263
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "R4"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   262
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "R3"
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   261
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "R2"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   260
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "R1"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   259
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Tiptyp"
         Height          =   255
         Index           =   48
         Left            =   120
         TabIndex        =   258
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Nagelbesonderheiten"
         Height          =   255
         Index           =   47
         Left            =   120
         TabIndex        =   256
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Hände"
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
         Index           =   24
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":106A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   95
         Top             =   5880
         Width           =   1935
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1920
      TabIndex        =   86
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      FontTransparent =   0   'False
      Height          =   1455
      Left            =   120
      MouseIcon       =   "frmWKL94.frx":1374
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   85
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame6"
      Height          =   3255
      Left            =   0
      TabIndex        =   77
      Top             =   4680
      Width           =   2775
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Einwilligungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":167E
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   306
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Füße"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":1988
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   98
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hände"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":1C92
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   97
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gesundheit / Risiken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":1F9C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   82
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Kundentyp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":22A6
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   81
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enthaarung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":25B0
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   80
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "verwendete Produkte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":28BA
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   79
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hautbild"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":2BC4
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   78
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   46
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Bikini-Zone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   52
         Top             =   3240
         Width           =   4575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Beine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   51
         Top             =   2760
         Width           =   4575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Achseln"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   5175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Arme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   4935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Augenbrauen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   5295
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Gesicht"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":2ECE
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   88
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Enthaarung"
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
         Index           =   23
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9360
      TabIndex        =   45
      Top             =   6720
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   290
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   289
         Top             =   2145
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   288
         Top             =   2625
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   287
         Top             =   3585
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   286
         Top             =   4065
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   285
         Top             =   4545
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   284
         Top             =   5025
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   283
         Top             =   1185
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   282
         Top             =   1665
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   10
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   281
         Top             =   3105
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   12
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   280
         Top             =   5520
         Width           =   4335
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   1560
         TabIndex        =   273
         Text            =   "Combo4"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text3 
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   71
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ListBox List3 
         Height          =   1425
         Left            =   120
         TabIndex        =   70
         Top             =   3960
         Width           =   4095
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   120
         TabIndex        =   69
         Top             =   1200
         Width           =   4095
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   1
         Left            =   1560
         TabIndex        =   275
         ToolTipText     =   "Kalender"
         Top             =   2760
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
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   4560
         MouseIcon       =   "frmWKL94.frx":31D8
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   318
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
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
         Index           =   8
         Left            =   5760
         TabIndex        =   305
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
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
         Index           =   0
         Left            =   4320
         TabIndex        =   304
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "vorhandene Termine"
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
         Left            =   120
         TabIndex        =   303
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Mit Klick auf vorhandene Termine die verwendeten Produkte einsehen."
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
         Index           =   16
         Left            =   120
         TabIndex        =   302
         Top             =   5760
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Reinigung / Peeling:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   301
         Top             =   500
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Masken:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   300
         Top             =   1960
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Extras:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   299
         Top             =   2440
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Tagespflege / Make-up:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   298
         Top             =   3400
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Beratung für zu Hause:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   297
         Top             =   3860
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Proben:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   296
         Top             =   4360
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "gekaufte Produkte:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   295
         Top             =   4820
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Tonic:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4320
         TabIndex        =   294
         Top             =   1010
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Peeling:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   293
         Top             =   1480
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Massage:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   292
         Top             =   2910
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Nagellack:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   4320
         TabIndex        =   291
         Top             =   5300
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Gliederung:"
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   274
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "alle Löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         MouseIcon       =   "frmWKL94.frx":34E2
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   93
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "einzeln Löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2280
         MouseIcon       =   "frmWKL94.frx":37EC
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   92
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hinzufügen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2280
         MouseIcon       =   "frmWKL94.frx":3AF6
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   91
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":3E00
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   87
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Diese Behandlungen wurden durchgeführt"
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
         TabIndex        =   72
         Top             =   3720
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "verwendete Produkte"
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
         Index           =   25
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   2
      Top             =   5640
      Width           =   2415
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Komedonen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   4320
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Pickel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   4800
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Milien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   74
         Top             =   5280
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Couperose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   73
         Top             =   5760
         Width           =   3495
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
         Style           =   2  'Dropdown-Liste
         TabIndex        =   63
         Top             =   600
         Width           =   5535
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "sehr"
            Height          =   255
            Index           =   16
            Left            =   3000
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "ab und zu"
            Height          =   255
            Index           =   15
            Left            =   1560
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "unempfindlich"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "EMPFINDLICHKEIT"
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
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "erschlafft"
            Height          =   255
            Index           =   13
            Left            =   4320
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "nachlassend"
            Height          =   255
            Index           =   12
            Left            =   3000
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "normal"
            Height          =   255
            Index           =   11
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "straff"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "SPANNKRAFT"
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
            TabIndex        =   59
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "groß"
            Height          =   255
            Index           =   9
            Left            =   3000
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "normal"
            Height          =   255
            Index           =   8
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "klein"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "PORENGRÖSSE"
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
            TabIndex        =   58
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "normal"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "arm"
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "sehr arm"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "FEUCHTIGKEIT"
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
            TabIndex        =   57
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "T-Zone-Nase"
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "fett"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "normal"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "trocken"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "FETTGEHALT"
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
            TabIndex        =   56
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":410A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   317
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":4414
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   83
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Hautbild"
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
         Index           =   17
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
      Begin VB.CheckBox Check5 
         Caption         =   "HIV-positiv"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   249
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Bluter"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   248
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Diabetes"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   247
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   1095
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   102
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox Text1 
         Height          =   1095
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   100
         Top             =   2520
         Width           =   5415
      End
      Begin VB.TextBox Text1 
         Height          =   1095
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   1
         Top             =   4080
         Width           =   5415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":471E
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   316
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Risiken:"
         Height          =   255
         Index           =   44
         Left            =   5880
         TabIndex        =   246
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Allergien und Besonderheiten:"
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   103
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Erkrankungen:"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   101
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sonstiges:"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   99
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6720
         MouseIcon       =   "frmWKL94.frx":4A28
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   90
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Gesundheit / Risiken"
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
         Index           =   14
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1800
      Top             =   3600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Schließen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9720
      MouseIcon       =   "frmWKL94.frx":4D32
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   84
      Top             =   8160
      Width           =   1935
   End
End
Attribute VB_Name = "frmWKL94"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub LeseKundenStammWKL90()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    If gckundnr = "" Then
        Exit Sub
    End If
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & gckundnr
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!name) Then
            Label2(0).Caption = rsrs!name
        Else
            Label2(0).Caption = ""
        End If
        If Not IsNull(rsrs!vorname) Then
            Label2(1).Caption = rsrs!vorname
        Else
            Label2(1).Caption = ""
        End If
        If Not IsNull(rsrs!strasse) Then
            Label2(2).Caption = rsrs!strasse
        Else
            Label2(2).Caption = ""
        End If
        If Not IsNull(rsrs!Plz) Then
            Label2(3).Caption = rsrs!Plz
        Else
            Label2(3).Caption = ""
        End If
        If Not IsNull(rsrs!STADT) Then
            Label2(4).Caption = rsrs!STADT
        Else
            Label2(4).Caption = ""
        End If
        Label2(5).Caption = ""
        If Not IsNull(rsrs!Tel) Then
            Label2(6).Caption = rsrs!Tel
        Else
            Label2(6).Caption = ""
        End If
        
        If Not IsNull(rsrs!Mobiltel) Then
            Label2(7).Caption = rsrs!Mobiltel
        Else
            Label2(7).Caption = ""
        End If
        
        If Not IsNull(rsrs!KurzTEXT2) Then
            Label2(8).Caption = rsrs!KurzTEXT2
        Else
            Label2(8).Caption = ""
        End If
    Else
        Label2(0).Caption = ""
        Label2(1).Caption = ""
        Label2(2).Caption = ""
        Label2(3).Caption = ""
        Label2(4).Caption = ""
        Label2(5).Caption = ""
        Label2(6).Caption = ""
        Label2(7).Caption = ""
        Label2(8).Caption = ""
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Label2(0).Refresh
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseKundenStammWKL90"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo1()
    On Error GoTo LOKAL_ERROR
    
    Combo1.Clear
    Combo1.AddItem "bitte wählen"
    Combo1.AddItem "feuchtigkeitsarme Haut"
    Combo1.AddItem "trockene Haut"
    Combo1.AddItem "trockene Mischhaut"
    Combo1.AddItem "fettige Mischhaut"
    Combo1.AddItem "fettige Haut"
    Combo1.AddItem "trockene Aknehaut"
    Combo1.AddItem "fettige Aknehaut"
    Combo1.AddItem "allergische Haut"
    Combo1.AddItem "atrophische Haut"
    Combo1.AddItem "empfindliche Haut"
    Combo1.Text = "bitte wählen"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleCombo1"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LesePflegeDatenWKL90()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim lWert           As Long
    Dim sNageldetail    As String
    Dim sNagel          As String
    Dim sFuss           As String
    Dim sRISIKEN        As String
    Dim i               As Integer
    
    If gckundnr = "" Then
        Exit Sub
    End If
    
    cSQL = "Select * from KUNPFLEG where KUNDNR = " & gckundnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!HAUTTYP) Then
            Combo1.Text = rsrs!HAUTTYP
        Else
            Combo1.Text = "bitte wählen"
        End If
        
        If Not IsNull(rsrs!SCHUHGR) Then
            Text4(0).Text = rsrs!SCHUHGR
        Else
            Text4(0).Text = ""
        End If
        
        If Not IsNull(rsrs!FUSSINFO) Then
            Text4(1).Text = rsrs!FUSSINFO
        Else
            Text4(1).Text = ""
        End If
        
        If Not IsNull(rsrs!NAGELINFO) Then
            Text4(2).Text = rsrs!NAGELINFO
        Else
            Text4(2).Text = ""
        End If
        
        If Not IsNull(rsrs!FINGERINFO) Then
            Text4(3).Text = rsrs!FINGERINFO
        Else
            Text4(3).Text = ""
        End If
        
        If Not IsNull(rsrs!TIPTYP) Then
            Text4(4).Text = rsrs!TIPTYP
        Else
            Text4(4).Text = ""
        End If
        
        If Not IsNull(rsrs!VERPRONA) Then
            Text4(5).Text = rsrs!VERPRONA
        Else
            Text4(5).Text = ""
        End If
        
        If Not IsNull(rsrs!VERPROFU) Then
            Text4(6).Text = rsrs!VERPROFU
        Else
            Text4(6).Text = ""
        End If
        
        
        
        
        
        If Not IsNull(rsrs!VORGESCHM) Then
            Text1(0).Text = rsrs!VORGESCHM
        Else
            Text1(0).Text = ""
        End If
        
        If Not IsNull(rsrs!ERKRANK) Then
            Text1(1).Text = rsrs!ERKRANK
        Else
            Text1(1).Text = ""
        End If
        
        If Not IsNull(rsrs!ALLERGIEN) Then
            Text1(2).Text = rsrs!ALLERGIEN
        Else
            Text1(2).Text = ""
        End If
        
        If Not IsNull(rsrs!FETT) Then
            lWert = rsrs!FETT
            Option1(lWert).Value = True
        Else
            Option1(1).Value = True
        End If
        If Not IsNull(rsrs!FEUCHT) Then
            lWert = rsrs!FEUCHT
            Option1(lWert).Value = True
        Else
            Option1(6).Value = True
        End If
        If Not IsNull(rsrs!POREN) Then
            lWert = rsrs!POREN
            Option1(lWert).Value = True
        Else
            Option1(8).Value = True
        End If
        If Not IsNull(rsrs!SPANN) Then
            lWert = rsrs!SPANN
            Option1(lWert).Value = True
        Else
            Option1(11).Value = True
        End If
        If Not IsNull(rsrs!EMPFIND) Then
            lWert = rsrs!EMPFIND
            Option1(lWert).Value = True
        Else
            Option1(15).Value = True
        End If
        If Not IsNull(rsrs!ERSCHEIN) Then
            lWert = rsrs!ERSCHEIN
            'Binärvergleich!!!
            If lWert And 1 Then
                Check1(0).Value = vbChecked
            Else
                Check1(0).Value = vbUnchecked
            End If
            If lWert And 2 Then
                Check1(1).Value = vbChecked
            Else
                Check1(1).Value = vbUnchecked
            End If
            If lWert And 4 Then
                Check1(2).Value = vbChecked
            Else
                Check1(2).Value = vbUnchecked
            End If
            If lWert And 8 Then
                Check1(3).Value = vbChecked
            Else
                Check1(3).Value = vbUnchecked
            End If
        Else
            Check1(0).Value = vbUnchecked
            Check1(1).Value = vbUnchecked
            Check1(2).Value = vbUnchecked
            Check1(3).Value = vbUnchecked
            
        End If
        

        
        If Not IsNull(rsrs!ENTHAAR) Then
            lWert = rsrs!ENTHAAR
            'Binärvergleich!!!
            If lWert And 1 Then
                Check3(0).Value = vbChecked
            Else
                Check3(0).Value = vbUnchecked
            End If
            If lWert And 2 Then
                Check3(1).Value = vbChecked
            Else
                Check3(1).Value = vbUnchecked
            End If
            If lWert And 4 Then
                Check3(2).Value = vbChecked
            Else
                Check3(2).Value = vbUnchecked
            End If
            If lWert And 8 Then
                Check3(3).Value = vbChecked
            Else
                Check3(3).Value = vbUnchecked
            End If
            If lWert And 16 Then
                Check3(4).Value = vbChecked
            Else
                Check3(4).Value = vbUnchecked
            End If
            If lWert And 32 Then
                Check3(5).Value = vbChecked
            Else
                Check3(5).Value = vbUnchecked
            End If
        Else
            Check3(0).Value = vbUnchecked
            Check3(1).Value = vbUnchecked
            Check3(2).Value = vbUnchecked
            Check3(3).Value = vbUnchecked
            Check3(4).Value = vbUnchecked
            Check3(5).Value = vbUnchecked
        End If
        
        If Not IsNull(rsrs!HAAR) Then
            Text2(0).Text = rsrs!HAAR
        Else
            Text2(0).Text = ""
        End If
                
        If Not IsNull(rsrs!AUGEN) Then
            Text2(1).Text = rsrs!AUGEN
        Else
            Text2(1).Text = ""
        End If
                
        If Not IsNull(rsrs!TEINT) Then
            Text2(2).Text = rsrs!TEINT
        Else
            Text2(2).Text = ""
        End If
                
        If Not IsNull(rsrs!GESICHT) Then
            Text2(3).Text = rsrs!GESICHT
        Else
            Text2(3).Text = ""
        End If
                
        If Not IsNull(rsrs!FARBEN) Then
            Text2(4).Text = rsrs!FARBEN
        Else
            Text2(4).Text = ""
        End If
                
        If Not IsNull(rsrs!DUFT) Then
            Text2(5).Text = rsrs!DUFT
        Else
            Text2(5).Text = ""
        End If
                
        If Not IsNull(rsrs!Makeup) Then
            Text2(6).Text = rsrs!Makeup
        Else
            Text2(6).Text = ""
        End If
        
        If Not IsNull(rsrs!RAUCHER) Then
            Text2(11).Text = rsrs!RAUCHER
        Else
            Text2(11).Text = ""
        End If
        
        If Not IsNull(rsrs!KAFFEEGENUSS) Then
            Text2(10).Text = rsrs!KAFFEEGENUSS
        Else
            Text2(10).Text = ""
        End If
        
        If Not IsNull(rsrs!SONNENBANK) Then
            Text2(9).Text = rsrs!SONNENBANK
        Else
            Text2(9).Text = ""
        End If
                
        If Not IsNull(rsrs!WASSER) Then
            Text2(8).Text = rsrs!WASSER
        Else
            Text2(8).Text = ""
        End If
    
        If Not IsNull(rsrs!sPort) Then
            Text2(7).Text = rsrs!sPort
        Else
            Text2(7).Text = ""
        End If
        
        If Not IsNull(rsrs!FUSS) Then
            sFuss = rsrs!FUSS
        Else
            sFuss = ""
        End If
        
        If Not IsNull(rsrs!RISIKEN) Then
            sRISIKEN = rsrs!RISIKEN
        Else
            sRISIKEN = ""
        End If
        
        If Not IsNull(rsrs!FNAGEL) Then
            sNagel = rsrs!FNAGEL
        Else
            sNagel = ""
        End If
        
        If Not IsNull(rsrs!FNAGELD) Then
            sNageldetail = rsrs!FNAGELD
        Else
            sNageldetail = ""
        End If
        
        If sNageldetail <> "" Then
    
            Dim sArray() As String
            sArray = Split(sNageldetail, " ")
            
            For i = 0 To UBound(sArray) - 1
                Check4(CInt(Trim(sArray(i)))).Value = vbChecked
            Next i
        End If
        
        If sNagel <> "" Then
    
            sArray = Split(sNagel, " ")
            
            For i = 0 To UBound(sArray) - 1
                Check2(CInt(Trim(sArray(i)))).Value = vbChecked
            Next i
        End If
        
        If sRISIKEN <> "" Then
    
            sArray = Split(sRISIKEN, " ")
            
            For i = 0 To UBound(sArray) - 1
                Check5(CInt(Trim(sArray(i)))).Value = vbChecked
            Next i
        End If
        
        If sFuss <> "" Then
    
            sArray = Split(sFuss, " ")
            
            For i = 0 To UBound(sArray) - 1
                Check6(CInt(Trim(sArray(i)))).Value = vbChecked
            Next i
        End If
        
        For i = 0 To 14
            If Check2(i).Value = vbChecked Then
                Fuss_OPT_einausblenden i, True
            Else
                Fuss_OPT_einausblenden i, False
            End If
        Next i
        
        Option2(0).Value = True
        If Not IsNull(rsrs!EINLAGEHERST) Then
            If rsrs!EINLAGEHERST = True Then
                Option2(0).Value = True
            Else
                Option2(1).Value = True
            End If
        End If
        
    Else
        Text1(0).Text = ""
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text2(0).Text = ""
        Text2(1).Text = ""
        Text2(2).Text = ""
        Text2(3).Text = ""
        Text2(4).Text = ""
        Text2(5).Text = ""
        Text2(6).Text = ""
        
        Text2(7).Text = ""
        Text2(8).Text = ""
        Text2(9).Text = ""
        Text2(10).Text = ""
        Text2(11).Text = ""
        
        Option1(1).Value = True
        Option1(6).Value = True
        Option1(8).Value = True
        Option1(11).Value = True
        Option1(15).Value = True
        Check1(0).Value = vbUnchecked
        Check1(1).Value = vbUnchecked
        Check1(2).Value = vbUnchecked
        Check1(3).Value = vbUnchecked
'        Check2(0).Value = vbUnchecked
'        Check2(1).Value = vbUnchecked
'        Check2(2).Value = vbUnchecked
'        Check2(3).Value = vbUnchecked
'        Check2(4).Value = vbUnchecked
        Check3(0).Value = vbUnchecked
        Check3(1).Value = vbUnchecked
        Check3(2).Value = vbUnchecked
        Check3(3).Value = vbUnchecked
        Check3(4).Value = vbUnchecked
        Check3(5).Value = vbUnchecked
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LesePflegeDatenWKL90"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub LesePflegeHistorieDetailWKL90(cDatum As String, cBeh As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from KUNPFLGH where KUNDNR = " & gckundnr & " and DATUM = " & cDatum & " and BEHANDLUNG = '" & cBeh & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!REINIGUNG) Then
            Text3(1).Text = rsrs!REINIGUNG
        Else
            Text3(1).Text = ""
        End If
        If Not IsNull(rsrs!MASKEN) Then
            Text3(2).Text = rsrs!MASKEN
        Else
            Text3(2).Text = ""
        End If
        If Not IsNull(rsrs!EXTRAS) Then
            Text3(3).Text = rsrs!EXTRAS
        Else
            Text3(3).Text = ""
        End If
        If Not IsNull(rsrs!TAGESPFLEG) Then
            Text3(4).Text = rsrs!TAGESPFLEG
        Else
            Text3(4).Text = ""
        End If
        If Not IsNull(rsrs!BERATUNG) Then
            Text3(5).Text = rsrs!BERATUNG
        Else
            Text3(5).Text = ""
        End If
        If Not IsNull(rsrs!PROBEN) Then
            Text3(6).Text = rsrs!PROBEN
        Else
            Text3(6).Text = ""
        End If
        If Not IsNull(rsrs!KAEUFE) Then
            Text3(7).Text = rsrs!KAEUFE
        Else
            Text3(7).Text = ""
        End If
    
        If Not IsNull(rsrs!TONIC) Then
            Text3(8).Text = rsrs!TONIC
        Else
            Text3(8).Text = ""
        End If
        
        If Not IsNull(rsrs!PEELING) Then
            Text3(9).Text = rsrs!PEELING
        Else
            Text3(9).Text = ""
        End If
        
        If Not IsNull(rsrs!MASSAGE) Then
            Text3(10).Text = rsrs!MASSAGE
        Else
            Text3(10).Text = ""
        End If
        
        If Not IsNull(rsrs!NAGELLACK) Then
            Text3(12).Text = rsrs!NAGELLACK
        Else
            Text3(12).Text = ""
        End If
    Else
        Text3(11).Text = Format$(Val(cDatum), "DD.MM.YYYY")
        Text3(1).Text = ""
        Text3(2).Text = ""
        Text3(3).Text = ""
        Text3(4).Text = ""
        Text3(5).Text = ""
        Text3(6).Text = ""
        Text3(7).Text = ""
        Text3(8).Text = ""
        Text3(9).Text = ""
        Text3(10).Text = ""
        Text3(12).Text = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LesePflegeHistorieDetailWKL90"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SchreibePflegeHistorieDetailWKL90()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cDatum As String
    Dim cBehandlung As String
    Dim lWert As Long
    
    If Label3(0).Caption = "" Then
        Exit Sub
    End If
    
    If Label3(8).Caption = "" Then
        Exit Sub
    End If
    
    cDatum = Label3(0).Caption
    If Not IsDate(cDatum) Then
'        MsgBox "Ungültige Eingabe im Feld 'Datum'!", vbCritical, "Winkiss Hinweis:"
'        Text3(11).SetFocus
        Exit Sub
    End If
    
    lWert = DateValue(cDatum)
    cDatum = Format$(lWert, "DD.MM.YYYY")
    
    cBehandlung = Trim(Label3(8).Caption)
    
    cSQL = "Select * from KUNPFLGH where KUNDNR = " & gckundnr & " and DATUM = " & Trim$(Str$(lWert)) & " "
    cSQL = cSQL & " and Behandlung = '" & cBehandlung & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    rsrs!Kundnr = gckundnr
    rsrs!Datum = lWert
    rsrs!REINIGUNG = Text3(1).Text
    rsrs!MASKEN = Text3(2).Text
    rsrs!EXTRAS = Text3(3).Text
    rsrs!TAGESPFLEG = Text3(4).Text
    rsrs!BERATUNG = Text3(5).Text
    rsrs!PROBEN = Text3(6).Text
    rsrs!KAEUFE = Text3(7).Text
    rsrs!TONIC = Text3(8).Text
    rsrs!PEELING = Text3(9).Text
    rsrs!MASSAGE = Text3(10).Text
    rsrs!NAGELLACK = Text3(12).Text
    rsrs!Behandlung = cBehandlung
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
'    LesePflegeHistorieWKL90
'    LesePflegeBH
    
'    Text3(11).Text = Format$(Now, "DD.MM.YYYY")
'    Text3(1).Text = ""
'    Text3(2).Text = ""
'    Text3(3).Text = ""
'    Text3(4).Text = ""
'    Text3(5).Text = ""
'    Text3(6).Text = ""
'    Text3(7).Text = ""
'    Text3(8).Text = ""
'    Text3(9).Text = ""
'    Text3(10).Text = ""
'    Text3(12).Text = ""
    
'    List1.SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibePflegeHistorieDetailWKL90"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub SchreibePflegeDatenWKL90()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim lWert           As Long
    Dim iIndex          As Integer
    Dim iCount          As Integer
    Dim sNageldetail    As String
    Dim sNagel          As String
    Dim sFuss           As String
    Dim sRISIKEN        As String
    ReDim iHautBild(0 To 4) As Integer
    
    cSQL = "Select * from KUNPFLEG where KUNDNR = " & gckundnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    rsrs!Kundnr = gckundnr
    rsrs!VORGESCHM = Text1(0).Text
    rsrs!ERKRANK = Text1(1).Text
    rsrs!ALLERGIEN = Text1(2).Text
    
    rsrs!SCHUHGR = Text4(0).Text
    rsrs!FUSSINFO = Text4(1).Text
    rsrs!NAGELINFO = Text4(2).Text
    rsrs!FINGERINFO = Text4(3).Text
    rsrs!TIPTYP = Text4(4).Text
    rsrs!VERPRONA = Text4(5).Text
    rsrs!VERPROFU = Text4(6).Text
    
    sRISIKEN = ""
    For iCount = 0 To 2
        If Check5(iCount).Value = vbChecked Then
            sRISIKEN = sRISIKEN & iCount & " "
        End If
    Next iCount
    rsrs!RISIKEN = sRISIKEN
    
    sFuss = ""
    For iCount = 0 To 16
        If Check6(iCount).Value = vbChecked Then
            sFuss = sFuss & iCount & " "
        End If
    Next iCount
    rsrs!FUSS = sFuss
    
    sNageldetail = ""
    For iCount = 0 To 79
        If Check4(iCount).Value = vbChecked Then
            sNageldetail = sNageldetail & iCount & " "
        End If
    Next iCount
    rsrs!FNAGELD = sNageldetail
    
    sNagel = ""
    For iCount = 0 To 14
        If Check2(iCount).Value = vbChecked Then
            sNagel = sNagel & iCount & " "
        End If
    Next iCount
    rsrs!FNAGEL = sNagel
    
   
    
    If Option2(0).Value = True Then
        rsrs!EINLAGEHERST = True
    ElseIf Option2(1).Value = True Then
        rsrs!EINLAGEHERST = False
    End If
    
    
    
    
    
    iIndex = 0
    For iCount = 0 To 16
        If Option1(iCount).Value = True Then
            iHautBild(iIndex) = iCount
            iIndex = iIndex + 1
        End If
    Next iCount
    
    rsrs!FETT = iHautBild(0)
    rsrs!FEUCHT = iHautBild(1)
    rsrs!POREN = iHautBild(2)
    rsrs!SPANN = iHautBild(3)
    rsrs!EMPFIND = iHautBild(4)
    
    lWert = 0
    For iCount = 0 To 3
        If Check1(iCount).Value = vbChecked Then
            lWert = lWert + (2 ^ iCount)
        End If
    Next iCount
    rsrs!ERSCHEIN = lWert
    
    lWert = 0
'    For iCount = 0 To 4
'        If Check2(iCount).Value = vbChecked Then
'            lWert = lWert + (2 ^ iCount)
'        End If
'    Next iCount
    rsrs!BEHAND = lWert
    
    lWert = 0
    For iCount = 0 To 5
        If Check3(iCount).Value = vbChecked Then
            lWert = lWert + (2 ^ iCount)
        End If
    Next iCount
    rsrs!ENTHAAR = lWert
    
    rsrs!HAAR = Text2(0).Text
    rsrs!AUGEN = Text2(1).Text
    rsrs!TEINT = Text2(2).Text
    rsrs!GESICHT = Text2(3).Text
    rsrs!FARBEN = Text2(4).Text
    rsrs!DUFT = Text2(5).Text
    rsrs!Makeup = Text2(6).Text
    rsrs!HAUTTYP = Combo1.Text
    
    rsrs!RAUCHER = Text2(11).Text
    rsrs!SONNENBANK = Text2(9).Text
    rsrs!KAFFEEGENUSS = Text2(10).Text
    rsrs!sPort = Text2(7).Text
    rsrs!WASSER = Text2(8).Text
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibePflegeDatenWKL90"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub LesePflegeBH()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim i As Integer
    Dim lDate As Long
    Dim sBehandlung As String
    Dim sTmp As String
    
    If gckundnr = "" Then
        Exit Sub
    End If
    
    List3.Clear
    
    cSQL = "select * from  KUNPFLGH where KUNDNR = " & gckundnr & " order by datum desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!Datum) Then
                lDate = rsrs!Datum
            Else
            
            End If
            
            If Not IsNull(rsrs!Behandlung) Then
                sBehandlung = rsrs!Behandlung
            Else
            
            End If
            
            List3.AddItem Format$(lDate, "DD.MM.YYYY") & Space(1) & sBehandlung
        
        rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    List3.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LesePflegeBH"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Check2(Index).Value = vbChecked Then
        Fuss_OPT_einausblenden Index, True
    Else
        Fuss_OPT_einausblenden Index, False
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Fuss_OPT_einausblenden(iBegin As Integer, bEinblenden As Boolean)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim j As Integer
    
    If iBegin < 8 Then
        For i = 0 To 9
            
            If iBegin = 0 Then
                j = i
            Else
                j = (iBegin * 10) + i
            End If
            Check4(j).Visible = bEinblenden
            
            If bEinblenden = False Then
                Check4(j).Value = vbUnchecked
            End If
        Next i
    
    ElseIf iBegin >= 8 Then
    
        Select Case iBegin
            Case 8
                Check6(0).Visible = bEinblenden
                Check6(1).Visible = bEinblenden
                Check6(2).Visible = bEinblenden
                Check6(3).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(0).Value = vbUnchecked
                    Check6(1).Value = vbUnchecked
                    Check6(2).Value = vbUnchecked
                    Check6(3).Value = vbUnchecked
                End If
            Case 9
                Check6(4).Visible = bEinblenden
                Check6(5).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(4).Value = vbUnchecked
                    Check6(5).Value = vbUnchecked
                End If
                
            Case 10
                Check6(6).Visible = bEinblenden
                Check6(7).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(6).Value = vbUnchecked
                    Check6(7).Value = vbUnchecked
                End If
            
            Case 11
                Check6(8).Visible = bEinblenden
                Check6(9).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(8).Value = vbUnchecked
                    Check6(9).Value = vbUnchecked
                End If
                
            Case 12
                Check6(10).Visible = bEinblenden
                Check6(11).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(10).Value = vbUnchecked
                    Check6(11).Value = vbUnchecked
                End If
                
            Case 13
                Check6(12).Visible = bEinblenden
                Check6(13).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(12).Value = vbUnchecked
                    Check6(13).Value = vbUnchecked
                End If
                
            Case 14
                Check6(16).Visible = bEinblenden
                Check6(15).Visible = bEinblenden
            
                If bEinblenden = False Then
                    Check6(16).Value = vbUnchecked
                    Check6(15).Value = vbUnchecked
                End If
            
        
        End Select
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Fuss_OPT_einausblenden"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Check3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    If Index = 12 Then
        If Check3(12).Value = vbChecked Then
            Check3(6).Value = vbChecked
            Check3(9).Value = vbChecked
            Check3(16).Value = vbChecked
            Check3(7).Value = vbChecked
            Check3(8).Value = vbChecked
            Check3(11).Value = vbChecked
            Check3(13).Value = vbChecked
'            Check3(21).Value = vbChecked
'            Check3(22).Value = vbChecked
'            Check3(23).Value = vbChecked
            Check3(14).Value = vbChecked
            Check3(25).Value = vbChecked
            Check3(26).Value = vbChecked
            Check3(27).Value = vbChecked
'            Check3(28).Value = vbChecked
            Check3(29).Value = vbChecked
        Else
            Check3(6).Value = vbUnchecked
            Check3(9).Value = vbUnchecked
            Check3(16).Value = vbUnchecked
            Check3(7).Value = vbUnchecked
            Check3(8).Value = vbUnchecked
            Check3(11).Value = vbUnchecked
            Check3(13).Value = vbUnchecked
'            Check3(21).Value = vbUnchecked
'            Check3(22).Value = vbUnchecked
'            Check3(23).Value = vbUnchecked
            Check3(14).Value = vbUnchecked
            Check3(25).Value = vbUnchecked
            Check3(26).Value = vbUnchecked
            Check3(27).Value = vbUnchecked
'            Check3(28).Value = vbUnchecked
            Check3(29).Value = vbUnchecked
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check6_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 14 Then
        If Check6(14).Value = vbChecked Then
            Option2(0).Visible = True
            Option2(1).Visible = True
            Option2(0).Value = True
        Else
            Option2(0).Visible = False
            Option2(1).Visible = False
            
            Text4(0).Visible = False
            Text4(0).Text = ""
            Label1(42).Visible = False
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check6_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo4_Click()
On Error GoTo LOKAL_ERROR

LeseStandardTexteWKL94 Combo4.Text

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseStandardTexteWKL94(sKrit As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim ctmp As String
    Dim lWert As Long
    
    List2.Clear
    
    If sKrit = "" Then
        cSQL = "Select * from TERM_STD order by NR "
    Else
        cSQL = "Select * from TERM_STD where Gliederung = '" & sKrit & "' order by NR "
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEZEICH) Then
                ctmp = rsrs!BEZEICH
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            
            List2.AddItem ctmp
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseStandardTexteWKL94"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            Text3(0).Text = Format(Datumschreiben11a(4000, 5000), "DD.MM.YYYY")
        Case 1
            Text3(11).Text = Format(Datumschreiben11a(4000, 5000), "DD.MM.YYYY")
            'fertig
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub wähle()
    On Error GoTo LOKAL_ERROR

    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "wähle"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKL90
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fuellecombo1
    FuelleListBeh
    
    LeseKundenStammWKL90
    LesePflegeDatenWKL90
'    LesePflegeHistorieWKL90
    LesePflegeBH
    
    Label1(15).ForeColor = vbRed
    
    fuellecombo
    
    Zeige_Kunden_Bilder gckundnr, 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    Combo4.Clear
    
    sSQL = "select distinct(gliederung) from TERM_STD  order by gliederung "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Gliederung) Then
                Combo4.AddItem rsrs!Gliederung
                If Combo4.Text = "" Then
                    Combo4.Text = rsrs!Gliederung
                End If
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
    Fehler.gsFunktion = "fuellecombo"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zeige_Kunden_Bilder(sKUNDNR As String, iStatus As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim sSpeicherpfad   As String
    Dim i               As Integer
    Dim sBildname       As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\KUNDEN"
    
'    For i = 0 To 2
'        Picture5(i).Enabled = False
'        Picture5(i).Visible = False
'    Next i
    
    If FileExists(sPfad & "\" & sKUNDNR & ".jpg") Then
        Image1.Picture = LoadPicture(sPfad & "\" & sKUNDNR & ".jpg")
'        Image2.Picture = LoadPicture(sPfad & "\" & sKUNDNR & ".jpg")
        sSpeicherpfad = sPfad & "\" & sKUNDNR & "_s.jpg"
'        Label5(12).Caption = sKUNDNR & ".jpg"
        
        Picture3.Visible = True
        
        File1.Path = sPfad
        File1.Pattern = sKUNDNR & "*.jpg"
        File1.Refresh
                    
'''        If File1.ListCount > 1 Then
'''            'dann zeige weitere
'''            For i = 0 To File1.ListCount - 1
'''
'''                If i > 2 Then
'''                    Exit For
'''                End If
'''                sBildname = File1.list(i)
'''                Picture5(i).Tag = sBildname
'''                Zeige_weitere_Kunden_Bilder sBildname, 0, Picture5(i), Image4, 50
'''
'''                If i > 0 Then
'''                    Picture5(i).Top = Picture5(0).Top
'''                    Picture5(i).Left = Picture5(i - 1).Left + 50 + Picture5(i - 1).Width
'''                End If
'''            Next i
'''        End If
    Else
        sSpeicherpfad = ""
        
        
        Dim sGeschlecht As String
    
        sGeschlecht = lookingForKundendaten(sKUNDNR).geschlecht
        
        If UCase(sGeschlecht) = "W" Then
            sBildname = "keinBildw.jpg"
        ElseIf UCase(sGeschlecht) = "M" Then
            sBildname = "keinBildm.jpg"
        Else
            Dim sAnrede As String
    
            sAnrede = lookingForKundendaten(sKUNDNR).anrede
            
            If UCase(sAnrede) = "FRAU" Then
                sBildname = "keinBildw.jpg"
            ElseIf UCase(sAnrede) = "HERR" Then
                sBildname = "keinBildm.jpg"
            Else
                sBildname = "keinBildw.jpg"
            End If
        End If
        
        
        
        
        If FileExists(sPfad & "\" & sBildname) Then
            Image1.Picture = LoadPicture(sPfad & "\" & sBildname)
'            Image2.Picture = LoadPicture(sPfad & "\" & sBildname)
            
'            Label5(12).Caption = "keinBild.jpg"
            Picture3.Visible = True
        Else
            Picture3.Visible = False
'            Label5(12).Caption = ""
            
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
    
    Picture3.Tag = sKUNDNR & ".jpg"
    zeigImage_Kunden_In_Picture Image1, Picture3, iDiv, "" 'sSpeicherpfad

    Picture3.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 481 Then
        If iStatus = 1 Then
            MsgBox "Dieses Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        End If
        Kill sPfad & "\" & sKUNDNR & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Zeige_Kunden_Bilder"
        Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub zeigImage_Kunden_In_Picture(imgx As Image, PicX As PictureBox, iDiv As Integer, sSpeicherpfad As String)
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
    PicX.Refresh
    
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
            SavePicture .Image, sSpeicherpfad
        End If
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigImage_Kunden_In_Picture"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleListBeh()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim ctmp As String
    Dim lWert As Long
    
    List2.Clear
    
    cSQL = "Select * from TERM_STD "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!BEZEICH) Then
                ctmp = rsrs!BEZEICH
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            
            List2.AddItem ctmp
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListBeh"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL90()
    On Error GoTo LOKAL_ERROR
    
    Frame3.Top = 1800
    Frame3.Left = 3000
    Frame3.Height = 6255
    Frame3.Width = 8775
    
    Frame2.Top = 1800
    Frame2.Left = 3000
    Frame2.Height = 6255
    Frame2.Width = 8775
    Frame2.Visible = False
    
    Frame5.Top = 1800
    Frame5.Left = 3000
    Frame5.Height = 6255
    Frame5.Width = 8775
    Frame5.Visible = False
    
    Frame7.Top = 1800
    Frame7.Left = 3000
    Frame7.Height = 6255
    Frame7.Width = 8775
    Frame7.Visible = False
    
    Frame8.Top = 1800
    Frame8.Left = 3000
    Frame8.Height = 6255
    Frame8.Width = 8775
    Frame8.Visible = False

    Frame9.Top = 1800
    Frame9.Left = 3000
    Frame9.Height = 6255
    Frame9.Width = 8775
    Frame9.Visible = False
    
    Frame10.Top = 1800
    Frame10.Left = 3000
    Frame10.Height = 6255
    Frame10.Width = 8775
    Frame10.Visible = False
    
    Frame11.Top = 1800
    Frame11.Left = 3000
    Frame11.Height = 6255
    Frame11.Width = 8775
    Frame11.Visible = False
   

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL90"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(6).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(2).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame10_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(1).ForeColor = glS1
    Label5(11).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame11_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(4).ForeColor = glS1
    Label5(12).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame2_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(26).ForeColor = glS1
    Label5(13).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(3).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame5_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label4(0).ForeColor = glS1
    Label4(1).ForeColor = glS1
    Label4(2).ForeColor = glS1
    Label4(3).ForeColor = glS1
    Label4(4).ForeColor = glS1
    Label4(5).ForeColor = glS1
    Label4(6).ForeColor = glS1
    Label4(7).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame6_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(9).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame7_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(10).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame8_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(0).ForeColor = glS1
    Label5(5).ForeColor = glS1
    Label5(8).ForeColor = glS1
    Label5(7).ForeColor = glS1
    Label5(14).ForeColor = glS1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame9_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Frame3.Visible = False
    Frame9.Visible = False
    Frame10.Visible = False
    Frame5.Visible = False
    Frame2.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
    Frame11.Visible = False
    Text3(11).Text = Format$(Now, "DD.MM.YYYY")

    Select Case Index
        Case 0  'Hautbild
            Frame3.Visible = True
        Case 1  'Einwilligungen
            Frame11.Visible = True
        Case 2  'verwendete Produkte
            Frame9.Visible = True
        Case 3  'Enthaarung
            Frame10.Visible = True
        Case 4  'Kundentyp
            Frame5.Visible = True
        Case 5  'kosm. Vorgeschichte
            Frame2.Visible = True
        Case 6  'Hände
            Frame7.Visible = True
        Case 7  'Füße
            Frame8.Visible = True
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label4(Index).ForeColor = glLink
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim lcount As Long
    Dim bFound As Boolean
    Dim sBeh As String
    Dim lDate As Long

    Select Case Index
    
        Case 14
            drucken_Kosmetik gckundnr, "verwendete Produkte"
        Case 13
            drucken_Kosmetik gckundnr, "Hautbild"
        Case 12
            drucken_Kosmetik gckundnr, "Gesundheit"
        Case 11
            drucken_einwilligung gckundnr
        Case 6
            'vorher immer speichern
            SchreibePflegeDatenWKL90
            SchreibePflegeHistorieDetailWKL90
            
            gckundnr = ""
            Unload frmWKL94
        Case 26, 0, 2, 3, 4, 9, 10, 1  'Speichern Reihenfolge!
            SchreibePflegeDatenWKL90
            SchreibePflegeHistorieDetailWKL90
            
        Case 5
            If List2.list(List2.ListIndex) <> "" Then
                List3.AddItem Text3(11).Text & Space(1) & List2.list(List2.ListIndex)
                
                sBeh = Trim(List2.list(List2.ListIndex))
                lDate = CLng(DateValue(Text3(11).Text))
                
                cSQL = "Delete from  KUNPFLGH where KUNDNR = " & gckundnr & " "
                cSQL = cSQL & " and BEHANDLUNG = '" & sBeh & "'"
                cSQL = cSQL & " and DATUM = " & lDate & " "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Insert into KUNPFLGH (KUNDNR,DATUM,BEHANDLUNG) values "
                cSQL = cSQL & " ( " & gckundnr & "," & lDate & ",'" & sBeh & "')"
                gdBase.Execute cSQL, dbFailOnError
                
                anzeige "normal", "Diese Behandlungen wurden durchgeführt", Label1(30)
            End If
        Case 7
            If Label3(0).Caption <> "" Then
                lDate = CLng(DateValue(Label3(0).Caption))
                    
                If lDate > 0 Then
                    cSQL = "Delete from  KUNPFLGH where KUNDNR = " & gckundnr & " "
                    
                    If Label3(8).Caption <> "" Then
                        cSQL = cSQL & " and behandlung = '" & Label3(8).Caption & "'"
                    End If
                    
                    cSQL = cSQL & " and DATUM = " & lDate & " "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    anzeige "normal", "Die Datei/en wurde/n gelöscht!", Label1(30)
                Else
                    anzeige "rot", "Bitte Eintrag markieren!", Label1(30)
                End If
            Else
                anzeige "rot", "Bitte Eintrag markieren!", Label1(30)
            End If
            
            Label3(8).Caption = ""
            Label3(0).Caption = ""
            
            LesePflegeBH
        Case 8
            List3.Clear
            Label3(8).Caption = ""
            Label3(0).Caption = ""
            
            cSQL = "Delete from  KUNPFLGH where KUNDNR = " & gckundnr & " "
            gdBase.Execute cSQL, dbFailOnError
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub drucken_einwilligung(cKundnr As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    loeschNEW "KU_EINWIL", gdBase
    CreateTableT2 "KU_EINWIL", gdBase
    
    sSQL = "Insert into KU_EINWIL select "
    sSQL = sSQL & " TEL "
    sSQL = sSQL & ", VORNAME "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", NAME "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", stadt as ort "
    sSQL = sSQL & ", TITEL "
    sSQL = sSQL & ", FIRMA "
    sSQL = sSQL & ", RABATT "
    sSQL = sSQL & ", DATUM1 "
    sSQL = sSQL & ", trim(Titel) & ' ' & Name & ', ' & Vorname as NameVorname "
    sSQL = sSQL & ", EMAIL  "
    sSQL = sSQL & " from Kunden where kundnr = " & cKundnr & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If Check3(7).Value = vbChecked Then 'EV JetPeel
        reportbildschirm "", "aWKL94x"
    End If
    
    
'    If Check3(11).Value = vbChecked Then
'        reportbildschirm "", "aWKL94f"
'    End If
    
'    If Check3(15).Value = vbChecked Then 'LCN Micro Needling
'        reportbildschirm "", "aWKL94i"
'    End If
    
    If Check3(16).Value = vbChecked Then 'Monteil neu
        reportbildschirm "", "aWKL94j"
    End If
    
'    '*neue
'    If Check3(17).Value = vbChecked Then 'Pflegehinweis Braue
'        reportbildschirm "", "aWKL94k"
'    End If
'
'    If Check3(18).Value = vbChecked Then 'Pflegehinweis Lippen
'        reportbildschirm "", "aWKL94l"
'    End If
'
'    If Check3(19).Value = vbChecked Then 'Pflegehinweis Lid
'        reportbildschirm "", "aWKL94m"
'    End If
    
'    If Check3(20).Value = vbChecked Then 'Pflegehinweis Wimpern
'        reportbildschirm "", "aWKL94n"
'    End If
    
'    If Check3(21).Value = vbChecked Then 'Neu permanent Make-Up
'        reportbildschirm "", "aWKL94o"
'    End If
'
'    If Check3(22).Value = vbChecked Then 'Neu Micro Needling
'        reportbildschirm "", "aWKL94p"
'    End If
'
'    If Check3(23).Value = vbChecked Then 'Neu Säurepeeling
'        reportbildschirm "", "aWKL94q"
'    End If
    
'    If Check3(24).Value = vbChecked Then 'Neu Lashes
'        reportbildschirm "", "aWKL94r"
'    End If
    
    If Check3(25).Value = vbChecked Then 'Neu Fragebogen
        reportbildschirm "", "aWKL94s"
    End If
    
    If Check3(6).Value = vbChecked Then 'Neu
        reportbildschirm "", "aWKL94st"
    End If
    
    If Check3(26).Value = vbChecked Then 'Neu SMS
    
        If SpalteInTabellegefundenNEW("KU_EINWIL", "MOBILTEL", gdBase) = False Then
            sSQL = " Alter table KU_EINWIL add MOBILTEL Text(20)  "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        sSQL = "Update KU_EINWIL inner join Kunden on KU_EINWIL.kundnr= Kunden.kundnr set KU_EINWIL.MOBILTEL  = kunden.MOBILTEL "
        gdBase.Execute sSQL, dbFailOnError
        
        
        If Modul6.FindFile(gcDBPfad, "aWKL94t_woidtke.rpt") Then
            reportbildschirm "", "aWKL94t_woidtke"
        Else
            reportbildschirm "", "aWKL94t"
        End If
        
        If SpalteInTabellegefundenNEW("KU_EINWIL", "MOBILTEL", gdBase) = True Then
            sSQL = " Alter table KU_EINWIL drop MOBILTEL   "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
    End If
    
    If Check3(27).Value = vbChecked Then 'Neu Fragebogen
        If SpalteInTabellegefundenNEW("KU_EINWIL", "MOBILTEL", gdBase) = False Then
            sSQL = " Alter table KU_EINWIL add MOBILTEL Text(20)  "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        sSQL = "Update KU_EINWIL inner join Kunden on KU_EINWIL.kundnr= Kunden.kundnr set KU_EINWIL.MOBILTEL  = kunden.MOBILTEL "
        gdBase.Execute sSQL, dbFailOnError
        
        reportbildschirm "", "aWKL94u"
        
        If SpalteInTabellegefundenNEW("KU_EINWIL", "MOBILTEL", gdBase) = True Then
            sSQL = " Alter table KU_EINWIL drop MOBILTEL   "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
'    If Check3(28).Value = vbChecked Then 'EV Hydrafacial
'
'        sSQL = "Update KU_EINWIL set Firma  = '" & gFirma.FirmaName & "' "
'        gdBase.Execute sSQL, dbFailOnError
'
'        reportbildschirm "", "aWKL94v"
'    End If
    
    If Check3(29).Value = vbChecked Then 'Einwilligung Minderjährige
        reportbildschirm "", "aWKL94w"
    End If
    
    If Check3(8).Value = vbChecked Then
        reportbildschirm "", "aWKL94c"
    End If
    
    If Check3(11).Value = vbChecked Then
        reportbildschirm "", "aWKL94a"
    End If
    
    If Check3(13).Value = vbChecked Then
        reportbildschirm "", "aWKL94b"
    End If
    
    If Check3(14).Value = vbChecked Then
        reportbildschirm "", "aWKL94e"
    End If
    
    If Check3(9).Value = vbChecked Then
        reportbildschirm "", "aWKL94d"
    End If
    
    
    Pause 4
    loeschNEW "KU_EINWIL", gdBase

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken_einwilligung"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub drucken_Kosmetik(cKundnr As String, sThema As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    loeschNEW "KU_EINWIL", gdBase
    CreateTableT2 "KU_EINWIL", gdBase
    
    sSQL = "Insert into KU_EINWIL select "
    sSQL = sSQL & " TEL "
    sSQL = sSQL & ", VORNAME "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", NAME "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", stadt as ort "
    sSQL = sSQL & ", TITEL "
    sSQL = sSQL & ", FIRMA "
    sSQL = sSQL & ", RABATT "
    sSQL = sSQL & ", DATUM1 "
    sSQL = sSQL & ", trim(Titel) & ' ' & Name & ', ' & Vorname as NameVorname "
    sSQL = sSQL & ", EMAIL  "
    sSQL = sSQL & " from Kunden where kundnr = " & cKundnr & " "
    gdBase.Execute sSQL, dbFailOnError
    
    Select Case sThema
        Case "Gesundheit"
        
'            sSQL = "Alter Table from Kunden where kundnr = " & cKundnr & " "
'            gdBase.Execute sSQL, dbFailOnError
        Case "Hautbild"
        
            loeschNEW "KUNHAUTBILD_PRINT", gdBase
            CreateTableT3 "KUNHAUTBILD_PRINT", gdBase
            
            sSQL = "Insert into KUNHAUTBILD_PRINT select "
            sSQL = sSQL & " KUNDNR  "
            
            
            sSQL = sSQL & ", '' as TEXT1 "
            sSQL = sSQL & ", '' as TEXT2 "
            sSQL = sSQL & ", '' as TEXT3 "
            sSQL = sSQL & ", '' as TEXT4 "
            sSQL = sSQL & ", '' as TEXT5 "
            sSQL = sSQL & ", '' as TEXT6 "
            sSQL = sSQL & ", '' as TEXT7 "
            sSQL = sSQL & ", '' as TEXT8 "
            sSQL = sSQL & ", '' as TEXT9 "
            sSQL = sSQL & ", '' as TEXT10 "
            
            sSQL = sSQL & " from KU_EINWIL where kundnr = " & cKundnr & " "
            gdBase.Execute sSQL, dbFailOnError
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text1 = '" & Combo1.Text & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
        
            
            Dim sWas As String
            
            sWas = "Nein"
            
            If Option1(0).Value = True Then
                sWas = "trocken"
            
            ElseIf Option1(1).Value = True Then
                sWas = "normal"
            
            ElseIf Option1(2).Value = True Then
                sWas = "fett"
            ElseIf Option1(3).Value = True Then
                sWas = "T-Zone-Nase"
            End If
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text2 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            
            If Option1(4).Value = True Then
                sWas = "sehr arm"
            
            ElseIf Option1(5).Value = True Then
                sWas = "arm"
            
            ElseIf Option1(6).Value = True Then
                sWas = "normal"
            
            End If
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text3 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            
            If Option1(7).Value = True Then
                sWas = "klein"
            
            ElseIf Option1(8).Value = True Then
                sWas = "normal"
            
            ElseIf Option1(9).Value = True Then
                sWas = "groß"
            
            End If
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text4 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            
            
            If Option1(10).Value = True Then
                sWas = "straff"
            
            ElseIf Option1(11).Value = True Then
                sWas = "normal"
            
            ElseIf Option1(12).Value = True Then
                sWas = "nachlassend"
            ElseIf Option1(13).Value = True Then
                sWas = "erschlafft"
            End If
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text5 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            
            
            If Option1(14).Value = True Then
                sWas = "unempfindlich"
            
            ElseIf Option1(15).Value = True Then
                sWas = "ab und zu"
            
            ElseIf Option1(16).Value = True Then
                sWas = "sehr"
            
            End If
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text6 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            
            
            
            If Check1(0).Value = vbChecked Then
                sWas = "Ja"
            Else
                sWas = "Nein"
            End If
            
           
            
            sSQL = "Update KUNHAUTBILD_PRINT set text7 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
             If Check1(1).Value = vbChecked Then
                sWas = "Ja"
            Else
                sWas = "Nein"
            End If
            
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text8 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            If Check1(2).Value = vbChecked Then
                sWas = "Ja"
            Else
                sWas = "Nein"
            End If
            
            
            
            sSQL = "Update KUNHAUTBILD_PRINT set text9 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            If Check1(3).Value = vbChecked Then
                sWas = "Ja"
            Else
                sWas = "Nein"
            End If
            
            sSQL = "Update KUNHAUTBILD_PRINT set text10 = '" & sWas & "'"
            gdBase.Execute sSQL, dbFailOnError
           
            
            
            reportbildschirm "", "aWKL95b"
        
        
        Case "verwendete Produkte"
        
            loeschNEW "KUNPFLGH_PRINT", gdBase
            CreateTableT2 "KUNPFLGH_PRINT", gdBase
            
            sSQL = "Insert into KUNPFLGH_PRINT select "
            sSQL = sSQL & " KUNDNR  "
            sSQL = sSQL & ", DATUM  "
            sSQL = sSQL & ", REINIGUNG  "
            sSQL = sSQL & ", MASKEN  "
            sSQL = sSQL & ", EXTRAS  "
            sSQL = sSQL & ", TAGESPFLEG  "
            sSQL = sSQL & ", BERATUNG  "
            sSQL = sSQL & ", PROBEN  "
            sSQL = sSQL & ", KAEUFE  "
            sSQL = sSQL & ", TONIC  "
            sSQL = sSQL & ", PEELING  "
            sSQL = sSQL & ", MASSAGE  "
            sSQL = sSQL & ", NAGELLACK  "
            sSQL = sSQL & ", BEHANDLUNG  "
            sSQL = sSQL & ", cstr(Datum) + Behandlung as DATBEHANDLUNG "
            sSQL = sSQL & " from KUNPFLGH where kundnr = " & cKundnr & " order by datum desc"
            gdBase.Execute sSQL, dbFailOnError
            
            reportbildschirm "", "aWKL95a"
            
    End Select
    
    

    loeschNEW "KU_EINWIL", gdBase

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken_Kosmetik"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(Index).ForeColor = glLink
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label5_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub List3_GotFocus()
    List3.BackColor = glSelBack1
End Sub
Private Sub List3_LostFocus()
    List3.BackColor = vbWhite
End Sub
Private Sub List3_Click()
On Error GoTo LOKAL_ERROR

    Dim cDatum  As String
    Dim cBeh    As String
    Dim lWert   As Long
    
    cDatum = Left(List3.list(List3.ListIndex), 10)
    Label3(0).Caption = cDatum
    
    lWert = DateValue(cDatum)
    cDatum = Trim$(Str$(lWert))
    
    cBeh = Trim(Right(List3.list(List3.ListIndex), Len(List3.list(List3.ListIndex)) - 10))
    Label3(8).Caption = cBeh
    
    LesePflegeHistorieDetailWKL90 cDatum, cBeh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Option2(1).Value = True Then
        Text4(0).Visible = True
        Label1(42).Visible = True
    Else
        Text4(0).Visible = False
'        Text4(0).Text = ""
        Label1(42).Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kosmetik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = glSelBack1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
End Sub
Private Sub Text2_GotFocus(Index As Integer)
    Text2(Index).BackColor = glSelBack1
End Sub
Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).BackColor = vbWhite
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    Text3(Index).BackColor = glSelBack1
End Sub
Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).BackColor = vbWhite
End Sub
Private Sub Text4_GotFocus(Index As Integer)
    Text4(Index).BackColor = glSelBack1
End Sub
Private Sub Text4_LostFocus(Index As Integer)
    Text4(Index).BackColor = vbWhite
End Sub



