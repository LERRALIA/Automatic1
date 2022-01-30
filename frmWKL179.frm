VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL179 
   Caption         =   "Einstellungen Kassenbon"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL179.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check17 
      Caption         =   "Sonderpreis auf Kassenbon darstellen"
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
      Left            =   3720
      TabIndex        =   67
      Top             =   7560
      Width           =   3735
   End
   Begin VB.CheckBox Check16 
      Caption         =   "bei 'Gutscheinverkauf' Gültigkeit 4 Jahre drucken"
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
      Left            =   3720
      TabIndex        =   66
      Top             =   7200
      Width           =   4815
   End
   Begin VB.CheckBox Check15 
      Caption         =   "Terminerstellung"
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
      TabIndex        =   65
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CheckBox Check14 
      Caption         =   "bei Storno -> kurzer Rückgabebon"
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
      Left            =   3720
      TabIndex        =   64
      Top             =   6600
      Width           =   4335
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Unterschriftenfeld bei Gutscheinauszahlung"
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
      Left            =   3720
      TabIndex        =   63
      Top             =   6360
      Width           =   4335
   End
   Begin VB.CheckBox Check12 
      Caption         =   "3. Artikelzeile (Es bediente: ...) drucken"
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
      Left            =   3720
      TabIndex        =   62
      Top             =   6120
      Width           =   3615
   End
   Begin VB.ComboBox cboBONFONTNAME 
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
      ItemData        =   "frmWKL179.frx":0442
      Left            =   8760
      List            =   "frmWKL179.frx":0444
      Style           =   2  'Dropdown-Liste
      TabIndex        =   60
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   11160
      MaxLength       =   2
      TabIndex        =   58
      Text            =   "8"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   11160
      MaxLength       =   2
      TabIndex        =   56
      Text            =   "32"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Brutto/Nettoumsatz gesplittet drucken"
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
      Left            =   3720
      TabIndex        =   55
      Top             =   4080
      Width           =   3855
   End
   Begin VB.CheckBox Check10 
      Caption         =   "ohne Druckvorschau - direkt auf dem hinterlegten DINA4 Listendrucker drucken"
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
      Left            =   3720
      TabIndex        =   54
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CheckBox Check48 
      Caption         =   "Bon Logo drucken"
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
      Left            =   8760
      TabIndex        =   53
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame Frame16 
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
      Left            =   8760
      TabIndex        =   48
      Top             =   1320
      Width           =   2535
      Begin VB.CheckBox Check50 
         Caption         =   "Angebotslogo"
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
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check50 
         Caption         =   "Weihnachtslogo"
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
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox Check50 
         Caption         =   "Standardlogo"
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
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lbl6 
         Caption         =   "Welches Logo?"
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
         Index           =   63
         Left            =   120
         TabIndex        =   52
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
      Height          =   255
      Index           =   11
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   47
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check33 
      Caption         =   "Bon auf Warengruppe"
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
      Left            =   8760
      TabIndex        =   46
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CheckBox Check26 
      Caption         =   "Parkvorgänge drucken"
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
      Left            =   8760
      TabIndex        =   45
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CheckBox Check31 
      Caption         =   """Kopie"" bei 2. Bon"
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
      Left            =   3720
      TabIndex        =   44
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CheckBox Check53 
      Caption         =   """gilt als Rechnung"" bei Kassenbon als Lieferschein"
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
      Left            =   3720
      TabIndex        =   43
      Top             =   2880
      Width           =   4575
   End
   Begin VB.CheckBox Check9 
      Caption         =   "kein Bon 'Preisänderung wurde vorgenommen'"
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
      Left            =   3720
      TabIndex        =   42
      Top             =   5880
      Width           =   4095
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Ein- und Auszahlungen"
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
      TabIndex        =   41
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Nur die 1. Seite des Zollbelegs drucken"
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
      Left            =   3720
      TabIndex        =   40
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CheckBox Check6 
      Caption         =   "MWST im Zollbeleg drucken"
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
      Left            =   3720
      TabIndex        =   39
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Zahlung mit Gutschein"
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
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Retoure mit VK statt EK drucken"
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
      Left            =   3720
      TabIndex        =   37
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   35
      Text            =   "Parfümerieartikel"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CheckBox Check58 
      Caption         =   "kein Bon 'Gutscheinverkauf'"
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
      Left            =   3720
      TabIndex        =   34
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CheckBox Check54 
      Caption         =   "keine Grafik bei Filialtauschbon"
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
      Left            =   3720
      TabIndex        =   33
      Top             =   5400
      Width           =   3615
   End
   Begin VB.CheckBox Check81 
      Caption         =   "Unterschriftenbon bei ' Schublade öffnen'"
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
      Left            =   3720
      TabIndex        =   32
      Top             =   5160
      Width           =   4335
   End
   Begin VB.CheckBox Check88 
      Caption         =   "Rabatte auf Kassenbon drucken"
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
      Left            =   3720
      TabIndex        =   31
      Top             =   4920
      Width           =   3735
   End
   Begin VB.CheckBox Check85 
      Caption         =   "Rabatt Summe drucken"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      TabIndex        =   30
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Bonnummer spiegeln"
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
      Left            =   3720
      TabIndex        =   29
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Bonnummer unterdrücken"
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
      Left            =   3720
      TabIndex        =   28
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kassennummer unterdrücken"
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
      Left            =   3720
      TabIndex        =   27
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CheckBox Check27 
      Caption         =   "Kartenzahlung"
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
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CheckBox Check28 
      Caption         =   "Storno"
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
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CheckBox Check40 
      Caption         =   "Filialtausch"
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
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CheckBox Check106 
      Caption         =   "Kundenbestellung"
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
      TabIndex        =   22
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CheckBox Check77 
      Caption         =   "Barzahlung"
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
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CheckBox Check49 
      Caption         =   "Kreditverkauf"
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
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CheckBox Check21 
      Caption         =   "Bonusmeldung"
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
      TabIndex        =   19
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CheckBox Check35 
      Caption         =   "Kollegen Verkauf"
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
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CheckBox Check66 
      Caption         =   "Parken/Artikelverleih"
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
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CheckBox chkKundendaten 
      Caption         =   "Kundendaten im Bon"
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
      TabIndex        =   16
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame Frame17 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox Check94 
         Caption         =   "Kundennachname"
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
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox Check95 
         Caption         =   "Kundenvorname"
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
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Check96 
         Caption         =   "Telefonnr"
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
         TabIndex        =   13
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CheckBox Check97 
         Caption         =   "Mobilfunknr"
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
         TabIndex        =   12
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox Check98 
         Caption         =   "Firma"
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
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Check99 
         Caption         =   "Titel"
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
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox Check100 
         Caption         =   "Plz"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox Check101 
         Caption         =   "Ort"
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
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox Check102 
         Caption         =   "Strasse"
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
         TabIndex        =   7
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "*"
      Top             =   4200
      Width           =   495
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   7200
      Width           =   2055
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
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
      Index           =   3
      Left            =   8760
      TabIndex        =   61
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label lbl6 
      Caption         =   "Schriftgröße"
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
      Left            =   8760
      TabIndex        =   59
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label lbl6 
      Caption         =   "Zeilenlänge/Anzahl Zeichen"
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
      Index           =   1
      Left            =   8760
      TabIndex        =   57
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lbl6 
      Caption         =   "Standardtext für den Zollbeleg"
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
      Index           =   0
      Left            =   3600
      TabIndex        =   36
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lbl6 
      Caption         =   "2. Bon drucken bei"
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
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lbl6 
      Caption         =   "Sternchen - Trennlinien mit diesem Symbol drucken"
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
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
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
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Einstellungen Kassenbon"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmWKL179"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBONFONTNAME_Click()
On Error GoTo LOKAL_ERROR

    If cboBONFONTNAME.Text = "Standard" Then
        Text1(2).Visible = False
        lbl6(2).Visible = False
    Else
        Text1(2).Visible = True
        lbl6(2).Visible = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboBONFONTNAME_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check33_Click()
On Error GoTo LOKAL_ERROR

    If Check33.value = vbChecked Then
        Text1(11).Visible = True
    Else
        Text1(11).Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check33_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check48_Click()

On Error GoTo LOKAL_ERROR

    If Check48.value = vbChecked Then
        Frame16.Visible = True
    Else
        Frame16.Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check48_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case index
        Case 0
            Unload frmWKL179
        Case 1
            speicherKundenimBon
            leseKundenimBon
            speicher2Druck
            speicherBonLayout
            speicherSpiegel
            speicherBONWG
            speicherPark
            speicherPL
    End Select
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherLOGOdetails()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sStart As String
    
    loeschNEW "LOGOS", gdBase
    CreateTable "LOGOS", gdBase

    If Check50(1).value = vbChecked Then
        sSQL = "Insert into LOGOS (LOGO1) values (true)"
        gdBase.Execute sSQL, dbFailOnError
        gbLOGO1 = True
    ElseIf Check50(1).value = vbUnchecked Then
        sSQL = "Insert into LOGOS (LOGO1) values (False)"
        gdBase.Execute sSQL, dbFailOnError
        gbLOGO1 = False
    End If
    
    If Check50(0).value = vbChecked Then
    
        sSQL = "Update LOGOS Set LOGO2 = true "
        gdBase.Execute sSQL, dbFailOnError
        gbLOGO2 = True
    ElseIf Check50(0).value = vbUnchecked Then
        sSQL = "Update LOGOS Set LOGO2 = False "
        gdBase.Execute sSQL, dbFailOnError
        gbLOGO2 = False
    End If
    
    If Check50(2).value = vbChecked Then
        sSQL = "Update LOGOS Set LOGO3 = true "
        gdBase.Execute sSQL, dbFailOnError
        gbLOGO3 = True
    ElseIf Check50(2).value = vbUnchecked Then
        sSQL = "Update LOGOS Set LOGO3 = False "
        gdBase.Execute sSQL, dbFailOnError
        gbLOGO3 = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherLOGOdetails"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherPL()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    
    If Check48.value = vbChecked Then
        sSQL = "Update KASSEIN Set PL = true"
        gdBase.Execute sSQL, dbFailOnError
        gbPrintLOGO = True
    ElseIf Check48.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set PL = False"
        gdBase.Execute sSQL, dbFailOnError
        gbPrintLOGO = False
    End If
    
    If gbPrintLOGO Then
        speicherLOGOdetails
        iRet = fnLeseIniPrinterWKL00()
    Else
        loeschNEW "LOGOS", gdBase
        iRet = fnLeseIniPrinterWKL00()
    End If
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherPL"
        Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    
End Sub
Private Sub speicherPark()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "Park", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "Park", "BIT", gdApp
    End If
    
    If Check26.value = vbChecked Then
        sSQL = "Update WKEINSTE Set PARK = true"
        gdApp.Execute sSQL, dbFailOnError
        gbPark = True
        
    ElseIf Check26.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set PARK = False"
        gdApp.Execute sSQL, dbFailOnError
        gbPark = False
        
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "DritteArtikelzeile", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "DritteArtikelzeile", "BIT", gdApp
    End If
    
    If Check12.value = vbChecked Then
        sSQL = "Update WKEINSTE Set DritteArtikelzeile = true"
        gdApp.Execute sSQL, dbFailOnError
        gbDritteArtikelzeile = True
        
    ElseIf Check12.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set DritteArtikelzeile = False"
        gdApp.Execute sSQL, dbFailOnError
        gbDritteArtikelzeile = False
        
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherPark"
        Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    
End Sub
Private Sub speicherBONWG()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "BONWG", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "BONWG", "BIT", gdApp
    End If
    
    If SpalteInTabellegefundenNEW("WKEINSTE", "WGNR", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "WGNR", "BYTE", gdApp
    End If
    
    If Check33.value = vbChecked Then
    
        If Text1(11).Text = 0 Then Text1(11).Text = ""
        If IsNumeric(Text1(11).Text) Then
        
            
            
            sSQL = "Update WKEINSTE Set WGNR = " & Val(Text1(11).Text)
            gdApp.Execute sSQL, dbFailOnError
            gBYTEWGNR = Val(Text1(11).Text)
            
            gsWGart = ermartnrausWGN(CStr(gBYTEWGNR))
            gsWGBEZEICH = ermBezeichausWGN(gsWGart)
        
            sSQL = "Update WKEINSTE Set BONWG = true"
            gdApp.Execute sSQL, dbFailOnError
            gbBONWG = True
        
        Else
            
            sSQL = "Update WKEINSTE Set BONWG = false"
            gdApp.Execute sSQL, dbFailOnError
            gbBONWG = False
            
            sSQL = "Update WKEINSTE Set WGNR = 0"
            gdApp.Execute sSQL, dbFailOnError
            gBYTEWGNR = 0
            gsWGart = ""
            gsWGBEZEICH = ""
        End If
        
    ElseIf Check33.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BONWG = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBONWG = False
        
        sSQL = "Update WKEINSTE Set WGNR = 0"
        gdApp.Execute sSQL, dbFailOnError
        gBYTEWGNR = 0
        
        gsWGart = ""
        gsWGBEZEICH = ""
        
    End If
    
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "speicherBONWG"
        Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub speicher2Druck()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check31.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONKOPIE = true "
        gdApp.Execute sSQL, dbFailOnError
        gbBonkopie = True
    Else
        sSQL = "Update WKEINSTE Set BONKOPIE = false "
        gdApp.Execute sSQL, dbFailOnError
        gbBonkopie = False
    End If
    
    If Check53.value = vbChecked Then
        sSQL = "Update WKEINSTE Set GILTRE = true "
        gdApp.Execute sSQL, dbFailOnError
        gbGiltAlsRechnung = True
    Else
        sSQL = "Update WKEINSTE Set GILTRE = false "
        gdApp.Execute sSQL, dbFailOnError
        gbGiltAlsRechnung = False
    End If
    
    If Check77.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BARBON2 = true"
        gdApp.Execute sSQL, dbFailOnError
        gbBARBON2 = True
    ElseIf Check77.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BARBON2 = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBARBON2 = False
    End If
    
    If Check49.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONKR = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKR = True
    Else
        sSQL = "Update WKEINSTE Set BONKR = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKR = False
    End If
    
    If Check5.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONGUVK = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONGUVK = True
    Else
        sSQL = "Update WKEINSTE Set BONGUVK = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONGUVK = False
    End If
    
    If Check8.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONEA = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONEA = True
    Else
        sSQL = "Update WKEINSTE Set BONEA = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONEA = False
    End If
    
    If Check15.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONTERMIN = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONTermin = True
    Else
        sSQL = "Update WKEINSTE Set BONTERMIN = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONTermin = False
    End If
    
   
    
    If Check21.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONUSMESS = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONUSMESS = True
    Else
        sSQL = "Update WKEINSTE Set BONUSMESS = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONUSMESS = False
    End If
    
    
    
    If Check27.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONKA = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKA = True
    Else
        sSQL = "Update WKEINSTE Set BONKA = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKA = False
    End If
    
    If Check106.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONKB = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKB = True
    Else
        sSQL = "Update WKEINSTE Set BONKB = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKB = False
    End If
    
    If Check28.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONst = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONST = True
    Else
        sSQL = "Update WKEINSTE Set BONst = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONST = False
    End If
    
    If Check40.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONFI = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONFI = True
    Else
        sSQL = "Update WKEINSTE Set BONFI = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONFI = False
    End If
    
    If Check66.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONVerleih = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONVerleih = True
    Else
        sSQL = "Update WKEINSTE Set BONVerleih = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONVerleih = False
    End If
    
    
    
    If Check35.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONKOLLVK = true "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKOLLVK = True
    Else
        sSQL = "Update WKEINSTE Set BONKOLLVK = false "
        gdApp.Execute sSQL, dbFailOnError
        gb2BONKOLLVK = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher2Druck"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherBonLayout()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check1.value = vbChecked Then
        sSQL = "Update WKEINSTE Set BONNRUNTER = true"
        gdApp.Execute sSQL, dbFailOnError
        gbBONNRUNTER = True
    ElseIf Check1.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set BONNRUNTER = False"
        gdApp.Execute sSQL, dbFailOnError
        gbBONNRUNTER = False
    End If
    
    If Check2.value = vbChecked Then
        sSQL = "Update WKEINSTE Set KASSNRUNTER = true"
        gdApp.Execute sSQL, dbFailOnError
        gbKASSNRUNTER = True
    ElseIf Check2.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set KASSNRUNTER = False"
        gdApp.Execute sSQL, dbFailOnError
        gbKASSNRUNTER = False
    End If
    
    If CDbl(Text1(2).Text) > 0 Then
    
    Else
        Text1(2).Text = "8"
    End If
    
    
    sSQL = "Update WKEINSTE Set BONFONTSIZE = '" & CDbl(Text1(2).Text) & "'"
    gdApp.Execute sSQL, dbFailOnError
    glBONFONTSIZE = CDbl(Text1(2).Text)
    
    If Val(Text1(1).Text) > 0 Then
    
    Else
        Text1(1).Text = "32"
    End If
    
    
    sSQL = "Update WKEINSTE Set ANZZEICHENBON = " & Val(Text1(1).Text) & ""
    gdApp.Execute sSQL, dbFailOnError
    glZeichenAnzahlBon = Val(Text1(1).Text)
    
    
    sSQL = "Update WKEINSTE Set BONFONTNAME = '" & cboBONFONTNAME.Text & "'"
    gdApp.Execute sSQL, dbFailOnError
    gsBONFONTNAME = cboBONFONTNAME.Text
    
    
    
    
    
    sSQL = "Update WKEINSTE Set Sternzeich = '" & Text1(3).Text & "'"
    gdApp.Execute sSQL, dbFailOnError
    gsSTERNZEICH = Text1(3).Text
    
    sSQL = "Update KASSEIN Set ZOLLARTBEZ = '" & Text1(0).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    gsZOLLARTBEZ = Text1(0).Text
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBonLayout"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub speicherKundenimBon()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If chkKundendaten.value = vbChecked Then
        sSQL = "Update DBEINSTE Set KuImBoY = true "
        gdBase.Execute sSQL, dbFailOnError
        gbKUNDENA = True
    Else
        sSQL = "Update DBEINSTE Set KuImBoY = false "
        gdBase.Execute sSQL, dbFailOnError
        gbKUNDENA = False
    End If
    
    If gbKUNDENA = False Then
        loeschNEW "KUIBON", gdBase
        Exit Sub
    End If
    
    loeschNEW "KUIBON", gdBase
    CreateTableT2 "KUIBON", gdBase
    
    'Kundenname Nachname
    If Check94.value = vbChecked Then
        sSQL = "Insert into KUIBON (Name) values (true) "
        gdBase.Execute sSQL, dbFailOnError
    
        gbKUIBONname = True
    
    ElseIf Check94.value = vbUnchecked Then
    
        sSQL = "Insert into KUIBON (Name) values (false) "
        gdBase.Execute sSQL, dbFailOnError
        
        gbKUIBONname = False
    
    End If
    
    'Vorname
    
    If Check95.value = vbChecked Then

        sSQL = "Update KUIBON Set vorname = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONvorname = True

    ElseIf Check95.value = vbUnchecked Then

        sSQL = "Update KUIBON Set vorname = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONvorname = False

    End If
    
    'firma
    
    If Check98.value = vbChecked Then
        sSQL = "Update KUIBON Set firma = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONfirma = True

    ElseIf Check98.value = vbUnchecked Then

        sSQL = "Update KUIBON Set firma = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONfirma = False

    End If
    
    'titel
    If Check99.value = vbChecked Then
        sSQL = "Update KUIBON Set titel = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONtitel = True

    ElseIf Check99.value = vbUnchecked Then

        sSQL = "Update KUIBON Set titel = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONtitel = False

    End If
    
    'strasse
    If Check102.value = vbChecked Then
        sSQL = "Update KUIBON Set strasse = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONstrasse = True

    ElseIf Check102.value = vbUnchecked Then

        sSQL = "Update KUIBON Set strasse = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONstrasse = False

    End If
    
    'plz
    If Check100.value = vbChecked Then
        sSQL = "Update KUIBON Set plz = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONplz = True

    ElseIf Check100.value = vbUnchecked Then

        sSQL = "Update KUIBON Set plz = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONplz = False

    End If
    
    'ort
    If Check101.value = vbChecked Then
        sSQL = "Update KUIBON Set ort = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONort = True

    ElseIf Check101.value = vbUnchecked Then

        sSQL = "Update KUIBON Set ort = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONort = False

    End If
    
    'tel
    If Check96.value = vbChecked Then
        sSQL = "Update KUIBON Set tel = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONtel = True

    ElseIf Check96.value = vbUnchecked Then

        sSQL = "Update KUIBON Set tel = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONtel = False

    End If
    
    'mobil
    If Check97.value = vbChecked Then
        sSQL = "Update KUIBON Set mobil = true "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONmobil = True

    ElseIf Check97.value = vbUnchecked Then

        sSQL = "Update KUIBON Set mobil = false "
        gdBase.Execute sSQL, dbFailOnError

        gbKUIBONmobil = False

    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherKundenimBon"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Einstellungen Kassenbon auf. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    If gbKUNDENA = True Then
        chkKundendaten.value = vbChecked
        Frame17.Visible = True
        
        leseKundenimBon
        
        If gbKUIBONname = True Then
            Check94.value = vbChecked
        Else
            Check94.value = vbUnchecked
        End If
        
        If gbKUIBONvorname = True Then
            Check95.value = vbChecked
        Else
            Check95.value = vbUnchecked
        End If
        
        If gbKUIBONfirma = True Then
            Check98.value = vbChecked
        Else
            Check98.value = vbUnchecked
        End If
        
        If gbKUIBONtitel = True Then
            Check99.value = vbChecked
        Else
            Check99.value = vbUnchecked
        End If
        
        If gbKUIBONstrasse = True Then
            Check102.value = vbChecked
        Else
            Check102.value = vbUnchecked
        End If
        
        If gbKUIBONplz = True Then
            Check100.value = vbChecked
        Else
            Check100.value = vbUnchecked
        End If
        
        If gbKUIBONort = True Then
            Check101.value = vbChecked
        Else
            Check101.value = vbUnchecked
        End If
        
        If gbKUIBONtel = True Then
            Check96.value = vbChecked
        Else
            Check96.value = vbUnchecked
        End If
        
        If gbKUIBONmobil = True Then
            Check97.value = vbChecked
        Else
            Check97.value = vbUnchecked
        End If
        
    Else
        Frame17.Visible = False
        chkKundendaten.value = vbUnchecked
    End If
    
    'Bonlayout
    If gbBONNRUNTER Then
        Check1.value = vbChecked
    Else
        Check1.value = vbUnchecked
    End If
    
    If gbKASSNRUNTER Then
        Check2.value = vbChecked
    Else
        Check2.value = vbUnchecked
    End If
    
    Text1(3).Text = gsSTERNZEICH
    Text1(1).Text = glZeichenAnzahlBon
    
    cboBONFONTNAME.AddItem "Standard"
    cboBONFONTNAME.AddItem "Lucida Console"
    cboBONFONTNAME.AddItem "Courier New"
    
    cboBONFONTNAME.Text = gsBONFONTNAME
    
    Text1(2).Text = gdBONFONTSIZE
    
    Text1(0).Text = gsZOLLARTBEZ
    
    
    
    If cboBONFONTNAME.Text = "Standard" Then
        Text1(2).Visible = False
        lbl6(2).Visible = False
    Else
        Text1(2).Visible = True
        lbl6(2).Visible = True
    End If
    
    
    
    '2.Bon
    If gb2BONKOLLVK Then
        Check35.value = vbChecked
    Else
        Check35.value = vbUnchecked
    End If
    
    If gb2BONFI Then
        Check40.value = vbChecked
    Else
        Check40.value = vbUnchecked
    End If
    
    If gb2BONVerleih Then
        Check66.value = vbChecked
    Else
        Check66.value = vbUnchecked
    End If
    
    If gb2BONST Then
        Check28.value = vbChecked
    Else
        Check28.value = vbUnchecked
    End If
    
    If gb2BONKB Then
        Check106.value = vbChecked
    Else
        Check106.value = vbUnchecked
    End If
    
    If gbBARBON2 = True Then
        Check77.value = vbChecked
    Else
        Check77.value = vbUnchecked
    End If
    
    If gb2BONKA Then
        Check27.value = vbChecked
    Else
        Check27.value = vbUnchecked
    End If
    
    If gb2BONKR Then
        Check49.value = vbChecked
    Else
        Check49.value = vbUnchecked
    End If
    
    If gb2BONGUVK Then
        Check5.value = vbChecked
    Else
        Check5.value = vbUnchecked
    End If
    
    If gb2BONEA Then
        Check8.value = vbChecked
    Else
        Check8.value = vbUnchecked
    End If
    
    If gb2BONTermin Then
        Check15.value = vbChecked
    Else
        Check15.value = vbUnchecked
    End If
    
    If gb2BONUSMESS Then
        Check21.value = vbChecked
    Else
        Check21.value = vbUnchecked
    End If
    
    
    
    If gbSPIEGEL = True Then
        Check3.value = vbChecked
    Else
        Check3.value = vbUnchecked
    End If
    
    If gbZOLLmMWST = True Then
        Check6.value = vbChecked
    Else
        Check6.value = vbUnchecked
    End If
    
    If gbZOLLonlyFirstPage = True Then
        Check7.value = vbChecked
    Else
        Check7.value = vbUnchecked
    End If
    
    If gbZOLLPrintDirekt = True Then
        Check10.value = vbChecked
    Else
        Check10.value = vbUnchecked
    End If
    
    If gbRETVK = True Then
        Check4.value = vbChecked
    Else
        Check4.value = vbUnchecked
    End If
    
    If gbSparsatz = True Then
        Check85.value = vbChecked
    Else
        Check85.value = vbUnchecked
    End If
    
    If gbNoBonGu = True Then
        Check58.value = vbChecked
    Else
        Check58.value = vbUnchecked
    End If
    
    If gbBonGu2J = True Then
        Check16.value = vbChecked
    Else
        Check16.value = vbUnchecked
    End If
    
    If gbNoBonPÄ = True Then
        Check9.value = vbChecked
    Else
        Check9.value = vbUnchecked
    End If
    
    If gbNoGrafik = True Then
        Check54.value = vbChecked
    Else
        Check54.value = vbUnchecked
    End If
    
    If gbKurzerStorni Then
        Check14.value = vbChecked
    Else
        Check14.value = vbUnchecked
    End If
    
    If gbGUTSCHBARAUSZAHLUNGMITUNTER Then
        Check13.value = vbChecked
    Else
        Check13.value = vbUnchecked
    End If
    
    If gbSCHUBMB Then
        Check81.value = vbChecked
    Else
        Check81.value = vbUnchecked
    End If
    
    If gbRabatt = True Then
        Check88.value = vbChecked
    Else
        Check88.value = vbUnchecked
    End If
    
    If gbSonderPreisDarstellen = True Then
        Check17.value = vbChecked
    Else
        Check17.value = vbUnchecked
    End If
    
    If gbBonkopie Then
        Check31.value = vbChecked
    Else
        Check31.value = vbUnchecked
    End If
    
    If gbGiltAlsRechnung Then
        Check53.value = vbChecked
    Else
        Check53.value = vbUnchecked
    End If
    
    If gbDritteArtikelzeile = True Then
        Check12.value = vbChecked
    Else
        Check12.value = vbUnchecked
    End If
    
    If gbPark = True Then
        Check26.value = vbChecked
    Else
        Check26.value = vbUnchecked
    End If
    
    If gbBONWG = True Then
        Check33.value = vbChecked
        Text1(11).Text = gBYTEWGNR
    Else
        Check33.value = vbUnchecked
        Text1(11).Text = ""
    End If
    
    If gbPrintLOGO = True Then
        Frame16.Visible = True
        leseLOGOS
        Check48.value = vbChecked
        
        
        
        If gbLOGO1 = True Then
            Check50(1).value = vbChecked
        Else
            Check50(1).value = vbUnchecked
        End If
        
        If gbLOGO2 = True Then
            Check50(0).value = vbChecked
        Else
            Check50(0).value = vbUnchecked
        End If
        
        If gbLOGO3 = True Then
            Check50(2).value = vbChecked
        Else
            Check50(2).value = vbUnchecked
        End If
        
        
    Else
        Frame16.Visible = False
        gbLOGO1 = False
        gbLOGO2 = False
        gbLOGO3 = False
        Check48.value = vbUnchecked
    End If
    
    If gbMitMwstAnteile = True Then
        Check11.value = vbChecked
    Else
        Check11.value = vbUnchecked
    End If
            
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub chkKundendaten_Click()
On Error GoTo LOKAL_ERROR

    If chkKundendaten.value = vbChecked Then
        Frame17.Visible = True
    Else
        Frame17.Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "chkKundendaten_Click"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherSpiegel()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If SpalteInTabellegefundenNEW("WKEINSTE", "SPIEGEL", gdApp) = False Then
        SpalteAnfuegenNEW "WKEINSTE", "SPIEGEL", "BIT", gdApp
    End If

    If Check3.value = vbChecked Then
        sSQL = "Update WKEINSTE Set SPIEGEL = true"
        gdApp.Execute sSQL, dbFailOnError
        gbSPIEGEL = True

    ElseIf Check3.value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set SPIEGEL = False"
        gdApp.Execute sSQL, dbFailOnErro
        gbSPIEGEL = False
    End If
    
    
    
    
    If Check11.value = vbChecked Then
        sSQL = "Update KASSEIN Set MitMwstAnteile = true"
        gdBase.Execute sSQL, dbFailOnError
        gbMitMwstAnteile = True
    ElseIf Check11.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set MitMwstAnteile = False"
        gdBase.Execute sSQL, dbFailOnError
        gbMitMwstAnteile = False
    End If
    
    
    
    
    If Check4.value = vbChecked Then
        sSQL = "Update KASSEIN Set RETVK = true"
        gdBase.Execute sSQL, dbFailOnError
        gbRETVK = True
    ElseIf Check4.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set RETVK = False"
        gdBase.Execute sSQL, dbFailOnError
        gbRETVK = False
    End If
    
    If Check6.value = vbChecked Then
        sSQL = "Update KASSEIN Set ZOLLmMWST = true"
        gdBase.Execute sSQL, dbFailOnError
        gbZOLLmMWST = True
    ElseIf Check6.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set ZOLLmMWST = False"
        gdBase.Execute sSQL, dbFailOnError
        gbZOLLmMWST = False
    End If
    
    If Check7.value = vbChecked Then
        sSQL = "Update KASSEIN Set ZOLLonlyfirstpage = true"
        gdBase.Execute sSQL, dbFailOnError
        gbZOLLonlyFirstPage = True
    ElseIf Check7.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set ZOLLonlyfirstpage = False"
        gdBase.Execute sSQL, dbFailOnError
        gbZOLLonlyFirstPage = False
    End If
    
    If Check10.value = vbChecked Then
        sSQL = "Update KASSEIN Set ZOLLPrintDirekt = true"
        gdBase.Execute sSQL, dbFailOnError
        gbZOLLPrintDirekt = True
    ElseIf Check10.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set ZOLLPrintDirekt = False"
        gdBase.Execute sSQL, dbFailOnError
        gbZOLLPrintDirekt = False
    End If
    
    
    
    
    If Check85.value = vbChecked Then
        sSQL = "Update KASSEIN Set SparSatz = true"
        gdBase.Execute sSQL, dbFailOnError
        gbSparsatz = True
    ElseIf Check85.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set SparSatz = False"
        gdBase.Execute sSQL, dbFailOnError
        gbSparsatz = False
    End If
    
    If Check54.value = vbChecked Then
        sSQL = "Update KASSEIN Set nografik = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNoGrafik = True
    ElseIf Check54.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set nografik = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNoGrafik = False
    End If
    
    If Check58.value = vbChecked Then
        sSQL = "Update KASSEIN Set nobongu = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNoBonGu = True
    ElseIf Check58.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set nobongu = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNoBonGu = False
    End If
    
    If Check16.value = vbChecked Then
        sSQL = "Update KASSEIN Set BonGu2J = True"
        gdBase.Execute sSQL, dbFailOnError
        gbBonGu2J = True
    ElseIf Check16.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set BonGu2J = False"
        gdBase.Execute sSQL, dbFailOnError
        gbBonGu2J = False
    End If
    
    If Check9.value = vbChecked Then
        sSQL = "Update KASSEIN Set nobonPAE = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNoBonPÄ = True
    ElseIf Check9.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set nobonPAE = False"
        gdBase.Execute sSQL, dbFailOnError
        gbNoBonPÄ = False
    End If
    
    
    
    If Check88.value = vbChecked Then
        sSQL = "Update KASSEIN Set Rabatt = true"
        gdBase.Execute sSQL, dbFailOnError
        gbRabatt = True
    ElseIf Check88.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set Rabatt = False"
        gdBase.Execute sSQL, dbFailOnError
        gbRabatt = False
    End If
    
    
    If Check17.value = vbChecked Then
        sSQL = "Update KASSEIN Set SonderPreisDarstellen = true"
        gdBase.Execute sSQL, dbFailOnError
        gbSonderPreisDarstellen = True
    ElseIf Check17.value = vbUnchecked Then
        sSQL = "Update KASSEIN Set SonderPreisDarstellen = False"
        gdBase.Execute sSQL, dbFailOnError
        gbSonderPreisDarstellen = False
    End If
    
    
    
    
    If Check81.value = vbChecked Then
        sSQL = "Update DBEINSTE Set SCHUBMB = True "
        gdBase.Execute sSQL, dbFailOnError
        gbSCHUBMB = True
    Else
        sSQL = "Update DBEINSTE Set SCHUBMB = False "
        gdBase.Execute sSQL, dbFailOnError
        gbSCHUBMB = False
    End If
    
    If Check13.value = vbChecked Then
        sSQL = "Update DBEINSTE Set GUTSCHBARAUSZAHLUNGMITUNTER = True "
        gdBase.Execute sSQL, dbFailOnError
        gbGUTSCHBARAUSZAHLUNGMITUNTER = True
    ElseIf Check13.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set GUTSCHBARAUSZAHLUNGMITUNTER = False "
        gdBase.Execute sSQL, dbFailOnError
        gbGUTSCHBARAUSZAHLUNGMITUNTER = False
    End If
    
    If Check14.value = vbChecked Then
        sSQL = "Update DBEINSTE Set KurzerStorni = True "
        gdBase.Execute sSQL, dbFailOnError
        gbKurzerStorni = True
    ElseIf Check14.value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set KurzerStorni = False "
        gdBase.Execute sSQL, dbFailOnError
        gbKurzerStorni = False
    End If
    
    
    
            

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherSpiegel"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Einstellungen Kassenbon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



