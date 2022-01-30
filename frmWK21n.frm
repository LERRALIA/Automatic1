VERSION 5.00
Begin VB.Form frmWK21n 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   8910
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'Kein
      Height          =   2415
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   5355
         TabIndex        =   78
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label labglo 
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
         Left            =   120
         TabIndex        =   83
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label lab 
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
         Left            =   120
         TabIndex        =   82
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Datenbank kopieren"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   81
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   2175
         Index           =   34
         Left            =   6000
         TabIndex        =   80
         Top             =   120
         Width           =   5655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   5880
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5880
         X2              =   5880
         Y1              =   120
         Y2              =   2280
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FF00&
      Height          =   6975
      Left            =   5400
      TabIndex        =   74
      Top             =   4680
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   7
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   11535
      End
      Begin VB.Label Label90 
         Caption         =   "l vk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8640
         TabIndex        =   94
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label90 
         Caption         =   "l vk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8640
         TabIndex        =   93
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label90 
         Caption         =   "l vk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8640
         TabIndex        =   92
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label80 
         Caption         =   "Bestand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   91
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label80 
         Caption         =   "Bestand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6360
         TabIndex        =   90
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label80 
         Caption         =   "Bestand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6360
         TabIndex        =   89
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label70 
         Caption         =   "bezeich"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   88
         Top             =   2520
         Width           =   5055
      End
      Begin VB.Label Label70 
         Caption         =   "bezeich"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   87
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label Label70 
         Caption         =   "bezeich"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   86
         Top             =   1560
         Width           =   5055
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Rechts
         Caption         =   "artnr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   85
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Rechts
         Caption         =   "artnr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   84
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Rechts
         Caption         =   "artnr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   29
      Top             =   6840
      Width           =   3615
      Begin VB.Label Label2 
         Caption         =   "Top 10 der Verkaufsmenge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   11
         Left            =   9120
         TabIndex        =   73
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   6
         Left            =   5280
         TabIndex        =   72
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   5
         Left            =   4440
         TabIndex        =   71
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   4
         Left            =   3600
         TabIndex        =   70
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   3
         Left            =   2760
         TabIndex        =   69
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   2
         Left            =   1920
         TabIndex        =   68
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   1
         Left            =   1080
         TabIndex        =   67
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   7
         Left            =   6120
         TabIndex        =   65
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   8
         Left            =   6960
         TabIndex        =   64
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   9
         Left            =   7800
         TabIndex        =   63
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   5430
         TabIndex        =   62
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   4590
         TabIndex        =   61
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   3750
         TabIndex        =   60
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   2910
         TabIndex        =   59
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   2070
         TabIndex        =   58
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   57
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   56
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   6270
         TabIndex        =   55
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   7110
         TabIndex        =   54
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   9
         Left            =   7950
         TabIndex        =   53
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   0
         Left            =   240
         TabIndex        =   52
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   1080
         TabIndex        =   51
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   1920
         TabIndex        =   50
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   2760
         TabIndex        =   49
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   3600
         TabIndex        =   48
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   4440
         TabIndex        =   47
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   5280
         TabIndex        =   46
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   6120
         TabIndex        =   45
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   8
         Left            =   6960
         TabIndex        =   44
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   9
         Left            =   7800
         TabIndex        =   43
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   5
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1080
         TabIndex        =   38
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   1920
         TabIndex        =   37
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2760
         TabIndex        =   36
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3600
         TabIndex        =   35
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   4440
         TabIndex        =   34
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   5280
         TabIndex        =   33
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   6120
         TabIndex        =   32
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   6960
         TabIndex        =   31
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   7800
         TabIndex        =   30
         Top             =   4320
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   11415
      Begin VB.Label Label2 
         Caption         =   "Verkaufsmenge pro Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   22
         Left            =   8760
         TabIndex        =   28
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   2
         Left            =   1920
         TabIndex        =   27
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   1
         Left            =   1080
         TabIndex        =   26
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   3
         Left            =   6120
         TabIndex        =   24
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   4
         Left            =   6960
         TabIndex        =   23
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   5
         Left            =   7800
         TabIndex        =   22
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   2070
         TabIndex        =   21
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   20
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   19
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   6270
         TabIndex        =   18
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   7110
         TabIndex        =   17
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   7950
         TabIndex        =   16
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   6960
         TabIndex        =   11
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Left            =   7800
         TabIndex        =   10
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Label8"
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Label8"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   3
         Left            =   6120
         TabIndex        =   3
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   4
         Left            =   6960
         TabIndex        =   2
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         Caption         =   "Label2"
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
         Index           =   5
         Left            =   7800
         TabIndex        =   1
         Top             =   4320
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"frmWK21n.frx":0000
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
      Index           =   10
      Left            =   120
      TabIndex        =   96
      Top             =   2520
      Width           =   11535
   End
End
Attribute VB_Name = "frmWK21n"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim gitop As Integer
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    Unload Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    PositionierenWK21n
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Label1(0)
    
    gitop = Shape1(0).Top
    
    Select Case giCopyMod
    
        Case 1
            Label1(0).Caption = "Datenbank wird kopiert..."
            Label1(0).Refresh
        Case 2
            Label1(0).Caption = "Datenbank wird gesichert..."
            Label1(0).Refresh
    End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Activate()
On Error GoTo LOKAL_ERROR

    Dim cVon As String
    Dim cBis As String
    Dim iFil As Integer
    
    
    
    If gbKopOhneAuswertung = False Then

        Label1(5).Caption = "Die besten Lieferanten - Verkaufszahlen im Zeitraum"
        
        Label1(1).Caption = "Die besten Bediener(Verkaufsschnitt) - Verkaufszahlen im Zeitraum"
    
        cVon = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
        cVon = Format$(cVon, "DD.MM.YY")
        cBis = DateValue(Now) - 1
        
        iFil = 0  'alle fils
        TopUmsatz cVon, cBis, iFil, "preis" '"menge"
    Else
        Frame2.Visible = False
    End If
    
    Frame3.Visible = False
    Me.Refresh
    
    starte giCopyMod
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Activate"
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TotalaltNoVerkauft()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    Dim lMonat As Long
    Dim lJahr As Long
    Dim i As Integer
    
    For i = 0 To 2
        Label60(i).Caption = ""
        Label70(i).Caption = ""
        Label80(i).Caption = ""
        Label90(i).Caption = ""
    Next i
    i = 0
    lHeute = CLng(DateValue(Now))
    
    loeschNEW "Btemp55", gdBase
    
    lMonat = Month(Now)
    lJahr = Year(Now)

    sSQL = "select * into Btemp55 from Bestaend "
    sSQL = sSQL & " where Jahr = " & lJahr
    sSQL = sSQL & " and monat = " & lMonat
    sSQL = sSQL & " and Bestand > 0  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART551", gdBase
    CreateTable "ART551", gdBase

    sSQL = " Insert into ART551 select  ARTNR,BESTAND from Btemp55 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART551 set lastvk = '01.01.2000'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART551 inner join Artikel  on ART551.ARTNR = artikel.ARTNR"
    sSQL = sSQL & " set ART551.aufdat = artikel.aufdat "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ART551 "
    sSQL = sSQL & " where aufdat is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Delete from ART551 "
    sSQL = sSQL & " where aufdat >  " & CLng(DateValue(Now)) - 730
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Delete from ART551 "
    sSQL = sSQL & " where bestand < 1  "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("ART551")


    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            
            
            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                ldifferenz = 0
                rsrs.Edit
                datLVK = ErmlzVK(cART)
                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk


                Select Case ldifferenz
                    Case Is > 365
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 12 Monaten"
                        End If
                        
                        Set rsArt = gdBase.OpenRecordset("Select Bestand from Artikel where artnr = " & cART)
                        If Not rsArt.EOF Then
                        
                            If Not IsNull(rsArt!BESTAND) Then
                                If Val(rsArt!BESTAND) > 0 Then
                                    rsrs!BESTAND = rsArt!BESTAND
                                    i = i + 1
                                End If
                            End If
                        
                        End If
                        rsArt.Close: Set rsArt = Nothing

                    Case Else
                        ctmp = ""
                End Select

                rsrs!Monat = ctmp
                rsrs!lastvk = datLVK

                rsrs.Update

            End If
            
            If i = 3 Then
                rsrs.MoveLast
            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing

 
    sSQL = "Delete from art551 where Monat = '' "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Delete from art551 where Monat is null "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update art551 inner join Artikel on art551.artnr = artikel.artnr "
    sSQL = sSQL & " Set art551.bezeich = artikel.bezeich "
    gdBase.Execute sSQL, dbFailOnError
    
    i = 0
    Set rsrs = gdBase.OpenRecordset("ART551")
    If Not rsrs.EOF Then

        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                Label60(i).Caption = rsrs!artnr
            Else
                Label60(i).Caption = ""
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                Label70(i).Caption = rsrs!BEZEICH
            Else
                Label70(i).Caption = ""
            End If
            
            If Not IsNull(rsrs!BESTAND) Then
                Label80(i).Caption = "Bestand: " & rsrs!BESTAND
            Else
                Label80(i).Caption = ""
            End If
            
            If Not IsNull(rsrs!lastvk) Then
                If Trim(rsrs!lastvk) = "00:00:00" Then
                    Label90(i).Caption = "letzter Verkauf: noch nie"
                Else
                    Label90(i).Caption = "letzter Verkauf: " & rsrs!lastvk
                End If
            Else
                Label90(i).Caption = ""
            End If
            i = i + 1
            If i = 3 Then
                rsrs.MoveLast
            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "ART551", gdBase
    loeschNEW "Btemp55", gdBase
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TotalaltNoVerkauft"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub TopUmsatz(cVon As String, cBis As String, iFil, cWert As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cvon1 As String
    Dim cbis1 As String
    
    loeschNEW "topi" & srechnertab, gdBase
    loeschNEW "TOPUMSATZ", gdBase
    CreateTable "TOPI" & srechnertab, gdBase
    
    cvon1 = cVon
    cbis1 = cBis
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    Screen.MousePointer = 11
    
    If cWert = "rertrag" Then

        loeschNEW "Ert", gdBase
        CreateTable "ERT", gdBase
        
        sSQL = "Insert into Ert Select "
        
        sSQL = sSQL & "  preis "
        sSQL = sSQL & " , menge "
        sSQL = sSQL & " , linr "
        sSQL = sSQL & " , ekpr "
        sSQL = sSQL & " , mwst "
       
        sSQL = sSQL & " from Kassjour "
        sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
        sSQL = sSQL & " and UMS_OK = 'J' "
        
        If iFil = 0 Then
        
        Else
            sSQL = sSQL & " and filiale = " & iFil
        End If
        
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Ert set rertrag = ((Preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * menge) "
        sSQL = sSQL & " where mwst = 'V' "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Ert set rertrag = ((Preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * menge) "
        sSQL = sSQL & " where mwst = 'E' "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Ert set rertrag = ((Preis * 100)/(100 + " & gdMWStO & " )) - (EKPR * menge) "
        sSQL = sSQL & " where mwst = 'O' "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Select sum(" & cWert & ") as maxi ,linr into TOPUMSATZ "
        sSQL = sSQL & " from Ert group by linr "
        gdBase.Execute sSQL, dbFailOnError
    
        loeschNEW "Ert", gdBase

    Else
        sSQL = "Select sum(" & cWert & ") as maxi ,linr into TOPUMSATZ "
        sSQL = sSQL & " from Kassjour "
        sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
        sSQL = sSQL & " and UMS_OK = 'J' "
        If iFil = 0 Then

        Else
            sSQL = sSQL & " and filiale = " & iFil
        End If

        sSQL = sSQL & " group by linr "
        gdBase.Execute sSQL, dbFailOnError

    End If

    sSQL = "insert into topi" & srechnertab & " SELECT TOP 10 LINR, maxi "
    sSQL = sSQL & " from TOPUMSATZ order by maxi desc"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topi" & srechnertab & " inner join lisrt on topi" & srechnertab & ".linr = lisrt.linr "
    sSQL = sSQL & " SET topi" & srechnertab & ".liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topi" & srechnertab
    sSQL = sSQL & " SET topi" & srechnertab & ".von = '" & cvon1 & "'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topi" & srechnertab
    sSQL = sSQL & " SET topi" & srechnertab & ".bis = '" & cbis1 & "'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topi" & srechnertab
    sSQL = sSQL & " SET topi" & srechnertab & ".fil  = 'alle Filialen'"
    gdBase.Execute sSQL, dbFailOnError
    
    Label1(3).Caption = cvon1
    Label1(4).Caption = cbis1
    
    Zeigegrafik "topi" & srechnertab, cWert

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TopUmsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TopArtikelUmsatz(cVon As String, cBis As String, iFil, cWert As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cvon1 As String
    Dim cbis1 As String
    
    loeschNEW "topiA", gdBase
    loeschNEW "TOPARTUMSATZ", gdBase
    CreateTable "TOPIA", gdBase
    
    cvon1 = cVon
    cbis1 = cBis
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    Screen.MousePointer = 11
    
    sSQL = "Select sum(" & cWert & ") as maxi ,artnr,bezeich into TOPARTUMSATZ "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & " and UMS_OK = 'J' "
    If iFil = 0 Then

    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If

    sSQL = sSQL & " group by artnr,bezeich "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "insert into topiA SELECT TOP 10 artnr,bezeich, maxi "
    sSQL = sSQL & " from TOPARTUMSATZ order by maxi desc"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topiA "
    sSQL = sSQL & " SET TOPiA.von = '" & cvon1 & "'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topiA "
    sSQL = sSQL & " SET TOPiA.bis = '" & cbis1 & "'"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update topiA "
    sSQL = sSQL & " SET TOPiA.fil  = 'alle Filialen'"
    gdBase.Execute sSQL, dbFailOnError
    
    Label1(3).Caption = cvon1
    Label1(4).Caption = cbis1
    
    Zeigegrafik "TOPIA", cWert

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TopartikelUmsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
    
End Sub
Private Sub TopBediener(cVon As String, cBis As String, iFil)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cvon1 As String
    Dim cbis1 As String
    
    loeschNEW "TB1" & srechnertab, gdBase
    loeschNEW "TOPBEDIENER" & srechnertab, gdBase
    
    sSQL = "Create Table TOPBEDIENER" & srechnertab
    sSQL = sSQL & "( BEDIENER INTEGER"
    sSQL = sSQL & ", SMENGE LONG"
    sSQL = sSQL & ", BONANZ LONG"
    sSQL = sSQL & ", KUCUT single"
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError

    cvon1 = cVon
    cbis1 = cBis
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    Screen.MousePointer = 11
    
    loeschNEW "AAT" & srechnertab, gdBase
    
    sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN , bediener, sum(Menge) as AMenge into AAT" & srechnertab
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & " and UMS_OK = 'J' and not belegnr is null"
    If iFil = 0 Then

    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    sSQL = sSQL & " group by adate,bediener,BELEGNR"
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Select sum(AMenge) as SMENGE ,BEDIENER,count(ANZKUNDEN) as belegnr into TB1" & srechnertab
    sSQL = sSQL & " from AAT" & srechnertab
    sSQL = sSQL & " group by bediener"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into TOPBEDIENER" & srechnertab & " Select SMENGE ,BEDIENER, belegnr as bonanz from TB1" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update TOPBEDIENER" & srechnertab & " set KUCUT = sMenge/bonanz "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "TH3" & srechnertab, gdBase
    sSQL = "Create Table TH3" & srechnertab
    sSQL = sSQL & "("
    sSQL = sSQL & " KUCUT Double"
    sSQL = sSQL & ", BEDIENER LONG"
    sSQL = sSQL & ", BEDNAME TEXT(35) "
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into TH3" & srechnertab & " SELECT TOP 3 Bediener, KUCUT "
    sSQL = sSQL & " from TOPBEDIENER" & srechnertab & " order by KUCUT desc"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into TH3" & srechnertab & " SELECT TOP 3 Bediener, KUCUT "
    sSQL = sSQL & " from TOPBEDIENER" & srechnertab & " order by KUCUT asc"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update TH3" & srechnertab & " inner join Bedname on TH3" & srechnertab & ".bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET TH3" & srechnertab & ".BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    Label1(6).Caption = cvon1
    Label1(2).Caption = cbis1
    
    ZeigegrafikB
    
    loeschNEW "AAT" & srechnertab, gdBase
    loeschNEW "TB1" & srechnertab, gdBase
    loeschNEW "TOPBEDIENER" & srechnertab, gdBase
    loeschNEW "TH3" & srechnertab, gdBase

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TopBediener"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
  
    
End Sub
Private Sub Zeigegrafik(cTab As String, cWert As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim siUms(9)        As Single
    Dim sLiefBez(9)     As String
    Dim rsUms           As Recordset
    Dim rsUmsVJ         As Recordset
    Dim lLinr(9)        As Long
    Dim i               As Integer
    
    
    Dim dbuffer     As Double
    Dim dMax        As Double
    Dim sDatum      As String
    Dim iJahr       As Integer
    Dim lHeuteVJ    As Long
    Dim cheuteVJ    As String
    Dim siWert      As Single
    Dim Wert1       As Integer
    
    
    
    
    
    For i = 0 To 9
        Shape1(i).Visible = False
    Next i
    
    sSQL = " Select * from " & cTab & " order by maxi desc "
    
                           
    Set rsUms = gdBase.OpenRecordset(sSQL)
    If Not rsUms.EOF Then
        i = 0
        rsUms.MoveFirst
    
        Do While Not rsUms.EOF
        If i > 9 Then Exit Do
        
            If Not IsNull(rsUms!maxi) Then
                siUms(i) = rsUms!maxi
            Else
                siUms(i) = 0
            End If
            
            If cTab = "TOPIA" Then
                If Not IsNull(rsUms!artnr) Then
                    lLinr(i) = rsUms!artnr
                Else
                    lLinr(i) = "0"
                End If
                
                If Not IsNull(rsUms!BEZEICH) Then
                    sLiefBez(i) = rsUms!BEZEICH
                Else
                    sLiefBez(i) = ""
                End If
            
            
            Else
            
                If Not IsNull(rsUms!linr) Then
                    lLinr(i) = rsUms!linr
                Else
                    lLinr(i) = "0"
                End If
                
                If Not IsNull(rsUms!LIEFBEZ) Then
                    sLiefBez(i) = rsUms!LIEFBEZ
                Else
                    sLiefBez(i) = ""
                End If
                
            End If
            
            i = i + 1
        rsUms.MoveNext
        Loop
    End If
    rsUms.Close
    
    For i = 0 To 9
        Label3(i).Caption = sLiefBez(i)
        Label3(i).Refresh
    
        Label2(i).Caption = lLinr(i)
        Label2(i).Refresh
        Label2(i).ToolTipText = sLiefBez(i)
        Label33(i).ToolTipText = sLiefBez(i)
        
        
    Next i
    
    dbuffer = 0
    dMax = 0
    
    For i = 0 To 9
        dbuffer = siUms(i)
        If dbuffer > dMax Then
            dMax = dbuffer
        End If
    Next i
    dMax = IIf(dMax = 0, 1, dMax)
    
    Dim screenfak As Long
    
    Select Case Screen.Height
        Case Is > 15000
            screenfak = 5000
        Case Is > 12000
            screenfak = 3800
        Case Is > 11000
            screenfak = 3500
        Case Is > 10000
            screenfak = 3000
        Case Is > 8000
            screenfak = 2800
        Case Else
            screenfak = 2000
    End Select
    
    For i = 0 To 9
        
        Shape1(0).Top = gitop
        Shape1(i).Height = 15
        Label33(i).Top = gitop - 250

    Next i

        
    
    For i = 0 To 9
        
        If siUms(i) > 0 Then
            Shape1(i).Height = (screenfak / dMax) * IIf(siUms(i) < 0, 0, siUms(i))
            Shape1(i).Top = gitop - ((screenfak / dMax) * siUms(i))
           
            Label33(i).Top = Shape1(i).Top - 250
            If cWert = "menge" Then
                Label33(i).Caption = Format$(siUms(i), "########")
            Else
                Label33(i).Caption = Format$(siUms(i), "###,##0.00")
            End If

            Label33(i).Refresh
            Shape1(i).Refresh
        Else
            Shape1(i).Height = 15
            Shape1(i).Top = gitop

            Label33(i).Top = gitop - 250
            Label33(i).Caption = siUms(i)
            Label33(i).Refresh
        End If

    Next i


    For i = 0 To 9
        Shape1(i).Visible = True
        If cTab = "TOPIA" Then
            Shape1(i).BackColor = vbRed
        Else
            Shape1(i).BackColor = vbYellow
        End If
        Shape1(i).Refresh
    Next i


    Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Zeigegrafik"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        
    End If
    
End Sub

Private Sub ZeigegrafikB()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim siUms(5)        As Single
    Dim sLiefBez(5)     As String
    Dim rsUms           As Recordset
    Dim rsUmsVJ         As Recordset
    Dim lLinr(5)        As Long
    Dim i               As Integer
    
    
    Dim dbuffer     As Double
    Dim dMax        As Double
    Dim sDatum      As String
    Dim iJahr       As Integer
    Dim lHeuteVJ    As Long
    Dim cheuteVJ    As String
    Dim siWert      As Single
    Dim Wert1       As Integer
    
    
    
    
    
    For i = 0 To 5
        Shape2(i).Visible = False
    Next i
    
    sSQL = " Select * from TH3" & srechnertab & " order by KUCUT desc "
    
                           
    Set rsUms = gdBase.OpenRecordset(sSQL)
    If Not rsUms.EOF Then
        i = 0
        rsUms.MoveFirst
    
        Do While Not rsUms.EOF
        If i > 5 Then Exit Do
        
            If Not IsNull(rsUms!KUCUT) Then
                siUms(i) = rsUms!KUCUT
            Else
                siUms(i) = 0
            End If
            
            If Not IsNull(rsUms!BEDIENER) Then
                lLinr(i) = rsUms!BEDIENER
            Else
                lLinr(i) = "0"
            End If
            
            If Not IsNull(rsUms!bedname) Then
                sLiefBez(i) = rsUms!bedname
            Else
                sLiefBez(i) = ""
            End If
            
            
            i = i + 1
        rsUms.MoveNext
        Loop
    End If
    rsUms.Close
    
    For i = 0 To 5
        Label8(i).Caption = sLiefBez(i)
        Label8(i).Refresh
    
        Label7(i).Caption = lLinr(i)
        Label7(i).Refresh
        Label7(i).ToolTipText = sLiefBez(i)
        Label34(i).ToolTipText = sLiefBez(i)
        
        
    Next i
    
    dbuffer = 0
    dMax = 0
    
    For i = 0 To 5
        dbuffer = siUms(i)
        If dbuffer > dMax Then
            dMax = dbuffer
        End If
    Next i
    dMax = IIf(dMax = 0, 1, dMax)
    
    Dim screenfak As Double
    
    Select Case Screen.Height
        Case Is > 15000
            screenfak = 5000
        Case Is > 12000
            screenfak = 3800
        Case Is > 11000
            screenfak = 3500
        Case Is > 10000
            screenfak = 3000
        Case Is > 8000
            screenfak = 2800
        Case Else
            screenfak = 2000
    End Select
        
    
    For i = 0 To 5
        
        If siUms(i) > 0 Then
            Shape2(i).Height = (screenfak / dMax) * IIf(siUms(i) < 0, 0, siUms(i))
            Shape2(i).Top = gitop - ((screenfak / dMax) * siUms(i))
           
            Label34(i).Top = Shape2(i).Top - 250
           
            Label34(i).Caption = Format$(siUms(i), "###,##0.00")
            
            Label34(i).Refresh
        Else
            Shape2(i).Height = 15
            Shape2(i).Top = gitop

            Label34(i).Top = gitop - 250
            Label34(i).Caption = siUms(i)
            Label34(i).Refresh
        End If

    Next i


    For i = 0 To 5
        Shape2(i).Visible = True
    Next i


    Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ZeigegrafikB"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        
    End If
    
End Sub

Private Sub PositionierenWK21n()
On Error GoTo LOKAL_ERROR

    

    With Frame4
        .Top = 3360
        .Left = 120
        .Width = 11655
        .Height = 5055
    End With

    With Frame3
        .Top = 3360
        .Left = 120
        .Width = 11655
        .Height = 5055
    End With
    
    With Frame2
        .Top = 3360
        .Left = 120
        .Width = 11655
        .Height = 5055
    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWK21n"
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub starte(imod As Integer)
On Error GoTo LOKAL_ERROR

    Dim sSQL                As String
    Dim sdateTimeDat        As String
    Dim sdateDateDat        As String
    Dim cPfad2              As String
    Dim iWochentag      As Integer
    
    Dim cVon                As String
    Dim cBis                As String
    Dim iFil                As Integer
    
    cVon = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))

    cVon = Format$(cVon, "DD.MM.YY")
    cBis = DateValue(Now) - 1

    cPfad2 = gcDBPfad
    If Right(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    Select Case imod
    
        Case 1
            sdateDateDat = DateValue(Now)
            sdateTimeDat = Format(TimeValue(Now), "hh:mm")
            
            db_CopyLokal cPfad2, "Kissdata.MDB", App.Path & "\lokal.mdb", lab, txtStatus, labglo, Label2(34)
        
        
        
        
            If gbKopOhneAuswertung = False Then
                iFil = 0  'alle fils
                TopBediener cVon, cBis, iFil
                
                Frame3.Visible = True
                Frame2.Visible = False
                Frame4.Visible = False
                Me.Refresh
            Else
                Frame3.Visible = False
            End If
        
            If gsAnforderung = "ALLES" Then
                Set dabalokal = OpenDatabase(App.Path & "\lokal.mdb", False, False, "MS Access;PWD=" & gsPasswort)
                db_Reindizieren dabalokal, lab, txtStatus, labglo
                dabalokal.Close
            Else
                Set dabalokal = OpenDatabase(App.Path & "\lokal.mdb", False, False, "MS Access;PWD=" & gsPasswort)
                db_ReindizierenLo dabalokal, lab, txtStatus, labglo, Frame4, Frame3
                dabalokal.Close
            End If
            
            sdateTimeDat = SwapStr(sdateTimeDat, ":", "")
    
            sSQL = "Update WKEINSTE Set LocalTime = " & sdateTimeDat
            gdApp.Execute sSQL, dbFailOnError
            
            sSQL = "Update WKEINSTE Set LocalDat = '" & sdateDateDat & "'"
            gdApp.Execute sSQL, dbFailOnError
                
        Case 2
           
            iWochentag = Weekday(Now)
            db_CopySicher_zip cPfad2, lab, txtStatus, labglo
        
        
        Case 3
            db_CopyLokal cPfad2, "Kissdata1.MDB", cPfad2 & "Kissdata.MDB", lab, txtStatus, labglo, Label2(34)
            Set gdBase = OpenDatabase(cPfad2 & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
            
            db_Reindizieren gdBase, lab, txtStatus, labglo
            
            gdBase.Close
    End Select
    
    giCopyMod = 0
    Unload Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "starte"
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
 
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
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function db_CopyLokal(sPfad As String, sDBOld As String, sDBNew As String, lab As Label, _
txtStatus As TextBox, labglo As Label, labb As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim dbOld               As DAO.Database
    Dim dbNew               As DAO.Database
    Dim lAnzTable           As Long
    Dim lCount              As Long
    Dim lgMax               As Long
    Dim lTabMax             As Long
    Dim name                As String
    Dim lMax                As Long
    Dim sAnzeigetext        As String
    Dim i                   As Integer
    Dim dErgebnis           As Double
    
    '1.Text
    
    
     

    sAnzeigetext = "Pennerartikel" & vbCrLf
    sAnzeigetext = sAnzeigetext & "Die Ermittlung der Pennerartikel finden" & vbCrLf
    sAnzeigetext = sAnzeigetext & "Sie unter Stammdaten/Artikel…/Pennerartikel"

    labb.Caption = sAnzeigetext
    labb.Refresh
    
    Set dbOld = OpenDatabase(sPfad & sDBOld, False, False, "MS Access;PWD=" & gsPasswort)

    If gsAnforderung = "ALLES" Then
        Kill sDBNew
        Set dbNew = CreateDatabase(sDBNew, dbLangGeneral, dbVersion40)
        dbNew.Close
    Else
        Set dbNew = OpenDatabase(sDBNew, False, False, "MS Access;PWD=" & gsPasswort)
        For i = 0 To UBound(gsLokalTabellen)
            loeschNEW gsLokalTabellen(i), dbNew
    
        Next i
        dbNew.TableDefs.Refresh
        lAnzTable = dbNew.TableDefs.Count
        
        For lCount = 0 To lAnzTable - 1
        name = dbNew.TableDefs(lCount).name
        
        If UCase(Left(name, 1)) = "Q" Then
            loeschNEW name, dbNew
        End If
        
        If UCase(Left(name, 1)) = "X" Then
            loeschNEW name, dbNew
        End If
        
        Next lCount
        dbNew.Close
    End If
      
    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbank wird kopiert, bitte warten..."
    labglo.Refresh
        
    dbOld.TableDefs.Refresh
    lAnzTable = dbOld.TableDefs.Count
    
    For lCount = 0 To lAnzTable - 1
        lMax = lMax + dbOld.TableDefs(lCount).RecordCount
    Next lCount
    
    
    dbOld.TableDefs.Refresh
    lAnzTable = dbOld.TableDefs.Count
    
    lgMax = 0
    
    For lCount = 0 To lAnzTable - 1
        name = dbOld.TableDefs(lCount).name
        
        If UCase(Left(name, 4)) = "MSYS" Then
'            MsgBox name
        Else
        
            lab.Caption = name
            lab.Refresh
            
            lTabMax = dbOld.TableDefs(lCount).RecordCount
            
            If gsAnforderung = "ALLES" Then
                TransferTab dbOld, sDBNew, name
            Else
                For i = 0 To UBound(gsLokalTabellen)
                    If UCase(gsLokalTabellen(i)) = UCase(name) Then
                        TransferTab dbOld, sDBNew, name
                    End If
            
                Next i
                
                If UCase(Left$(name, 1)) = "Q" Then
                    TransferTab dbOld, sDBNew, name
                End If
                
                If UCase(Left$(name, 1)) = "X" Then
                    TransferTab dbOld, sDBNew, name
                End If
            
            End If
    
            lgMax = lgMax + lTabMax
            dErgebnis = lgMax / (lMax / 100)
            txtStatus.Text = CStr(dErgebnis)
            
            Select Case CLng(txtStatus.Text)
                Case Is > 60
                    '3.Text
                    sAnzeigetext = "Pennerartikel" & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Des Weiteren können Sie auch die Neuheiten variabel definieren, die ausgeschlossen werden sollen." & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Die Ergebnisartikel können im Anschluss mit einer" & vbCrLf
                    sAnzeigetext = sAnzeigetext & "selbstdefinierten Farbe dauerhaft markiert werden." & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Noch Fragen? 0511/9559112" & vbCrLf
                    labb.Caption = sAnzeigetext
                    labb.Refresh
                
                Case Is > 40
                
   

                    '2.Text
                    sAnzeigetext = "Pennerartikel" & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Pennerartikel werden jetzt aufgrund" & vbCrLf
                    sAnzeigetext = sAnzeigetext & "der Lagerumschlagsgeschwindigkeit eines Artikels definiert." & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Sie können sich alle Artikel unterhalb einer bestimmten Lagerumschlagsgeschwindigkeit anzeigen lassen" & vbCrLf
                    labb.Caption = sAnzeigetext
                    labb.Refresh
                Case Is > 20
                    
                    If gbKopOhneAuswertung = False Then
                        If Label1(9).Caption = "Drei 'heimliche Verlustbringer' - eine Auswahl" Then
                     
                        Else
                            Label1(9).Caption = "Drei 'heimliche Verlustbringer' - eine Auswahl"
                            Label1(7).Caption = "Möchten Sie eine Übersicht über Ihr 'totes Kapital' erstellen, dann "
                            Label1(7).Caption = Label1(7).Caption & "gehen Sie auf ARTIKELLISTEN.../Diverse Listen"
                     
                     
                            TotalaltNoVerkauft
                            Frame4.Visible = True
                            Frame2.Visible = False
                            Frame3.Visible = False
                            Me.Refresh
                        End If
                    Else
                        Frame4.Visible = False
                    End If
                
                Case Is > 8
                
                    If gbKopOhneAuswertung = False Then
                        If Label1(5).Caption = "Die besten Artikel - Verkaufszahlen im Zeitraum" Then
                        
                        Else
                            Dim cVon                As String
                            Dim cBis                As String
                            Dim iFil                As Integer
                            
                            cVon = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
                        
                            cVon = Format$(cVon, "DD.MM.YY")
                            cBis = DateValue(Now) - 1
                            
                            Label1(5).Caption = "Die besten Artikel - Verkaufszahlen im Zeitraum"
                            TopArtikelUmsatz cVon, cBis, iFil, "menge"
                        End If
                    Else
                        Frame2.Visible = False
                    End If
                
                    sAnzeigetext = "Die Geschäftsanalyse" & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Unter Statistik/Geschäftsanalyse" & vbCrLf
                    sAnzeigetext = sAnzeigetext & "Hier sehen Sie die wichtigsten Eckzahlen Ihres Unternehmens im Überblick."
                    
                    labb.Caption = sAnzeigetext
                    labb.Refresh
            End Select
        End If
    Next lCount
    
    dbOld.Close

    labglo.ForeColor = vbBlack
    labglo.Caption = "Fertig"
    labglo.Refresh
    
Exit Function
LOKAL_ERROR:


If err.Number = 53 Then
    Resume Next
ElseIf err.Number = 70 Then
    Exit Function
Else
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "db_CopyLokal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End If
End Function
