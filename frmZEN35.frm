VERSION 5.00
Begin VB.Form frmZEN35 
   BackColor       =   &H000000C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   " - Datenbank "
   ClientHeight    =   8910
   ClientLeft      =   255
   ClientTop       =   1695
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
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   120
      TabIndex        =   52
      Top             =   3240
      Width           =   11655
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         Index           =   0
         Left            =   240
         TabIndex        =   75
         Top             =   4320
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
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   74
         Top             =   240
         Width           =   3735
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
         TabIndex        =   73
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
         Index           =   6
         Left            =   4440
         TabIndex        =   72
         Top             =   240
         Width           =   1095
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   4080
         Width           =   735
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         Index           =   2
         Left            =   2070
         TabIndex        =   60
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Shape2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   5
         Left            =   7800
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         Index           =   2
         Left            =   1920
         TabIndex        =   54
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
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
         TabIndex        =   53
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   11655
      Begin VB.Label Label3 
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
         Index           =   9
         Left            =   7800
         TabIndex        =   51
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Index           =   8
         Left            =   6960
         TabIndex        =   50
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Index           =   7
         Left            =   6120
         TabIndex        =   49
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Index           =   6
         Left            =   5280
         TabIndex        =   48
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   4440
         TabIndex        =   47
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   3600
         TabIndex        =   46
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   2760
         TabIndex        =   45
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         TabIndex        =   44
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   1080
         TabIndex        =   43
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   4320
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
         Height          =   615
         Index           =   5
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   3735
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
         TabIndex        =   30
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
         Index           =   3
         Left            =   4440
         TabIndex        =   29
         Top             =   240
         Width           =   1095
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   4080
         Width           =   735
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         Index           =   6
         Left            =   5430
         TabIndex        =   9
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Shape1 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   9
         Left            =   7800
         TabIndex        =   32
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
         TabIndex        =   33
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
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         Index           =   6
         Left            =   5280
         TabIndex        =   41
         Tag             =   "Shape"
         Top             =   3960
         Width           =   720
      End
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
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'Kein
      Height          =   3480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   5355
         TabIndex        =   2
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   13
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   5535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   6120
         X2              =   6120
         Y1              =   240
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   0
         X1              =   240
         X2              =   6120
         Y1              =   2400
         Y2              =   2400
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5415
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
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   5535
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmZEN35"
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
Private Sub Form_Activate()
On Error GoTo LOKAL_ERROR

    Dim cVon As String
    Dim cBis As String
    Dim iFil As Integer

    Label1(5).Caption = "Die besten Lieferanten - Verkaufszahlen im Zeitraum"
    
    Label1(1).Caption = "Die besten Bediener(Verkaufsschnitt) - Verkaufszahlen im Zeitraum"

    cVon = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
    cVon = Format$(cVon, "DD.MM.YY")
    cBis = DateValue(Now) - 1
    
    iFil = 0  'alle fils
    TopUmsatz cVon, cBis, iFil, "menge"
    
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
Private Sub TopUmsatz(cVon As String, cBis As String, iFil, cWert As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cvon1 As String
    Dim cbis1 As String
    
    loeschNEW "topi", gdbMdb
    loeschNEW "TOPUMSATZ", gdbMdb
    CreateTable "TOPI", gdbMdb
    
    cvon1 = cVon
    cbis1 = cBis
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    Screen.MousePointer = 11
    
    If cWert = "rertrag" Then

        loeschNEW "Ert", gdbMdb
        CreateTable "ERT", gdbMdb
        
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
        
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Update Ert set rertrag = ((Preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * menge) "
        sSQL = sSQL & " where mwst = 'V' "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Update Ert set rertrag = ((Preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * menge) "
        sSQL = sSQL & " where mwst = 'E' "
        gdbMdb.Execute sSQL, dbFailOnError
        
        sSQL = "Update Ert set rertrag = ((Preis * 100)/(100 + " & gdMWStO & " )) - (EKPR * menge) "
        sSQL = sSQL & " where mwst = 'O' "
        gdbMdb.Execute sSQL, dbFailOnError
    
    
    
    
    
        sSQL = "Select sum(" & cWert & ") as maxi ,linr into TOPUMSATZ "
        sSQL = sSQL & " from Ert group by linr "
        gdbMdb.Execute sSQL, dbFailOnError
    
    
        loeschNEW "Ert", gdbMdb

    
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
        gdbMdb.Execute sSQL, dbFailOnError

    End If

    
    
    


    sSQL = "insert into topi SELECT TOP 10 LINR, maxi "
    sSQL = sSQL & " from TOPUMSATZ order by maxi desc"
    gdbMdb.Execute sSQL, dbFailOnError

    sSQL = "Update topi inner join lisrt on topi.linr = lisrt.linr "
    sSQL = sSQL & " SET TOPi.liefbez = lisrt.liefbez "
    gdbMdb.Execute sSQL, dbFailOnError



    sSQL = "Update topi "
    sSQL = sSQL & " SET TOPi.von = '" & cvon1 & "'"
    gdbMdb.Execute sSQL, dbFailOnError

    sSQL = "Update topi "
    sSQL = sSQL & " SET TOPi.bis = '" & cbis1 & "'"
    gdbMdb.Execute sSQL, dbFailOnError

    sSQL = "Update topi "
    sSQL = sSQL & " SET TOPi.fil  = 'alle Filialen'"
    gdbMdb.Execute sSQL, dbFailOnError
    
    Label1(3).Caption = cvon1
    Label1(4).Caption = cbis1
    
    
    
    Zeigegrafik "TOPI", cWert

    
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
Private Sub TopBediener(cVon As String, cBis As String, iFil)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cvon1 As String
    Dim cbis1 As String
    
   
    loeschNEW "TB1" & sRechnertab, gdbMdb
    loeschNEW "TB" & sRechnertab, gdbMdb
    
    sSQL = "Create Table TB" & sRechnertab
    sSQL = sSQL & "( BEDIENER INTEGER"
    sSQL = sSQL & ", SMENGE LONG"
    sSQL = sSQL & ", BONANZ LONG"
    sSQL = sSQL & ", KUCUT single"
    sSQL = sSQL & ") "
    gdbMdb.Execute sSQL, dbFailOnError

    
    cvon1 = cVon
    cbis1 = cBis
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    Screen.MousePointer = 11
    
    loeschNEW "AAT" & sRechnertab, gdbMdb
    
    sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN , bediener, sum(Menge) as AMenge into AAT" & sRechnertab
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & " and UMS_OK = 'J' "
    If iFil = 0 Then

    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    sSQL = sSQL & " group by adate,bediener,BELEGNR"

    gdbMdb.Execute sSQL, dbFailOnError
    
    
    sSQL = "Select sum(AMenge) as SMENGE ,BEDIENER,count(ANZKUNDEN) as belegnr into TB1" & sRechnertab
    sSQL = sSQL & " from AAT" & sRechnertab
    sSQL = sSQL & " group by bediener"
    gdbMdb.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into TB" & sRechnertab & " Select SMENGE ,BEDIENER, belegnr as bonanz from TB1" & sRechnertab
    gdbMdb.Execute sSQL, dbFailOnError
    
    sSQL = "Update TB" & sRechnertab & " set KUCUT = sMenge/bonanz "
    gdbMdb.Execute sSQL, dbFailOnError
    
    
    

    loeschNEW "TH3" & sRechnertab, gdbMdb
    sSQL = "Create Table TH3" & sRechnertab
    sSQL = sSQL & "("
    sSQL = sSQL & " KUCUT Double"
    sSQL = sSQL & ", BEDIENER LONG"
    sSQL = sSQL & ", BEDNAME TEXT(35) "
    sSQL = sSQL & ") "
    gdbMdb.Execute sSQL, dbFailOnError

    sSQL = "Insert into TH3" & sRechnertab & " SELECT TOP 3 Bediener, KUCUT "
    sSQL = sSQL & " from TB" & sRechnertab & " order by KUCUT desc"
    gdbMdb.Execute sSQL, dbFailOnError


    sSQL = "Insert into TH3" & sRechnertab & " SELECT TOP 3 Bediener, KUCUT "
    sSQL = sSQL & " from TB" & sRechnertab & " order by KUCUT asc"
    gdbMdb.Execute sSQL, dbFailOnError


    

    
    sSQL = "Update TH3" & sRechnertab & " inner join Bedname on TH3" & sRechnertab & ".bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET TH3" & sRechnertab & ".BEDNAME = BEDNAME.bedname "
    gdbMdb.Execute sSQL, dbFailOnError
    
    Label1(6).Caption = cvon1
    Label1(2).Caption = cbis1
    
    ZeigegrafikB

    
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
    Dim lLINR(9)        As Long
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
    
                           
    Set rsUms = gdbMdb.OpenRecordset(sSQL)
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
            
            If Not IsNull(rsUms!LINR) Then
                lLINR(i) = rsUms!LINR
            Else
                lLINR(i) = "0"
            End If
            
            If Not IsNull(rsUms!liefBEZ) Then
                sLiefBez(i) = rsUms!liefBEZ
            Else
                sLiefBez(i) = ""
            End If
            
            
            i = i + 1
        rsUms.MoveNext
        Loop
    End If
    rsUms.Close
    
    For i = 0 To 9
        Label3(i).Caption = sLiefBez(i)
        Label3(i).Refresh
    
        Label2(i).Caption = lLINR(i)
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
    Dim lLINR(5)        As Long
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
    
    sSQL = " Select * from TH3" & sRechnertab & " order by KUCUT desc "
    
                           
    Set rsUms = gdbMdb.OpenRecordset(sSQL)
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
                lLINR(i) = rsUms!BEDIENER
            Else
                lLINR(i) = "0"
            End If
            
            If Not IsNull(rsUms!BEDNAME) Then
                sLiefBez(i) = rsUms!BEDNAME
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
    
        Label7(i).Caption = lLINR(i)
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

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    alternativFarbform Me, Label1(0)
    Skalieren Me, True, True: Schrift Me
    
    gitop = Shape1(0).Top
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Datenbank kopieren ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
    
End Sub
Private Sub starte(imod As Integer)
On Error GoTo LOKAL_ERROR

    Dim cQuelle             As String
    Dim cZiel               As String
    Dim lfail               As Long
    Dim lret                As Long
    Dim sSQL                As String
    Dim sdateTimeDat        As String
    Dim sdateDateDat        As String
    Dim cPfad2              As String
    
    Dim cVon                As String
    Dim cBis                As String
    Dim iFil                As Integer
    
    

    cVon = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))

    cVon = Format$(cVon, "DD.MM.YY")
    cBis = DateValue(Now) - 1

    
    cPfad2 = gcDBPfad
    If Right$(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    Select Case imod
        Case 1
        
            sdateDateDat = DateValue(Now)
            sdateTimeDat = Format(TimeValue(Now), "hh:mm")
            
            
            db_CopyLokal cPfad2, "Kissdata.MDB", App.Path & "\lokal.mdb", lab, txtStatus, labglo, Label2(13)
            

            iFil = 0  'alle fils
            TopBediener cVon, cBis, iFil
            
            
            Frame3.Visible = True
            Frame2.Visible = False
            Me.Refresh
            
            Set dabalokal = OpenDatabase(App.Path & "\lokal.mdb", False, False, "MS Access;PWD=" & gsPasswort)
            
            Label2(13).Caption = "Neu - Lagerumschlag" & vbCrLf
            Label2(13).Caption = Label2(13).Caption & "Unter Stammdaten/Artikel bearbeiten"
            Label2(13).Refresh
            If gsAnforderung = "ALLES" Then
                db_Reindizieren lab, txtStatus, labglo, dabalokal
            Else
                db_ReindizierenLo lab, txtStatus, labglo, dabalokal
            End If
            
            dabalokal.Close
            
            sdateTimeDat = SwapStr(sdateTimeDat, ":", "")
            
            sSQL = "Update zenteins Set LocalTime = " & sdateTimeDat
            dbApp.Execute sSQL, dbFailOnError
        
            sSQL = "Update zenteins Set LocalDat = '" & sdateDateDat & "'"
            dbApp.Execute sSQL, dbFailOnError
            
        Case 2
        
            db_CopyLokal cPfad2, "Kissdata1.MDB", cPfad2 & "Kissdata.MDB", lab, txtStatus, labglo, Label2(13)
            
            iFil = 0  'alle fils
            TopBediener cVon, cBis, iFil
            Frame3.Visible = True
            Frame2.Visible = False
            Me.Refresh
            
            Set gdbMdb = OpenDatabase(gcDBPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
            
            db_Reindizieren lab, txtStatus, labglo, gdbMdb
            
            gdbMdb.Close
            
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
