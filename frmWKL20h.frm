VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmWK20h 
   Caption         =   "Winkiss Zahlung über EC - Lastschrift"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame18 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.Frame Frame19 
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   7560
         TabIndex        =   33
         Top             =   240
         Width           =   4215
         Begin VB.ListBox List11 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7020
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   3975
         End
      End
      Begin VB.Frame Frame22 
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   3120
         TabIndex        =   27
         Top             =   4080
         Width           =   4455
         Begin VB.CommandButton Command10 
            Caption         =   "manuelle Eingabe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   4
            Left            =   2280
            TabIndex        =   32
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Schließen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   2160
            Width           =   4215
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Eingabe löschen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Kunden- daten..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   2280
            TabIndex        =   29
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Beleg drucken"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame20 
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
         Height          =   1335
         Left            =   5400
         TabIndex        =   15
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   3240
            MaxLength       =   10
            TabIndex        =   21
            Text            =   "Text6"
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   3240
            MaxLength       =   8
            TabIndex        =   20
            Text            =   "Text6"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   19
            Text            =   "Text6"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   18
            Text            =   "Text6"
            Top             =   2040
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   2640
            Width           =   2535
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Zurück"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   6360
            TabIndex        =   16
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            Caption         =   "Kontonummer:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            Caption         =   "Bankleitzahl:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   25
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            Caption         =   "gültig bis:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   24
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   4080
            TabIndex        =   23
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "manuelle Eingabe EC-Karte"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   8775
         End
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3960
         TabIndex        =   1
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton Command13 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   1680
            TabIndex        =   11
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   2400
            TabIndex        =   10
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   3120
            TabIndex        =   9
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   3840
            TabIndex        =   8
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   6
            Left            =   4560
            TabIndex        =   7
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   7
            Left            =   5280
            TabIndex        =   6
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   8
            Left            =   6000
            TabIndex        =   5
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   9
            Left            =   6720
            TabIndex        =   4
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   10
            Left            =   7440
            TabIndex        =   3
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   11
            Left            =   8160
            TabIndex        =   2
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFF00&
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9000
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   120
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
         Handshaking     =   1
         RThreshold      =   1
         ParitySetting   =   2
         SThreshold      =   1
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         Caption         =   "Bitte die EC-Karte durch den Kartenleser ziehen!"
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
         Left            =   840
         TabIndex        =   42
         Top             =   2040
         Width           =   6855
      End
      Begin VB.Label Label11 
         Caption         =   "zu zahlender Betrag:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   41
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         Caption         =   "Zahlung über EC-Lastschrift"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label11 
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   3
         Left            =   5160
         TabIndex        =   39
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         Caption         =   "Kunde / Kontoinhaber:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Index           =   4
         Left            =   600
         TabIndex        =   38
         Top             =   3240
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Zentriert
         Caption         =   "Karte geprüft?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   5640
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Zentriert
         Caption         =   "Unterschrift geprüft?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   6360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Zentriert
         Caption         =   "Wenn alles okay, dann Beleg drucken und die EC-Karte zurück an Kunden!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   7080
         Visible         =   0   'False
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmWK20h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


