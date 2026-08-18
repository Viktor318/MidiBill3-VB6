VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0518EEBD-7F0E-4513-8491-A0221C9008A2}#2.1#0"; "midiio2k.ocx"
Object = "{4424C993-EABF-4A03-9BA9-369E0F07466E}#1.2#0"; "midifl2k.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Rögzített egyetlen
   Caption         =   "MIDI billentyûzet megszólaltatása 4 hangszeren egyszerre : Kezdet.mb3"
   ClientHeight    =   6744
   ClientLeft      =   1860
   ClientTop       =   2376
   ClientWidth     =   9288
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.6
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "MidiBill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'Z-rendezés használata
   ScaleHeight     =   6744
   ScaleWidth      =   9288
   Begin MidifileLib.MIDIFile MIDIFile1 
      Left            =   8400
      Top             =   5760
      _Version        =   65538
      _ExtentX        =   677
      _ExtentY        =   677
      _StockProps     =   0
      Filename        =   ""
   End
   Begin MidiioLib.MIDIOutput MIDIOutput1 
      Left            =   7440
      Top             =   3240
      _Version        =   131073
      _ExtentX        =   677
      _ExtentY        =   677
      _StockProps     =   0
   End
   Begin VB.CheckBox Pedal 
      Alignment       =   1  'Jobbra igazít
      Caption         =   "Fordított pedál mûködés"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3240
      TabIndex        =   46
      Top             =   6120
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Caption         =   "Aktív hangszerek"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   40
      Top             =   240
      Width           =   2775
      Begin VB.CheckBox J1 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   44
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox J1 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   43
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox J1 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   42
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox J1 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Value           =   1  'Bejelölt
         Width           =   375
      End
   End
   Begin VB.Frame Keret1 
      Caption         =   "4. hangszer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   3
      Left            =   3240
      TabIndex        =   25
      Top             =   3120
      Width           =   2775
      Begin MSComctlLib.Slider Vibrato1 
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   36
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   868
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Tér1 
         Height          =   510
         Index           =   3
         Left            =   480
         TabIndex        =   35
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Min             =   -64
         Max             =   63
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Volume1 
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   34
         Top             =   840
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   868
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         SelStart        =   127
         TickStyle       =   1
         TickFrequency   =   16
         Value           =   127
      End
      Begin VB.ComboBox Hangszerek1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Text            =   "Hangszerek4"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Vib"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Tér"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Her"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Keret1 
      Caption         =   "3. hangszer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   3120
      Width           =   2775
      Begin MSComctlLib.Slider Vibrato1 
         Height          =   510
         Index           =   2
         Left            =   480
         TabIndex        =   29
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Tér1 
         Height          =   510
         Index           =   2
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Min             =   -64
         Max             =   63
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Volume1 
         Height          =   510
         Index           =   2
         Left            =   480
         TabIndex        =   27
         Top             =   840
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         SelStart        =   127
         TickStyle       =   1
         TickFrequency   =   16
         Value           =   127
      End
      Begin VB.ComboBox Hangszerek1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Text            =   "Hangszerek3"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Vib"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Tér"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Her"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Keret1 
      Caption         =   "2. hangszer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox Hangszerek1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Text            =   "Hangszerek2"
         Top             =   360
         Width           =   2295
      End
      Begin MSComctlLib.Slider Tér1 
         Height          =   510
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Min             =   -64
         Max             =   63
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Volume1 
         Height          =   510
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         SelStart        =   127
         TickStyle       =   1
         TickFrequency   =   16
         Value           =   127
      End
      Begin MSComctlLib.Slider Vibrato1 
         Height          =   510
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Vib"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Tér"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Her"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.CommandButton MIDI_Reset 
      Caption         =   "Csend legyen!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Alap hangerõ beállítása"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   6240
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
      Begin MSComctlLib.Slider Volume 
         Height          =   560
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2535
         _ExtentX        =   4466
         _ExtentY        =   995
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   5
         Max             =   100
         SelStart        =   70
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   70
      End
   End
   Begin VB.CommandButton Kilépés 
      Caption         =   "Kilépés"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Keret1 
      Caption         =   "1. hangszer: Aktív"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox Hangszerek1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         ItemData        =   "MidiBill.frx":030A
         Left            =   240
         List            =   "MidiBill.frx":030C
         TabIndex        =   12
         Text            =   "Hangszerek1"
         Top             =   360
         Width           =   2295
      End
      Begin MSComctlLib.Slider Vibrato1 
         Height          =   510
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Tér1 
         Height          =   510
         Index           =   0
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   910
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Min             =   -64
         Max             =   63
         TickStyle       =   1
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Volume1 
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   2175
         _ExtentX        =   3831
         _ExtentY        =   868
         _Version        =   393216
         LargeChange     =   0
         SmallChange     =   8
         Max             =   127
         SelStart        =   127
         TickStyle       =   1
         TickFrequency   =   16
         Value           =   127
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Vib"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Tér"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Jobbra igazít
         Caption         =   "Her"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MIDI Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
      Begin VB.ComboBox MIDI_Out 
         Appearance      =   0  'Lapos
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Lenyíló lista
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MIDI In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
      Begin MidiioLib.MIDIInput MIDIInput1 
         Left            =   1200
         Top             =   0
         _Version        =   131073
         _ExtentX        =   677
         _ExtentY        =   677
         _StockProps     =   0
      End
      Begin VB.ComboBox MIDI_In 
         Appearance      =   0  'Lapos
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Lenyíló lista
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label MIDI_uzi 
      Caption         =   "MIDI üzenetek: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Menu Fájl 
      Caption         =   "&Fájl"
      Index           =   0
      Begin VB.Menu Tölt 
         Caption         =   "Adatok &betöltése"
      End
      Begin VB.Menu Ment 
         Caption         =   "Adatok &mentése"
      End
      Begin VB.Menu v 
         Caption         =   "-"
      End
      Begin VB.Menu Kilép 
         Caption         =   "&Kilépés"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Változók definiálása
    Dim Inv_pedal As Integer
    Dim Last_pedal As Integer
    Dim KezdetFile As String  'A Kezdet.mb3 teljes, App.Path-hoz képest egységesített elérési útja
    Dim NoteOn(1 To 4, 0 To 127) As Boolean  'Nyilvántartás: melyik hang szól éppen az egyes csatornákon
    Dim HeldNotes(0 To 127) As Integer  'Melyik fizikai billentyû van lenyomva (0=nincs, egyébként a velocity), a csatorna be/ki állapotától FÜGGETLENÜL
    Dim MidiOutHibaJelezve As Boolean  'Hogy a kimeneti hiba csak egyszer villantson üzenetet, ne minden billentyûnél
    'Def. vége
    

Private Sub Form_Load()
    Dim i As Integer
    Inv_pedal = Pedal.Value
    Last_pedal = 0

    ' Center the form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    Form1.Show
    
    If (MIDIInput1.DeviceCount = 0) Then
     MsgBox ("MIDI bemenet nem található! A program most kilép.")
     Unload Me
     Exit Sub
    End If
    
    '
    ' Fill input device combo box
    '
    For i = 0 To MIDIInput1.DeviceCount - 1
     MIDIInput1.DeviceID = i
     MIDI_In.AddItem MIDIInput1.ProductName
    Next
    '
    ' Select first in list
    '
    MIDI_In.ListIndex = 0
    MIDIInput1.DeviceID = 0
    
    '
    ' Fill output device combo box
    '
    For i = -1 To MIDIOutput1.DeviceCount - 1
     MIDIOutput1.DeviceID = i
     MIDI_Out.AddItem MIDIOutput1.ProductName
    Next
    '
    ' Select first in list
    '
    MIDI_Out.ListIndex = 0
    MIDIOutput1.DeviceID = 0

    '
    ' General MIDI patches
    '
    MIDI_Hangok
    
    Dim c As Integer
    For c = 0 To 3
      For i = 1 To 128
        Hangszerek1(c).AddItem Hang(i)
      Next i
      Hangszerek1(c).ListIndex = 0
    Next c
    
    'A Kezdet.mb3 elérési útja mindig App.Path-hoz képest, konzisztensen a
    ' fájl-dialógusok InitDir beállításával (relatív útvonal a munkakönyvtártól
    ' függött volna, ami eltérhet App.Path-tól)
    KezdetFile = App.Path
    If Right(KezdetFile, 1) <> "\" Then KezdetFile = KezdetFile & "\"
    KezdetFile = KezdetFile & "Kezdet.mb3"

    If Dir(KezdetFile) <> "" Then 'megvizsgálja van-e már elmentett kezdõ beállítás
     Load_File (KezdetFile) 'utolsó elmentettbeállítás betöltése
    Else
     Call Init_Ports 'Alap értékek aktualizálása
    End If
    
    MIDI_Out.SetFocus 'egér fókusz beállítása
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next 'Ha a MIDI eszköz már nem elérhetõ (pl. kihúzták), ne akadjon el a bezárás
    MIDIInput1.Action = MIDIIN_STOP
    ' Close midiin port
    MIDIInput1.Action = MIDIIN_CLOSE
    ' Close
    MIDIOutput1.Action = MIDIOUT_CLOSE
    
End Sub

Private Sub Pedal_Click()
 Inv_pedal = Pedal.Value
 
 'Pedál mûködés megfordult
  Last_pedal = 127 - Last_pedal 'Utolsó mentett pedál üzenet invertálása
  
  'Küldendõ MIDI üzenet megjelenítése
  MIDI_uzi.Caption = "MIDI üzenetek: " & CONTROLLER_CHANGE & ", " & DAMPER_PEDAL & ", " & Last_pedal
  
  'Tell MIDIOutput1 to send the MIDI data
  Call Send_MIDI(CONTROLLER_CHANGE, DAMPER_PEDAL, Last_pedal)
 
End Sub

Private Sub Tölt_Click()
 Dim InpName As String, T As String
 On Error GoTo Megse_Tolt
 CommonDialog1.CancelError = True
 CommonDialog1.DialogTitle = "MidiBill adatok betöltése"
 CommonDialog1.Filter = "MidiBill adatok (*.mb3)|*.mb3"
 ' a könyvtár beállítása
 CommonDialog1.InitDir = App.Path
 ' a "Megnyitás" párbeszédablak megjelenítése
 CommonDialog1.ShowOpen
 InpName = CommonDialog1.FileName
 T = CommonDialog1.FileTitle
 If Len(InpName) > 0 Then
  ' ha a fájlnév nem üres
  Caption = "MIDI billentyûzet megszólaltatása 4 hangszeren egyszerre : " + T
  Load_File (InpName)
 End If
 Exit Sub

Megse_Tolt:
 ' A felhasználó Mégse-t nyomott a Megnyitás párbeszédablakban - nincs teendõ
End Sub

Private Sub Ment_Click()
 Dim OutName As String, T As String, myAnswer As String
 Dim SaveOK As Boolean
 On Error GoTo Megse_Ment
 CommonDialog1.CancelError = True
 SaveOK = False 'Menthetõ-e a fájl
 Do While (SaveOK = False)
    CommonDialog1.DialogTitle = "MidiBill adatok mentése"
    CommonDialog1.Filter = "MidiBill adatok (*.mb3)|*.mb3"
    ' a könyvtár beállítása
    CommonDialog1.InitDir = App.Path
    ' a "Mentés másként" párbeszédablak megjelenítése
    CommonDialog1.ShowSave
    OutName = CommonDialog1.FileName
    T = CommonDialog1.FileTitle
    If Len(OutName) > 0 Then  'ha a fájlnév nem üres
        If Dir(OutName) <> "" Then 'a fájl már létezik
            myAnswer = MsgBox("A fájl már létezik. Felülírjam?", vbYesNo + vbQuestion, "MidiBill")
            If myAnswer = 6 Then SaveOK = True 'a fájl felülírható
        Else 'a fájl még nem létezik
            SaveOK = True
        End If
    Else
        MsgBox ("Üres fájlnév!")
    End If
 Loop
 Caption = "MIDI billentyûzet megszólaltatása 4 hangszeren egyszerre : " + T
 Save_File (OutName)
 Exit Sub

Megse_Ment:
 ' A felhasználó Mégse-t nyomott a Mentés párbeszédablakban - egyszerûen kilépünk
 ' anélkül, hogy bármit mentenénk vagy hibaüzenetet mutatnánk.
End Sub

Private Sub Kilép_Click()
 'MIDI hangok kikapcsolása
 Call MIDI_Reset_Click
 'Utolsó beállítás mentése
 Save_File (KezdetFile)
 'A MIDI portok bezárását a Form_Unload intézi
 Unload Me
End Sub

Private Sub Hangszerek1_Click(Index As Integer)
  Call Prog_Change(Index + 1, Hangszerek1(Index).ListIndex)
End Sub

Private Sub Volume1_Change(Index As Integer)
  Call Her_Change(Index + 1, Volume1(Index).Value)
End Sub

Private Sub Volume1_Scroll(Index As Integer)
  Call Her_Change(Index + 1, Volume1(Index).Value)
End Sub

Private Sub Tér1_Change(Index As Integer)
  Call Tér_Change(Index + 1, Tér1(Index).Value)
End Sub

Private Sub Tér1_Scroll(Index As Integer)
  Call Tér_Change(Index + 1, Tér1(Index).Value)
End Sub

Private Sub Vibrato1_Change(Index As Integer)
  Call Vibrato_Change(Index + 1, Vibrato1(Index).Value)
End Sub

Private Sub Vibrato1_Scroll(Index As Integer)
  Call Vibrato_Change(Index + 1, Vibrato1(Index).Value)
End Sub

Private Sub Volume_Init()
'Alap hangerõ maximumra állítása jobb és bal csatornán
On Error Resume Next 'Ha a kimeneti eszköz nem elérhetõ, ne akadjon el az inicializálás
    If (MIDIOutput1.HMidiDevice <> 0) Then
      If (MIDIOutput1.HasLRVolume) Then
         MIDIOutput1.VolumeLeft = 32767
         MIDIOutput1.VolumeRight = 32767
      ElseIf (MIDIOutput1.HasVolume) Then
         MIDIOutput1.VolumeLeft = 32767
      End If
    End If
    Call Volume_Refresh
End Sub

Private Sub Volume_Change()
'Hangerõ módosítás kiküldése minden csatornára
    Call Volume_Refresh
End Sub

Private Sub Volume_Scroll()
'Hangerõ módosítás kiküldése minden csatornára
    Call Volume_Refresh
End Sub

Private Sub Volume_Refresh()
'Hangerõ módosítás kiküldése minden csatornára
    Dim c As Integer
    For c = 0 To 3
      Call Volume1_Change(c)
    Next c
End Sub

Private Sub J1_Click(Index As Integer)
 Dim i As Integer

 If J1(Index).Value Then
   Keret1(Index).Caption = Keret1(Index).Caption + ": Aktív"
   Call Sound_Held_Notes(Index + 1) 'A jelenleg lenyomott hangok megszólaltatása az újonnan aktivált csatornán
 Else
   i = Len(Keret1(Index).Caption)
   Keret1(Index).Caption = Left(Keret1(Index).Caption, i - 7)
   Call Turn_Notes_Off(Index + 1) 'Hangok kikapcsolása a megfelelõ MIDI csatornán
 End If

End Sub

Private Sub MIDI_In_Click()
    
    If MIDI_In.ListIndex < 0 Then Exit Sub 'Nincs kiválasztott bemeneti eszköz
    
    On Error GoTo MidiInHiba
    
    If MIDIInput1.State >= MIDISTATE_OPEN Then
        ' Close midiin port
        MIDIInput1.Action = MIDIIN_CLOSE
    End If

    Erase HeldNotes 'Új bemeneti eszközön a régi "lenyomott" állapot már nem releváns
    
    ' Get selected device id
    MIDIInput1.DeviceID = MIDI_In.ListIndex
    
    ' Open and display handle
    MIDIInput1.Action = MIDIIN_OPEN

    ' Start midi input
    MIDIInput1.Action = MIDIIN_START
    Exit Sub

MidiInHiba:
    MsgBox "Nem sikerült megnyitni a kiválasztott MIDI bemeneti eszközt." & vbCrLf & _
           "Lehet, hogy idõközben eltávolították, vagy más program használja." & vbCrLf & vbCrLf & _
           "Hiba: " & Err.Description, vbExclamation, "MidiBill"
End Sub

Private Sub MIDI_Out_Click()

    If MIDI_Out.ListIndex < 0 Then Exit Sub 'Nincs kiválasztott kimeneti eszköz

    On Error GoTo MidiOutHiba

    If MIDIOutput1.State >= MIDISTATE_OPEN Or MIDIOutput1.DeviceID = -1 Then
       ' Close
       MIDIOutput1.Action = MIDIOUT_CLOSE
    End If
    
    ' Open selected device
    MIDIOutput1.DeviceID = MIDI_Out.ListIndex - 1
    MIDIOutput1.Action = MIDIOUT_OPEN
    
    MidiOutHibaJelezve = False 'Új sikeres nyitás után újra figyelmeztethetünk, ha megint hiba lenne
    Erase NoteOn 'Új porton biztosan nem szól semmi - a nyilvántartást is nullázzuk
    
    Call Init_Channels 'MIDI csatornák beállítása kimeneti port váltás esetén(!)
    Exit Sub

MidiOutHiba:
    MsgBox "Nem sikerült megnyitni a kiválasztott MIDI kimeneti eszközt." & vbCrLf & _
           "Lehet, hogy idõközben eltávolították, vagy más program használja." & vbCrLf & vbCrLf & _
           "Hiba: " & Err.Description, vbExclamation, "MidiBill"
End Sub

Private Sub MIDI_Reset_Click()

Dim c As Integer

For c = 1 To 4
  Call Turn_Notes_Off(c, True) 'Vészleállítás: összes hang kikapcsolása a 4 MIDI csatornán, nyilvántartástól függetlenül
Next c

End Sub

Private Sub Kilépés_Click()

 'MIDI hangok kikapcsolása
 Call MIDI_Reset_Click
 'Utolsó beállítás mentése
 Save_File (KezdetFile)
 'A MIDI portok bezárását a Form_Unload intézi
 Unload Me

End Sub

Private Sub MIDIInput1_Message()
 Dim InMessage As Integer
 Dim InData1 As Integer
 Dim InData2 As Integer
 On Error GoTo BemenetHiba
 '
 'This do while loop allows you to take all the messages that are
 'waiting in the message queue.
 '
 Do While MIDIInput1.MessageCount > 0
  '
  'This is the incoming MIDI data
  '
  InMessage = MIDIInput1.Message
  InData1 = MIDIInput1.Data1
  InData2 = MIDIInput1.Data2
  'Sustain pedál üzenetek invertálása
  'Bekapcsolható a Pedal checkbox-ban
  If MIDIInput1.Message = CONTROLLER_CHANGE And InData1 = DAMPER_PEDAL Then
  'Pedál üzenet (normál: lenyomott: 127, felengedett: 0)
    If Inv_pedal = 1 Then 'Ha invertálni kell a mûködést, invertálja az üzenetet
      InData2 = 127 - InData2
    End If
    Last_pedal = InData2 'Pedál üzenet mentése
  End If
  'Küldendõ MIDI üzenet megjelenítése
  MIDI_uzi.Caption = "MIDI üzenetek: " & InMessage & ", " & InData1 & ", " & InData2
  
  'Tell MIDIOutput1 to send the MIDI data
  '
  Call Send_MIDI(InMessage, InData1, InData2)
  '
  'Remove the MIDI data from the MIDI IN queue
  '
  MIDIInput1.Action = MIDIIN_REMOVE
 Loop
 Exit Sub

BemenetHiba:
 ' Csendben kilépünk - ha a bemeneti eszköz menet közben elérhetetlenné vált
 ' (pl. kihúzták), nem akarjuk, hogy a program leálljon.
End Sub

Private Sub Send_MIDI(OutMessage As Integer, OutData1 As Integer, OutData2 As Integer)

  MIDIOutput1.Data1 = OutData1
  MIDIOutput1.Data2 = OutData2

  If OutMessage >= &H80 And OutMessage < &HF0 Then
    ' Csak a valódi "channel voice" üzenetek (Note On/Off, CC, Program Change,
    ' Pitch Bend stb.) csatornáját szabad eltolni, mert ezeknél az alsó 4 bit
    ' hordozza a csatornaszámot (itt mindig csatorna 0-ra normálva érkeznek be).

    ' A HeldNotes nyilvántartás a J1-4 kapcsolók állapotától FÜGGETLENÜL
    ' frissül, hogy tudjuk, mely fizikai billentyûk vannak lenyomva akkor
    ' is, ha épp egyik csatorna sincs bekapcsolva - lásd Sound_Held_Notes.
    If OutMessage = NOTE_ON And OutData2 > 0 Then
        HeldNotes(OutData1) = OutData2
    ElseIf (OutMessage = NOTE_ON And OutData2 = 0) Or OutMessage = NOTE_OFF Then
        HeldNotes(OutData1) = 0
    End If

    ' MIDI üzenetek küldése a 4 csatornára
    Dim c As Integer
    For c = 0 To 3
      If J1(c).Value = 1 Then
        MIDIOutput1.Message = OutMessage + c 'hang küldés a (c+1). csatornára
        Call MIDI_Out_Send
        Call Track_Note(c + 1, OutMessage, OutData1, OutData2)
      End If
    Next c
  Else
    ' Rendszerüzenet (pl. MIDI Clock, Start/Stop, Active Sensing, SysEx: &HF0-&HFF).
    ' Ezeknek nincs csatornájuk, ezért VÁLTOZATLANUL, csak EGYSZER küldjük tovább -
    ' a +1/+2/+3 eltolás itt érvénytelen MIDI státuszbájtot eredményezne.
    MIDIOutput1.Message = OutMessage
    Call MIDI_Out_Send
  End If

End Sub

Private Sub MIDI_Out_Send()
    ' Központi, hibatûrõ MIDI-küldés. Ha a kimeneti eszköz menet közben
    ' eltûnik (pl. kihúzott USB-MIDI interfész), itt kapjuk el a hibát,
    ' ahelyett hogy az egész program leállna. Csak EGYSZER figyelmeztetünk,
    ' nehogy minden lenyomott billentyûnél felugorjon egy hibaablak.
    On Error GoTo SendHiba
    MIDIOutput1.Action = MIDIOUT_SEND
    Exit Sub

SendHiba:
    If Not MidiOutHibaJelezve Then
        MidiOutHibaJelezve = True
        MsgBox "A MIDI kimeneti eszköz nem érhetõ el (lehet, hogy idõközben eltávolították)." & vbCrLf & _
               "Válassz újra egy eszközt a MIDI Out listából.", vbExclamation, "MidiBill"
    End If
End Sub

Private Sub Sound_Held_Notes(Channel As Integer)
    ' Amikor egy korábban kikapcsolt csatornát visszakapcsolunk, ez az
    ' eljárás megszólaltatja rajta az éppen lenyomva tartott hangokat -
    ' az eredeti velocity-vel -, hogy "csatlakozzon" a már játszott akkordhoz.
    Dim n As Integer
    For n = 0 To 127
        If HeldNotes(n) > 0 Then
            MIDIOutput1.Message = NOTE_ON + (Channel - 1)
            MIDIOutput1.Data1 = n
            MIDIOutput1.Data2 = HeldNotes(n)
            Call MIDI_Out_Send
            NoteOn(Channel, n) = True
        End If
    Next n
End Sub

Private Sub Track_Note(Channel As Integer, StatusByte As Integer, NoteNum As Integer, Velocity As Integer)
    ' Nyilvántartjuk, mely hangok szólnak jelenleg az egyes csatornákon,
    ' hogy a Turn_Notes_Off (normál esetben) csak a ténylegesen szóló
    ' hangokat kapcsolja ki, ne mind a 128-at.
    If StatusByte = NOTE_ON And Velocity > 0 Then
        NoteOn(Channel, NoteNum) = True
    ElseIf StatusByte = NOTE_ON And Velocity = 0 Then
        NoteOn(Channel, NoteNum) = False  'Velocity=0 Note On == Note Off (MIDI szabvány)
    ElseIf StatusByte = NOTE_OFF Then
        NoteOn(Channel, NoteNum) = False
    End If
End Sub

Private Sub Save_File(FileName As String)
  Dim c As Integer
  On Error GoTo MentesiHiba
  Open FileName For Output As #1
  'Az adatok fájlba írása
  For c = 0 To 3
    Write #1, Hangszerek1(c).ListIndex, Volume1(c).Value, Tér1(c).Value, Vibrato1(c).Value
  Next c
  Write #1, Volume.Value, J1(0).Value, J1(1).Value, J1(2).Value, J1(3).Value
  Write #1, MIDI_In.ListIndex, MIDI_Out.ListIndex
  'A fájl lezárása
  Close 1
  Exit Sub

MentesiHiba:
  Close 1
  MsgBox "Nem sikerült elmenteni a beállításokat ide:" & vbCrLf & FileName & vbCrLf & vbCrLf & _
         "Hiba: " & Err.Description, vbExclamation, "MidiBill"
End Sub

Private Sub Load_File(FileName As String)
  Dim Adat(5) As Integer
  Dim c As Integer
  On Error GoTo BetoltesiHiba
  Open FileName For Input As #1
  For c = 0 To 3
    Input #1, Adat(1), Adat(2), Adat(3), Adat(4)
    Hangszerek1(c).ListIndex = Adat(1)
    Volume1(c).Value = Adat(2)
    Tér1(c).Value = Adat(3)
    Vibrato1(c).Value = Adat(4)
  Next c
  Input #1, Adat(1), Adat(2), Adat(3), Adat(4), Adat(5)
  Volume.Value = Adat(1)
  J1(0).Value = Adat(2)
  J1(1).Value = Adat(3)
  J1(2).Value = Adat(4)
  J1(3).Value = Adat(5)
  If EOF(1) Then
    Close 1 'Régi típusú mentett adatokkal való kompatibilitás miatt
  Else
    Input #1, Adat(1), Adat(2)
    MIDI_In.ListIndex = Adat(1)
    MIDI_Out.ListIndex = Adat(2)
    Close 1
  End If

  Call Init_Ports 'Betöltött értékek aktualizálása
  Exit Sub

BetoltesiHiba:
  Close 1
  MsgBox "Nem sikerült betölteni a(z) """ & FileName & """ fájlt." & vbCrLf & _
         "Lehet, hogy sérült, vagy nem MidiBill adatfájl." & vbCrLf & vbCrLf & _
         "Hiba: " & Err.Description, vbExclamation, "MidiBill"
  Call Init_Ports 'Biztosítjuk, hogy a portok mindenképp beálljanak, ha a betöltés félbeszakadt
End Sub

Private Sub Init_Ports()
    'MIDI portok beállítása
     Call MIDI_In_Click
     Call MIDI_Out_Click
End Sub

Private Sub Init_Channels()
    'MIDI csatornák beállítása
    Dim c As Integer
    For c = 0 To 3
      Call Hangszerek1_Click(c)
      Call Volume1_Change(c)
      Call Tér1_Change(c)
      Call Vibrato1_Change(c)
    Next c
    Call Volume_Init
End Sub

Private Sub Prog_Change(Channel As Integer, Data As Integer)
    'Patch Midi Out routine
    MIDIOutput1.Message = PROGRAM_CHANGE + (Channel - 1) ' channel: ez a MIDI csatorna változik
    MIDIOutput1.Data1 = Data
    MIDIOutput1.Data2 = 0
    MIDIOutput1.Time = 0
    Call MIDI_Out_Send
End Sub

Private Sub Her_Change(Channel As Integer, Data As Integer)
    If (MIDIOutput1.HMidiDevice <> 0) Then
        MIDIOutput1.Message = CONTROLLER_CHANGE + (Channel - 1)
        MIDIOutput1.Data1 = MAIN_VOLUME
        MIDIOutput1.Data2 = CInt(Data * (Volume.Value / 100))
        MIDIOutput1.Time = 0
        Call MIDI_Out_Send
    End If
End Sub

Private Sub Tér_Change(Channel As Integer, Data As Integer)
    If (MIDIOutput1.HMidiDevice <> 0) Then
        MIDIOutput1.Message = CONTROLLER_CHANGE + (Channel - 1)
        MIDIOutput1.Data1 = PAN
        MIDIOutput1.Data2 = Data + 64
        MIDIOutput1.Time = 0
        Call MIDI_Out_Send
    End If
End Sub

Private Sub Vibrato_Change(Channel As Integer, Data As Integer)
    If (MIDIOutput1.HMidiDevice <> 0) Then
        MIDIOutput1.Message = CONTROLLER_CHANGE + (Channel - 1)
        MIDIOutput1.Data1 = MOD_WHEEL
        MIDIOutput1.Data2 = Data
        MIDIOutput1.Time = 0
        Call MIDI_Out_Send
    End If
End Sub

Private Sub Turn_Notes_Off(Channel As Integer, Optional FullSweep As Boolean = False)

 Dim n As Integer
 
 'Turn sustain pedal off on a MIDI channel
 MIDIOutput1.Message = CONTROLLER_CHANGE + (Channel - 1)
 MIDIOutput1.Data1 = DAMPER_PEDAL
 MIDIOutput1.Data2 = 0
 Call MIDI_Out_Send
 
 If FullSweep Then
   'Vészleállítás (pl. a "Csend legyen!" gombnál): MIND a 128 hangot
   ' lekapcsoljuk, függetlenül a NoteOn nyilvántartástól. Erre azért van
   ' szükség, mert egy esetleg "beragadt" hang (pl. a program indítása
   ' elõttrõl, vagy egy elveszett MIDI-üzenet miatt) nem biztos, hogy
   ' szerepel a nyilvántartásban - a pánikgombnak viszont garantáltan
   ' csendet kell teremtenie.
   For n = 0 To 127
    MIDIOutput1.Message = NOTE_ON + (Channel - 1)
    MIDIOutput1.Data1 = n
    MIDIOutput1.Data2 = 0
    Call MIDI_Out_Send
    NoteOn(Channel, n) = False
   Next
 Else
   'Normál eset (pl. egy hangszer kikapcsolása a Jx checkboxszal): csak
   ' azokat a hangokat kapcsoljuk ki, amikrõl a nyilvántartás szerint
   ' ténylegesen szólnak ezen a csatornán.
   For n = 0 To 127
    If NoteOn(Channel, n) Then
      MIDIOutput1.Message = NOTE_ON + (Channel - 1)
      MIDIOutput1.Data1 = n
      MIDIOutput1.Data2 = 0
      Call MIDI_Out_Send
      NoteOn(Channel, n) = False
    End If
   Next
 End If

End Sub
