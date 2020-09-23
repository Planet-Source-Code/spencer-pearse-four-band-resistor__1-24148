VERSION 5.00
Begin VB.Form frmFour 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Four Band Resistors - By S. Pearse (TPO Diagnostic Specialist)"
   ClientHeight    =   6570
   ClientLeft      =   3210
   ClientTop       =   2775
   ClientWidth     =   8745
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "Four Band Resistor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Four Band Resistor.frx":0E42
   MousePointer    =   99  'Custom
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6570
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7695
      MouseIcon       =   "Four Band Resistor.frx":114C
      Picture         =   "Four Band Resistor.frx":1456
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5940
      Width           =   1005
   End
   Begin VB.Frame fraBand4 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3030
      Left            =   3870
      TabIndex        =   48
      Top             =   3555
      Width           =   1050
      Begin VB.OptionButton optBnd4Gry 
         BackColor       =   &H00808080&
         Caption         =   "Grey"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   31
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Brn 
         BackColor       =   &H00004080&
         Caption         =   "Brown"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   35
         Top             =   1215
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Red 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   36
         Top             =   1440
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Grn 
         BackColor       =   &H0000C000&
         Caption         =   "Green"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   34
         Top             =   990
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Blu 
         BackColor       =   &H00C00000&
         Caption         =   "Blue"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   33
         Top             =   765
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Vio 
         BackColor       =   &H00FF80FF&
         Caption         =   "Violet"
         Height          =   195
         Left            =   45
         TabIndex        =   32
         Top             =   540
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Gld 
         BackColor       =   &H0000C0C0&
         Caption         =   "Gold"
         Height          =   195
         Left            =   45
         TabIndex        =   37
         Top             =   1665
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Sil 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Silver"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   38
         Top             =   1890
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optBnd4Whi 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NONE"
         Height          =   195
         Left            =   45
         TabIndex        =   39
         Top             =   2115
         Width           =   960
      End
   End
   Begin VB.Frame fraBand3 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2985
      Left            =   2565
      TabIndex        =   47
      Top             =   3555
      Width           =   1050
      Begin VB.OptionButton optBnd3Vio 
         BackColor       =   &H00FF80FF&
         Caption         =   "Violet"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   28
         Top             =   2295
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Gry 
         BackColor       =   &H00808080&
         Caption         =   "Grey"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   29
         Top             =   2520
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Whi 
         BackColor       =   &H00FFFFFF&
         Caption         =   "White"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   30
         Top             =   2745
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Gld 
         BackColor       =   &H0000C0C0&
         Caption         =   "Gold"
         Height          =   195
         Left            =   45
         TabIndex        =   20
         Top             =   495
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Sil 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Silver"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   19
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Blk 
         BackColor       =   &H00000000&
         Caption         =   "Black"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   21
         Top             =   720
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Brn 
         BackColor       =   &H00004080&
         Caption         =   "Brown"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   22
         Top             =   945
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Red 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   23
         Top             =   1170
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Orn 
         BackColor       =   &H000080FF&
         Caption         =   "Orange"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   24
         Top             =   1395
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Yel 
         BackColor       =   &H0000FFFF&
         Caption         =   "Yellow"
         Height          =   195
         Left            =   45
         TabIndex        =   25
         Top             =   1620
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Grn 
         BackColor       =   &H0000C000&
         Caption         =   "Green"
         Height          =   195
         Left            =   45
         TabIndex        =   26
         Top             =   1845
         Width           =   960
      End
      Begin VB.OptionButton optBnd3Blu 
         BackColor       =   &H00C00000&
         Caption         =   "Blue"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   27
         Top             =   2070
         Width           =   960
      End
   End
   Begin VB.Frame fraBand1 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2985
      Left            =   45
      TabIndex        =   46
      Top             =   3555
      Width           =   1050
      Begin VB.OptionButton optBnd1Whi 
         BackColor       =   &H00FFFFFF&
         Caption         =   "White"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   8
         Top             =   2115
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Gry 
         BackColor       =   &H00808080&
         Caption         =   "Grey"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   7
         Top             =   1890
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Vio 
         BackColor       =   &H00FF80FF&
         Caption         =   "Violet"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   1665
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Blu 
         BackColor       =   &H00C00000&
         Caption         =   "Blue"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   5
         Top             =   1440
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Grn 
         BackColor       =   &H0000C000&
         Caption         =   "Green"
         Height          =   195
         Left            =   45
         TabIndex        =   4
         Top             =   1215
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Yel 
         BackColor       =   &H0000FFFF&
         Caption         =   "Yellow"
         Height          =   195
         Left            =   45
         TabIndex        =   3
         Top             =   990
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Orn 
         BackColor       =   &H000080FF&
         Caption         =   "Orange"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   2
         Top             =   765
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Red 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   540
         Width           =   960
      End
      Begin VB.OptionButton optBnd1Brn 
         BackColor       =   &H00004080&
         Caption         =   "Brown"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   0
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame fraBand2 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2985
      Left            =   1305
      TabIndex        =   45
      Top             =   3555
      Width           =   1050
      Begin VB.OptionButton optBnd2Whi 
         BackColor       =   &H00FFFFFF&
         Caption         =   "White"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   18
         Top             =   2340
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Gry 
         BackColor       =   &H00808080&
         Caption         =   "Grey"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   17
         Top             =   2115
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Vio 
         BackColor       =   &H00FF80FF&
         Caption         =   "Violet"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   16
         Top             =   1890
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Blu 
         BackColor       =   &H00C00000&
         Caption         =   "Blue"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   15
         Top             =   1665
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Grn 
         BackColor       =   &H0000C000&
         Caption         =   "Green"
         Height          =   195
         Left            =   45
         TabIndex        =   14
         Top             =   1440
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Yel 
         BackColor       =   &H0000FFFF&
         Caption         =   "Yellow"
         Height          =   195
         Left            =   45
         TabIndex        =   13
         Top             =   1215
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Orn 
         BackColor       =   &H000080FF&
         Caption         =   "Orange"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   12
         Top             =   990
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Red 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   11
         Top             =   765
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Brn 
         BackColor       =   &H00004080&
         Caption         =   "Brown"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   10
         Top             =   540
         Width           =   960
      End
      Begin VB.OptionButton optBnd2Blk 
         BackColor       =   &H00000000&
         Caption         =   "Black"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   4
      X1              =   4455
      X2              =   4455
      Y1              =   3375
      Y2              =   3735
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   3
      X1              =   5175
      X2              =   4455
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   2
      X1              =   5175
      X2              =   5175
      Y1              =   2745
      Y2              =   3330
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   1
      X1              =   5175
      X2              =   5175
      Y1              =   1260
      Y2              =   1575
   End
   Begin VB.Label lblOr 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4950
      TabIndex        =   52
      Top             =   4455
      Width           =   3750
   End
   Begin VB.Label lblAns2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4950
      TabIndex        =   51
      Top             =   4905
      Width           =   3750
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3060
      X2              =   3060
      Y1              =   2745
      Y2              =   2880
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   4005
      X2              =   4005
      Y1              =   2745
      Y2              =   3285
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3060
      X2              =   3960
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3060
      X2              =   3060
      Y1              =   3645
      Y2              =   3285
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   1800
      X2              =   1800
      Y1              =   3825
      Y2              =   3150
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3465
      X2              =   1800
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3510
      X2              =   3510
      Y1              =   2745
      Y2              =   3105
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   540
      X2              =   540
      Y1              =   3915
      Y2              =   2880
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3015
      X2              =   585
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "The Resistor Value is :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4950
      TabIndex        =   50
      Top             =   3555
      Width           =   3750
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4950
      TabIndex        =   49
      Top             =   4005
      Width           =   3750
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   0
      X1              =   5805
      X2              =   5175
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   4005
      X2              =   4005
      Y1              =   855
      Y2              =   1575
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   5805
      X2              =   4005
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3510
      X2              =   3510
      Y1              =   540
      Y2              =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   5805
      X2              =   3510
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   3060
      X2              =   3060
      Y1              =   135
      Y2              =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   5805
      X2              =   3105
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Label lblBnd4 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 4 (Tolerence)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5895
      TabIndex        =   44
      Top             =   1080
      Width           =   2850
   End
   Begin VB.Label lblBnd2 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 2 (Second Digit)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5895
      TabIndex        =   43
      Top             =   360
      Width           =   2850
   End
   Begin VB.Label lblBnd3 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 3 (Multiplier)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5895
      TabIndex        =   42
      Top             =   720
      Width           =   2850
   End
   Begin VB.Label lblBnd1 
      BackColor       =   &H00FF8080&
      Caption         =   "Band 1 (First Digit)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5895
      TabIndex        =   41
      Top             =   0
      Width           =   2850
   End
   Begin VB.Shape shpBnd2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1005
      Left            =   3420
      Top             =   1650
      Width           =   195
   End
   Begin VB.Shape shpBnd3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1005
      Left            =   3915
      Top             =   1650
      Width           =   195
   End
   Begin VB.Shape shpBnd4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1005
      Left            =   5085
      Top             =   1650
      Width           =   195
   End
   Begin VB.Shape shpBnd1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1005
      Left            =   2970
      Top             =   1650
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   5775
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   2220
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   2700
      Top             =   1635
      Width           =   3090
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      X1              =   7500
      X2              =   6300
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      X1              =   2190
      X2              =   990
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Band1Val
Dim Band2Val
Dim Band3Val
Dim Band4Val
Private Sub SUM()
Dim ans
ans = ((Val(Band1Val) & Val(Band2Val)) * Val(Band3Val)) & " Ohms" & " " & Band4Val
lblAns.Caption = ans
SUM2
End Sub
Private Sub SUM2()
'Band 3 (Multiplier) value for lblAns2
Dim V
V = Val(Band3Val)
Select Case V
Case 0.01 'Silver
lblOr.Caption = ""
lblAns2.Caption = ""
Case 0.1 'Gold
lblOr.Caption = ""
lblAns2.Caption = ""
Case 1 'Black
lblOr.Caption = ""
lblAns2.Caption = ""
Case 10 'Brown
lblOr.Caption = ""
lblAns2.Caption = ""
Case 100 'Red
If Val(Band2Val) = 0 Then
lblOr.Caption = "Or"
lblAns2.Caption = Val(Band1Val) & " K " & " Ohms" & " " & Band4Val
Else
lblOr.Caption = "Or"
lblAns2.Caption = Val(Band1Val) & " K " & Val(Band2Val) & " Ohms" & " " & Band4Val
End If
Case 1000 'Orange
lblOr.Caption = "Or"
lblAns2.Caption = (Val(Band1Val) & Val(Band2Val)) & " K " & " Ohms" & " " & Band4Val
Case 10000 'Yellow
lblOr.Caption = "Or"
lblAns2.Caption = ((Val(Band1Val) & Val(Band2Val)) * 10) & " K " & " Ohms" & " " & Band4Val
Case 100000 'Green
If Val(Band2Val) = 0 Then
lblOr.Caption = "Or"
lblAns2.Caption = Val(Band1Val) & " M " & " Ohms" & " " & Band4Val
Else
lblOr.Caption = "Or"
lblAns2.Caption = Val(Band1Val) & " M " & Val(Band2Val) & " Ohms" & " " & Band4Val
End If
Case 1000000 'Blue
lblOr.Caption = "Or"
lblAns2.Caption = (Val(Band1Val) & Val(Band2Val)) & " M " & " Ohms" & " " & Band4Val
Case 10000000
lblOr.Caption = "Or"
lblAns2.Caption = ((Val(Band1Val) & Val(Band2Val)) * 10) & " M " & " Ohms" & " " & Band4Val
Case 100000000
lblOr.Caption = "Or"
lblAns2.Caption = ((Val(Band1Val) & Val(Band2Val)) * 100) & " M " & " Ohms" & " " & Band4Val
Case 1000000000
lblOr.Caption = "Or"
lblAns2.Caption = ((Val(Band1Val) & Val(Band2Val)) * 1000) & " M " & " Ohms" & " " & Band4Val
End Select
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
optBnd1Brn_Click
optBnd2blk_Click
optBnd3blk_Click
optBnd4Sil_Click
End Sub

Private Sub optBnd1Blu_Click()
shpBnd1.FillColor = optBnd1Blu.BackColor
lblBnd1.ForeColor = optBnd1Blu.BackColor
Band1Val = 6
SUM
End Sub

Private Sub optBnd1Brn_Click()
shpBnd1.FillColor = optBnd1Brn.BackColor
lblBnd1.ForeColor = optBnd1Brn.BackColor
Band1Val = 1
SUM
End Sub

Private Sub optBnd1Grn_Click()
shpBnd1.FillColor = optBnd1Grn.BackColor
lblBnd1.ForeColor = optBnd1Grn.BackColor
Band1Val = 5
SUM
End Sub

Private Sub optBnd1Gry_Click()
shpBnd1.FillColor = optBnd1Gry.BackColor
lblBnd1.ForeColor = optBnd1Gry.BackColor
Band1Val = 8
SUM
End Sub

Private Sub optBnd1Orn_Click()
shpBnd1.FillColor = optBnd1Orn.BackColor
lblBnd1.ForeColor = optBnd1Orn.BackColor
Band1Val = 3
SUM
End Sub

Private Sub optBnd1Red_Click()
shpBnd1.FillColor = optBnd1Red.BackColor
lblBnd1.ForeColor = optBnd1Red.BackColor
Band1Val = 2
SUM
End Sub

Private Sub optBnd1Vio_Click()
shpBnd1.FillColor = optBnd1Vio.BackColor
lblBnd1.ForeColor = optBnd1Vio.BackColor
Band1Val = 7
SUM
End Sub

Private Sub optBnd1Whi_Click()
shpBnd1.FillColor = optBnd1Whi.BackColor
lblBnd1.ForeColor = optBnd1Whi.BackColor
Band1Val = 9
SUM
End Sub

Private Sub optBnd1Yel_Click()
shpBnd1.FillColor = optBnd1Yel.BackColor
lblBnd1.ForeColor = optBnd1Yel.BackColor
Band1Val = 4
SUM
End Sub
Private Sub optBnd2blk_Click()
shpBnd2.FillColor = optBnd2Blk.BackColor
lblBnd2.ForeColor = optBnd2Blk.BackColor
Band2Val = 0
SUM
End Sub

Private Sub optBnd2Blu_Click()
shpBnd2.FillColor = optBnd2Blu.BackColor
lblBnd2.ForeColor = optBnd2Blu.BackColor
Band2Val = 6
SUM
End Sub

Private Sub optBnd2Brn_Click()
shpBnd2.FillColor = optBnd2Brn.BackColor
lblBnd2.ForeColor = optBnd2Brn.BackColor
Band2Val = 1
SUM
End Sub

Private Sub optBnd2Grn_Click()
shpBnd2.FillColor = optBnd2Grn.BackColor
lblBnd2.ForeColor = optBnd2Grn.BackColor
Band2Val = 5
SUM
End Sub

Private Sub optBnd2Gry_Click()
shpBnd2.FillColor = optBnd2Gry.BackColor
lblBnd2.ForeColor = optBnd2Gry.BackColor
Band2Val = 8
SUM
End Sub

Private Sub optBnd2Orn_Click()
shpBnd2.FillColor = optBnd2Orn.BackColor
lblBnd2.ForeColor = optBnd2Orn.BackColor
Band2Val = 3
SUM
End Sub

Private Sub optBnd2Red_Click()
shpBnd2.FillColor = optBnd2Red.BackColor
lblBnd2.ForeColor = optBnd2Red.BackColor
Band2Val = 2
SUM
End Sub

Private Sub optBnd2Vio_Click()
shpBnd2.FillColor = optBnd2Vio.BackColor
lblBnd2.ForeColor = optBnd2Vio.BackColor
Band2Val = 7
SUM
End Sub

Private Sub optBnd2Whi_Click()
shpBnd2.FillColor = optBnd2Whi.BackColor
lblBnd2.ForeColor = optBnd2Whi.BackColor
Band2Val = 9
SUM
End Sub

Private Sub optBnd2Yel_Click()
shpBnd2.FillColor = optBnd2Yel.BackColor
lblBnd2.ForeColor = optBnd2Yel.BackColor
Band2Val = 4
SUM
End Sub
Private Sub optBnd3blk_Click()
shpBnd3.FillColor = optBnd3Blk.BackColor
lblBnd3.ForeColor = optBnd3Blk.BackColor
Band3Val = 1
SUM
End Sub

Private Sub optBnd3Blu_Click()
shpBnd3.FillColor = optBnd3Blu.BackColor
lblBnd3.ForeColor = optBnd3Blu.BackColor
Band3Val = 1000000
SUM
End Sub

Private Sub optBnd3Brn_Click()
shpBnd3.FillColor = optBnd3Brn.BackColor
lblBnd3.ForeColor = optBnd3Brn.BackColor
Band3Val = 10
SUM
End Sub

Private Sub optBnd3Gld_Click()
shpBnd3.FillColor = optBnd3Gld.BackColor
lblBnd3.ForeColor = optBnd3Gld.BackColor
Band3Val = 0.1
SUM
End Sub

Private Sub optBnd3Grn_Click()
shpBnd3.FillColor = optBnd3Grn.BackColor
lblBnd3.ForeColor = optBnd3Grn.BackColor
Band3Val = 100000
SUM
End Sub

Private Sub optBnd3Gry_Click()
shpBnd3.FillColor = optBnd2Gry.BackColor
lblBnd3.ForeColor = optBnd3Gry.BackColor
Band3Val = 100000000
SUM
End Sub

Private Sub optBnd3Orn_Click()
shpBnd3.FillColor = optBnd3Orn.BackColor
lblBnd3.ForeColor = optBnd3Orn.BackColor
Band3Val = 1000
SUM
End Sub

Private Sub optBnd3Red_Click()
shpBnd3.FillColor = optBnd3Red.BackColor
lblBnd3.ForeColor = optBnd3Red.BackColor
Band3Val = 100
SUM
End Sub

Private Sub optBnd3Sil_Click()
shpBnd3.FillColor = optBnd3Sil.BackColor
lblBnd3.ForeColor = optBnd3Sil.BackColor
Band3Val = 0.01
SUM
End Sub

Private Sub optBnd3Vio_Click()
shpBnd3.FillColor = optBnd2Vio.BackColor
lblBnd3.ForeColor = optBnd3Vio.BackColor
Band3Val = 10000000
SUM
End Sub

Private Sub optBnd3Whi_Click()
shpBnd3.FillColor = optBnd2Whi.BackColor
lblBnd3.ForeColor = optBnd3Whi.BackColor
Band3Val = 1000000000
SUM
End Sub

Private Sub optBnd3Yel_Click()
shpBnd3.FillColor = optBnd2Yel.BackColor
lblBnd3.ForeColor = optBnd3Yel.BackColor
Band3Val = 10000
SUM
End Sub

Private Sub optBnd4Blu_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Blu.BackColor
lblBnd4.ForeColor = optBnd4Blu.BackColor
Band4Val = "+/-0.25%"
SUM
End Sub

Private Sub optBnd4Brn_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Brn.BackColor
lblBnd4.ForeColor = optBnd4Brn.BackColor
Band4Val = "+/-1%"
SUM
End Sub

Private Sub optBnd4Gld_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Gld.BackColor
lblBnd4.ForeColor = optBnd4Gld.BackColor
Band4Val = "+/-5%"
SUM
End Sub

Private Sub optBnd4Grn_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Grn.BackColor
lblBnd4.ForeColor = optBnd4Grn.BackColor
Band4Val = "+/-0.5%"
SUM
End Sub

Private Sub optBnd4Gry_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Gry.BackColor
lblBnd4.ForeColor = optBnd4Gry.BackColor
Band4Val = "+/-0.05%"
SUM
End Sub

Private Sub optBnd4Red_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Red.BackColor
lblBnd4.ForeColor = optBnd4Red.BackColor
Band4Val = "+/-2%"
SUM
End Sub

Private Sub optBnd4Sil_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Sil.BackColor
lblBnd4.ForeColor = optBnd4Sil.BackColor
Band4Val = "+/-10%"
SUM
End Sub

Private Sub optBnd4Vio_Click()
shpBnd4.Visible = True
lblBnd4.Visible = True
ShowLin1
shpBnd4.FillColor = optBnd4Vio.BackColor
lblBnd4.ForeColor = optBnd4Vio.BackColor
Band4Val = "+/-0.1%"
SUM
End Sub

Private Sub optBnd4Whi_Click()
shpBnd4.Visible = False
lblBnd4.Visible = False
HideLin1
Band4Val = "+/-20%"
SUM
End Sub
Private Sub HideLin1()
Dim i
For i = 0 To 4
lin1(i).Visible = False
Next i
End Sub
Private Sub ShowLin1()
Dim i
For i = 0 To 4
lin1(i).Visible = True
Next i
End Sub
