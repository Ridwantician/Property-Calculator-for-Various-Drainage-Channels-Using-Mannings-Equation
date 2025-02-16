VERSION 5.00
Begin VB.Form LblTitle 
   BackColor       =   &H80000000&
   Caption         =   "Uniform Flow in Trapezoidal Channel"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbSection 
      Height          =   315
      Left            =   16320
      TabIndex        =   35
      Top             =   8040
      Width           =   4815
   End
   Begin VB.Frame FrmResults 
      Caption         =   "Outputs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8400
      TabIndex        =   22
      Top             =   7440
      Width           =   7575
      Begin VB.TextBox TxtD 
         Height          =   285
         Left            =   5160
         TabIndex        =   34
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox TxtT 
         Height          =   285
         Left            =   5040
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtR 
         Height          =   285
         Left            =   5040
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtP 
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TxtA 
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtQ 
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblD 
         Caption         =   "Hydraulic Depth D (m):"
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label LblT 
         Caption         =   "Top Width T (m):"
         Height          =   375
         Left            =   3480
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LblR 
         Caption         =   "Hydraulic Radius R (m) :"
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LblP 
         Caption         =   "Wetted Per P (m) :"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label LblA 
         Caption         =   "Area A (m^2) :"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label LblQ 
         Caption         =   "Discharge Q (m^3/s:"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FrmInput 
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   5
      Top             =   7320
      Width           =   7815
      Begin VB.TextBox Txtk 
         Height          =   285
         Left            =   4200
         TabIndex        =   40
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TxtTheta 
         Height          =   285
         Left            =   4080
         TabIndex        =   39
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton CmdTerminate 
         Caption         =   "Terminate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   21
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton CmdSolve 
         Caption         =   "SOLVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Txty 
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Txtz 
         Height          =   285
         Left            =   4080
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtSo 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Txtn 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Txtb 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lblk 
         Caption         =   "Shape factor k:"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label LblTheta 
         Caption         =   "Theta (radians):"
         Height          =   495
         Left            =   3000
         TabIndex        =   37
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Lbly 
         Alignment       =   2  'Center
         Caption         =   "Depth y (m):"
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Lblz 
         Caption         =   "Side Slope z:"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbln 
         Caption         =   "Manning n:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label LblSo 
         Caption         =   "Bed Slope So:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Lblb 
         Alignment       =   2  'Center
         Caption         =   "Bottom Width b (m):"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FrmUnit 
      Caption         =   "Unit System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   16440
      TabIndex        =   3
      Top             =   8760
      Width           =   4815
      Begin VB.TextBox TxtSelected 
         Height          =   405
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton CmdInternational 
         Caption         =   "International"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CmdEnglish 
         Caption         =   "English"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblSelected 
         Caption         =   "Selected:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.PictureBox PicVariables 
      Height          =   6015
      Left            =   8040
      Picture         =   "Trapezoidal_Channel.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   9915
      TabIndex        =   2
      Top             =   960
      Width           =   9975
   End
   Begin VB.PictureBox PicObject 
      Height          =   5775
      Left            =   360
      Picture         =   "Trapezoidal_Channel.frx":11F3C
      ScaleHeight     =   5715
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   960
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Note: Fill Bottom width as Diameter incase of a circular section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   18480
      TabIndex        =   41
      Top             =   1920
      Width           =   3900
   End
   Begin VB.Label LblSection 
      Caption         =   "Select your Section Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   16440
      TabIndex        =   36
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Rectangular,Trapezoidal, Triangular, Circular and Parabolic Channel Property Calculator with Manning's Equation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   16335
   End
End
Attribute VB_Name = "LblTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare all the variables

Dim b As Single, y As Single, z As Single, So As Single, A As Single, P As Single, D As Single
Dim T As Single, n As Single, R As Single, Cu As Single, SI As Boolean, Q As Single, k As Single, Theta As Single

Private Sub CmdEnglish_Click()
'Activate Imperial Unit
SI = False
TxtSelected.Text = "E.S"
Cu = 1.486

End Sub

Private Sub CmdInternational_Click()
'Activate Metric Unit
SI = True
TxtSelected.Text = "S.I"
Cu = 1
End Sub

Private Sub CmdClear_Click()
'Clear the text boxes
Txtb.Text = "":     Txtn.Text = ""
TxtSo.Text = "":    Txtz.Text = ""
Txty.Text = "":     TxtQ.Text = ""
TxtA.Text = "":     TxtP.Text = ""
TxtR.Text = "":     TxtT.Text = ""
TxtD.Text = "":     TxtD.Text = ""
Txtk.Text = "":     TxtTheta.Text = ""

End Sub

Private Sub Form_Load()
CmbSection.AddItem "Trapezoidal"
CmbSection.AddItem "Rectangular"
CmbSection.AddItem "Triangular"
CmbSection.AddItem "Fully Filled Circular"
CmbSection.AddItem "Partially Filled Circular"
CmbSection.AddItem "Parabolic"

End Sub
Private Sub getdata()
b = Val(Txtb.Text)
n = Val(Txtn.Text)
So = Val(TxtSo.Text)
z = Val(Txtz.Text)
y = Val(Txty.Text)
k = Val(Txtk.Text)
Theta = Val(TxtTheta.Text)
End Sub
Private Function error() As Integer
If (b <= 0 Or n <= 0 Or So <= 0 Or z <= 0 Or y <= 0) Then
error = 1
End If

End Function
Private Function CalcDischarge() As Single
If CmbSection.Text = "Trapezoidal" Then

A = (b + z * y) * y
P = b + 2 * y * Sqr(1 + z ^ 2)
R = A / P
T = b + 2 * y
D = A / T
Q = (Cu / n) * A * Sqr(So) * R ^ (2 / 3)

PicObject.Picture = LoadPicture("C:\Users\ridoc\OneDrive\Desktop\Visual Basic\Trapezoidal Channel\trapezoidal.jpg")


ElseIf CmbSection.Text = "Fully Filled Circular" Then

A = (22 / 7) * b ^ 2 / 4
P = (22 / 7) * b
R = b / 4
T = b
D = (22 / 7) * b / 4
Q = (Cu / n) * A * Sqr(So) * R ^ (2 / 3)

PicObject.Picture = LoadPicture("C:\Users\ridoc\OneDrive\Desktop\Visual Basic\Trapezoidal Channel\full circular.jpg")


ElseIf CmbSection.Text = "Partially Filled Circular" Then

A = (b ^ 2 / 8) * (2 * Theta - Math.Sin(2 * Theta))
P = b * Theta
R = A / P
T = 2 * R * Math.Sin(Theta)
D = A / T
Q = (Cu / n) * A * Sqr(So) * R ^ (2 / 3)

PicObject.Picture = LoadPicture("C:\Users\ridoc\OneDrive\Desktop\Visual Basic\Trapezoidal Channel\partially full circular.jpg")

ElseIf CmbSection.Text = "Rectangular" Then

A = b * y
P = b + 2 * y
R = A / P
T = b
D = A / T
Q = (Cu / n) * A * Sqr(So) * R ^ (2 / 3)

PicObject.Picture = LoadPicture("C:\Users\ridoc\OneDrive\Desktop\Visual Basic\Trapezoidal Channel\rectangle.jpg")


ElseIf CmbSection.Text = "Triangular" Then
A = z * y ^ 2
P = 2 * y * Sqr(1 + z ^ 2)
R = A / P
T = 2 * z * y
D = A / T
Q = (Cu / n) * A * Sqr(So) * R ^ (2 / 3)

PicObject.Picture = LoadPicture("C:\Users\ridoc\OneDrive\Desktop\Visual Basic\Trapezoidal Channel\triangular.jpg")

ElseIf CmbSection.Text = "Parabolic" Then
A = (4 / 3) * k * y ^ (3 / 2)
P = (8 / 3) * k * y ^ (3 / 2)
R = A / P
T = 2 * Sqr(y / k)
D = A / T
Q = (Cu / n) * A * Sqr(So) * R ^ (2 / 3)

PicObject.Picture = LoadPicture("C:\Users\ridoc\OneDrive\Desktop\Visual Basic\Trapezoidal Channel\parabolic.jpg")


Else
MsgBox ("Select a Drainage Section")
PicObject.Picture = LoadPicture("")


End If
TxtA.Text = A
TxtP.Text = P
TxtR.Text = R
TxtT.Text = T
TxtD.Text = D
TxtQ.Text = Q



End Function

Private Sub CmdSolve_Click()
'Pop up a message when unit system isn't selected
If TxtSelected.Text <> "" Then
Call getdata
Else
MsgBox ("Pick a Unit System")
Exit Sub
End If

'Pop up a message when input data isn't complete
If error = 1 Then
MsgBox ("Some input data is missing")
Call CmdClear_Click
Exit Sub
Else
Q = CalcDischarge()



End If


End Sub

Private Sub CmdTerminate_Click()
'Terminate the program
End
End Sub




