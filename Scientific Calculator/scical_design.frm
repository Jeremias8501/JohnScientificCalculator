VERSION 5.00
Begin VB.Form Scientific_Calculator 
   BackColor       =   &H00000000&
   Caption         =   "Scientific Calculator"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   5685
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "scical_design.frx":0000
   ScaleHeight     =   6240
   ScaleMode       =   0  'User
   ScaleWidth      =   3301.244
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo5 
      Height          =   405
      Left            =   6000
      TabIndex        =   41
      Text            =   "Metre/second (m/s)"
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
      Height          =   405
      Left            =   6000
      TabIndex        =   39
      Text            =   "Celcius (°C)"
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   405
      Left            =   6000
      TabIndex        =   33
      Text            =   "Kilometre (km)"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ComboBox Combo4 
      Height          =   405
      Left            =   6000
      TabIndex        =   40
      Text            =   "Fahrenheit (°F)"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ComboBox Combo6 
      Height          =   405
      Left            =   6000
      TabIndex        =   42
      Text            =   "Kilometre/hour (km/h)"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H000080FF&
      Caption         =   "="
      Height          =   3255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   405
      Left            =   6000
      TabIndex        =   32
      Text            =   "Centimetre (cm)"
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   38
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   37
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Speed Conversion"
      Height          =   735
      Left            =   6000
      TabIndex        =   36
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Temperature Conversion"
      Height          =   735
      Left            =   6000
      TabIndex        =   35
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Length Conversion"
      Height          =   735
      Left            =   6000
      TabIndex        =   34
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton Add 
      BackColor       =   &H000080FF&
      Caption         =   "+"
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Clear 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "C"
      Height          =   735
      Left            =   4080
      MaskColor       =   &H000080FF&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "7"
      Height          =   735
      Left            =   2880
      TabIndex        =   29
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      Height          =   735
      Left            =   1680
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      Height          =   735
      Left            =   1680
      TabIndex        =   27
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "6"
      Height          =   735
      Left            =   480
      TabIndex        =   26
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "9"
      Height          =   735
      Left            =   480
      TabIndex        =   25
      Top             =   1815
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   24
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Del 
      BackColor       =   &H000000FF&
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "4"
      Height          =   735
      Left            =   2880
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "3"
      Height          =   735
      Left            =   480
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "2"
      Height          =   735
      Left            =   1680
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "1"
      Height          =   735
      Left            =   2880
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Subtrct 
      BackColor       =   &H000080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "0"
      Height          =   735
      Left            =   480
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "."
      Height          =   735
      Left            =   1680
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Dvd 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Mltply 
      BackColor       =   &H000080FF&
      Caption         =   "*"
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H80000010&
      Caption         =   "X^2"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H80000010&
      Caption         =   "X^Y"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H000080FF&
      Caption         =   "="
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H80000010&
      Caption         =   "10^y"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H80000010&
      Caption         =   "Sqrt"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H80000010&
      Caption         =   "Log"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H80000010&
      Caption         =   "Abs"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H80000010&
      Caption         =   "X^3"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H80000010&
      Caption         =   "Tan"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H80000010&
      Caption         =   "Cos"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H80000010&
      Caption         =   "Sin"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin VB.Menu VIEW 
      Caption         =   "VIEW"
      Begin VB.Menu standard 
         Caption         =   "standard"
      End
      Begin VB.Menu scientific 
         Caption         =   "scientific"
      End
      Begin VB.Menu Converter 
         Caption         =   "Unit Converter"
      End
   End
   Begin VB.Menu OFF 
      Caption         =   "OFF"
   End
End
Attribute VB_Name = "Scientific_Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b As Double
Dim opern As Integer
Dim fn As Double
Dim r As Double

Private Sub Add_Click()
opern = 1
Text1.Text = Text2.Text
Text2.Text = ""

End Sub

Private Sub Clear_Click()
If Text3.Visible = False Then
Text2.Text = " "
Text1.Text = " "
ElseIf Text3.Visible = True Then
Text3.Text = " "
Text4.Text = " "
End If
End Sub

Private Sub btn9_Click()
Text2.Text = Text2.Text & Val(btn9.Caption)
End Sub

Private Sub btn2_Click()
Text1.Text = Text1.Text & Val(btn2.Caption)
End Sub

Private Sub btn1_Click()
Text1.Text = Text1.Text & Val(btn1.Caption)
End Sub

Private Sub Command12_Click()
Text1.Text = "                 TEMPERATURE"
Text2.Text = "CONVERSION                "

Text3.Text = ""
Text4.Text = ""
Text3.Visible = True
Text4.Visible = True

Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = True
Combo4.Visible = True
Combo5.Visible = False
Combo6.Visible = False


End Sub

Private Sub btn0_Click()
Text1.Text = Text1.Text & Val(btn0.Caption)
End Sub

Private Sub btndot_Click()
Text1.Text = Text1.Text & Val(btndot.Caption)
End Sub

Private Sub Command15_Click()
Text1.Text = "                         SPEED"
Text2.Text = "CONVERSION                  "

Text3.Text = ""
Text4.Text = ""
Text3.Visible = True
Text4.Visible = True
Text4.Visible = True
Text3.Visible = True

Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = True
Combo6.Visible = True
End Sub

Private Sub Command16_Click()
If Combo3.Visible = True Then
    tempconvert
ElseIf Combo5.Visible = True Then
speedconvert
ElseIf Combo1.Visible = True Then
lengthconvert
End If
End Sub

Private Sub btn8_Click()
Text1.Text = Text1.Text & Val(btn8.Caption)
End Sub

Private Sub Command1_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command1.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command1.Caption)

End If

End Sub

Private Sub Command10_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command10.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command10.Caption)

End If
End Sub

Private Sub Command11_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command11.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command11.Caption)

End If
End Sub

Private Sub Command13_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command13.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command13.Caption)

End If
End Sub

Private Sub Command14_Click()
If Text3.Visible = False Then
Text2.Text = "."
ElseIf Text3.Visible = True Then
Text3.Text = "."
End If
End Sub

Private Sub Command17_Click()
Call sin
End Sub

Private Sub Command18_Click()
Call cos
End Sub

Private Sub Command19_Click()
Call tan
End Sub

Private Sub Command2_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command2.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command2.Caption)

End If
End Sub

Private Sub Command20_Click()
Call deg
End Sub

Private Sub Command21_Click()
Call ln
End Sub

Private Sub Command22_Click()
Call log
End Sub

Private Sub Command23_Click()
Call sqrt
End Sub

Private Sub Command24_Click()
Call teny
End Sub

Private Sub Command25_Click()
opern = 5
Text1.Text = Text2.Text
Text2.Text = ""
End Sub

Private Sub Command26_Click()
Call operation

End Sub

Private Sub btn7_Click()
Text1.Text = Text1.Text & Val(btn.Caption)
End Sub

Private Sub Command4_Click()
If Text3.Visible = False Then
a = Text2.Text
fn = Text2.Text * -1
Text2.Text = fn
ElseIf Text3.Visible = True Then
a = Text3.Text
fn = Text3.Text * -1
Text3.Text = fn
End If
End Sub

Private Sub Command27_Click()
opern = 6
Text1.Text = Text2.Text
Text2.Text = ""
End Sub

Private Sub Command28_Click()
Call x2
End Sub

Private Sub Command3_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command3.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command3.Caption)

End If
End Sub

Private Sub Command5_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command5.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command5.Caption)

End If
End Sub

Private Sub Command6_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command6.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command6.Caption)

End If
End Sub

Private Sub Command7_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command7.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command7.Caption)

End If
End Sub

Private Sub Command8_Click()
Text1.Text = "                        LENGTH"
Text2.Text = " CONVERSION                  "
Text3.Text = ""
Text4.Text = ""
Text3.Visible = True
Text4.Visible = True

Combo1.Visible = True
Combo2.Visible = True
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False
Text3.Refresh
Text4.Refresh

End Sub

Private Sub Command9_Click()
If Text3.Visible = False Then
Text2.Text = Text2.Text & Val(Command9.Caption)

ElseIf Text3.Visible = True Then
Text3.Text = Text3.Text & Val(Command9.Caption)

End If
End Sub

Private Sub Converter_Click()
Me.Width = 9945
Command17.Visible = False
Command21.Visible = False
Command18.Visible = False
Command22.Visible = False
Command19.Visible = False
Command23.Visible = False
Command20.Visible = False
Command24.Visible = False
Command27.Visible = False
Command28.Visible = False
Command8.Visible = True
Command12.Visible = True
Command15.Visible = True
Text3.Visible = True
Text4.Visible = True
Text1.Text = "                        LENGTH"
Text2.Text = " CONVERSION                  "
Combo1.Visible = True
Combo2.Visible = True
Command16.Visible = True

End Sub

Private Sub Del_Click()
Dim dlt, dlte As String, d, f As Integer
If Text3.Visible = False Then
On Error GoTo error_dlt
dlt = Text2.Text
d = Len(dlt)
Text2.Text = Left(dlt, d - 1)
Exit Sub

error_dlt:
Text2.Text = 0

ElseIf Text3.Visible = True Then
On Error GoTo error_dlte
dlte = Text3.Text
f = Len(dlte)
Text3.Text = Left(dlte, f - 1)
Exit Sub

error_dlte:
Text3.Text = 0
End If

End Sub

Private Sub Dvd_Click()
opern = 4
Text1.Text = Text2.Text
Text2.Text = ""
End Sub

Private Sub Form_Load()
Combo1.AddItem "Centimetre (cm)"
Combo1.AddItem "Kilometre (km)"
Combo1.AddItem "Metre (m)"
Combo2.AddItem "Kilometre (km)"
Combo2.AddItem "Centimetre (cm)"
Combo2.AddItem "Metre (m)"

Combo3.AddItem "Celcius (°C)"
Combo3.AddItem "Kelvin (K)"
Combo3.AddItem "Fahrenheit (°F)"
Combo4.AddItem "Celcius (°C)"
Combo4.AddItem "Kelvin (K)"
Combo4.AddItem "Fahrenheit (°F)"

Combo5.AddItem "Metre/second (m/s)"
Combo5.AddItem "Kilometre/hour (km/h)"
Combo5.AddItem "Kilometre/second (km/s)"
Combo6.AddItem "Metre/second (m/s)"
Combo6.AddItem "Kilometre/hour (km/h)"
Combo6.AddItem "Kilometre/second (km/s)"


Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False


Text3.Visible = False
Command16.Visible = False
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Mltply_Click()
opern = 3
Text1.Text = Text2.Text
Text2.Text = ""
End Sub

Private Sub OFF_Click()
Unload Me

End Sub

Private Sub scientific_Click()
Me.Width = 8745
Command16.Visible = False
Command17.Visible = True
Command21.Visible = True
Command18.Visible = True
Command22.Visible = True
Command19.Visible = True
Command23.Visible = True
Command20.Visible = True
Command24.Visible = True
Command27.Visible = True
Command28.Visible = True
Command8.Visible = False
Command12.Visible = False
Command15.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Text3.Visible = False
Text4.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub standard_Click()
Text1.Text = ""
Text2.Text = ""
Text1.Refresh
Text2.Refresh
Me.Width = 5925
Me.Height = 7125
Me.FillColor = &H80FF&
Command17.Visible = False
Command21.Visible = False
Command18.Visible = False
Command22.Visible = False
Command19.Visible = False
Command23.Visible = False
Command20.Visible = False
Command24.Visible = False
Command27.Visible = False
Command28.Visible = False
Command8.Visible = False
Command12.Visible = False
Command15.Visible = False
Command16.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False
End Sub

Private Sub Subtrct_Click()
opern = 2
Text1.Text = Text2.Text
Text2.Text = ""
End Sub

Public Sub operation()
If opern = 1 Then
a = Text1.Text
b = Text2.Text
Text2.Text = a + b
Text1.Text = a & " + " & b

ElseIf opern = 2 Then
a = Text1.Text
b = Text2.Text
Text2.Text = a - b
Text1.Text = a & " - " & b

ElseIf opern = 3 Then
a = Text1.Text
b = Text2.Text
Text2.Text = a * b
Text1.Text = a & " * " & b

ElseIf opern = 4 Then
a = Text1.Text
b = Text2.Text
Text2.Text = a / b
Text1.Text = a & " / " & b

ElseIf opern = 5 Then
a = Text1.Text
b = Text2.Text
Text2.Text = a * b / 100
Text1.Text = a & " % " & b

ElseIf opern = 6 Then
a = Text1.Text
b = Text2.Text
Text2.Text = a ^ b
Text1.Text = a & " ^ " & b
End If



End Sub

Public Sub sin()
a = Text2.Text
fn = Math.sin(Text2.Text * 3.14159265 / 180)
Text2.Text = fn
Text1.Text = "Sin " & a
End Sub

Public Sub cos()
a = Text2.Text
fn = Math.cos(Text2.Text * 1.74532925199433E-02)
Text2.Text = fn
Text1.Text = "Cos " & a
End Sub

Public Sub x2()
a = Text2.Text
fn = a ^ 2
Text1.Text = a & " ^2 "
Text2.Text = fn
End Sub

Public Sub sqrt()
a = Text2.Text
fn = Math.Sqr(Text2.Text)
Text2.Text = fn
Text1.Text = "Sqrt " & a
End Sub

Public Sub teny()
a = Text2.Text
fn = 10 ^ Text2.Text
Text2.Text = fn
Text1.Text = "10^ " & a
End Sub

Public Sub log()
a = Text2.Text
fn = Math.log(Text2.Text)
Text2.Text = fn
Text1.Text = "Log " & a
End Sub

Public Sub ln()
a = Text2.Text
fn = Math.Abs(Text2.Text)
Text2.Text = fn
Text1.Text = "Abs " & a
End Sub

Public Sub tan()
a = Text2.Text
fn = Math.tan(Text2.Text)
Text2.Text = fn
Text1.Text = "Tan " & a
End Sub

Public Sub deg()
a = Text2.Text
fn = a ^ 3
Text2.Text = fn
Text1.Text = a & "^3 "
End Sub

Public Sub tempconvert()
Select Case Combo3.Text
     Case "Celcius (°C)"
        If Combo4.Text = "Fahrenheit (°F)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1.8 + 32
        
        ElseIf Combo4.Text = "Kelvin (K)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text + 273
        
        ElseIf Combo4.Text = "Celcius (°C)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
        
    Case "Fahrenheit (°F)"
        If Combo4.Text = "Celcius (°C)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text - 32 * 0.56

    
        ElseIf Combo4.Text = "Kelvin (K)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text + 459 * 0.56
        
        ElseIf Combo4.Text = "Fahrenheit (°F)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
        
    Case "Kelvin (K)"
        If Combo4.Text = "Celcius (°C)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text - 273

        ElseIf Combo4.Text = "Fahrenheit (°F)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text - 273 * 1.8 + 32
        
        ElseIf Combo4.Text = "Kelvin (K)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
End Select


End Sub

Public Sub speedconvert()

Select Case Combo5.Text
     Case "Metre/second (m/s)"
        If Combo6.Text = "Kilometre/second (km/s)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 0.001
        
        ElseIf Combo6.Text = "Kilometre/hour (km/h)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 3.6
        
        ElseIf Combo6.Text = "Metre/second (m/s)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
        
    Case "Kilometre/second (km/s)"
        If Combo6.Text = "Metre/second (m/s)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1000
    
        ElseIf Combo6.Text = "Kilometre/hour (km/h)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 3600
        
        ElseIf Combo6.Text = "Kilometre/second (km/s)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
        
    Case "Kilometre/hour (km/h)"
        If Combo6.Text = "Metre/second (m/s)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 0.28

        ElseIf Combo6.Text = "Kilometre/second (km/s)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 0.00028
        
        ElseIf Combo6.Text = "Kilometre/hour (km/h)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
End Select

End Sub

Public Sub lengthconvert()

Select Case Combo1.Text
     Case "Centimetre (cm)"
        If Combo2.Text = "Kilometre (km)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 0.0001

        ElseIf Combo2.Text = "Metre (m)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 0.01
        
        ElseIf Combo2.Text = "Centimetre (cm)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
        
    Case "Metre (m)"
        If Combo2.Text = "Centimetre (cm)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 100
    
        ElseIf Combo2.Text = "Kilometre (km)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 0.001
        
        ElseIf Combo2.Text = "Metre (m)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
        
    Case "Kilometre (km)"
        If Combo2.Text = "Centimetre (cm)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 100000

        ElseIf Combo2.Text = "Metre (m)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1000
        
        ElseIf Combo2.Text = "Kilometre (km)" Then
        Text4.Text = " "
        Text4.Text = Text3.Text * 1
        End If
End Select

End Sub



