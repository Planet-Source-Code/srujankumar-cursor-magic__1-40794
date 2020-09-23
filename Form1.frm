VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Cursor Magic"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   FillColor       =   &H00808000&
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":09BA
   MousePointer    =   1  'Arrow
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Go Play>>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6990
      TabIndex        =   8
      Top             =   450
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2715
      Left            =   3000
      ScaleHeight     =   2655
      ScaleWidth      =   3645
      TabIndex        =   10
      Top             =   390
      Visible         =   0   'False
      Width           =   3705
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   2100
         Width           =   885
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1620
         Width           =   3375
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   465
         Left            =   2190
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   180
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   4
         Top             =   600
         Width           =   525
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Circle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Square"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font Size :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   15
         Top             =   2190
         Width           =   840
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   1290
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Second Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2790
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "First Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   810
         TabIndex        =   11
         Top             =   600
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1380
      Top             =   3210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "String"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   1
      Top             =   450
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shapes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   900
      TabIndex        =   0
      Top             =   450
      Width           =   885
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1620
      Top             =   2040
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ABCabc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6720
      TabIndex        =   14
      Top             =   2010
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   3870
      TabIndex        =   9
      Top             =   2820
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      Height          =   105
      Index           =   0
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   3930
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos, ypos As Long
Dim clickvalue, xmax, caseval, myval, lfts
Dim objcir, objlab
Dim str, stg, stb
Dim tval As String
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Combo1_Click()
Label5.FontName = Combo1.List(Combo1.ListIndex)
End Sub
Private Sub Combo1_GotFocus()
Label5.Visible = True
End Sub

Private Sub Combo1_LostFocus()
Label5.Visible = False
End Sub

Private Sub Command1_Click()
Option1.Enabled = True
Option2.Enabled = True
Combo1.Enabled = False
Timer1.Enabled = False
Text1.Enabled = False
unloadcir
Unloadlab
Picture1.Visible = True
Command3.Enabled = True
caseval = "cir"
MousePointer = 1
End Sub
Private Sub Command2_Click()
MousePointer = 1
unloadcir
Unloadlab
Combo1.Enabled = True
Text1.Enabled = True
Timer1.Enabled = False
clickvalue = 2
caseval = "lab"
Picture1.Visible = True
Command3.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
End Sub

Private Sub Command3_Click()
Select Case caseval
Case "cir"
    createcir
    clickvalue = 1
    For aaa = 0 To xmax
    Shape1(aaa).Visible = True
    Next
    Timer1.Enabled = True
    Command3.Enabled = False
    Picture1.Visible = False
Case "lab"
    createlab
    Timer1.Enabled = True
    Command3.Enabled = False
    Picture1.Visible = False
End Select
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
lfts = 1
End Sub
Private Sub Form_Resize()
If Form1.WindowState = 0 Then
Form1.WindowState = 2
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure do you want to exit the Application", vbExclamation + vbYesNoCancel) = vbYes Then
Unload Form1
Else
Cancel = 1
End If
End Sub

Private Sub Picture2_Click()
cd1.ShowColor
Picture2.BackColor = cd1.Color
End Sub
Private Sub Picture3_Click()
cd1.ShowColor
Picture3.BackColor = cd1.Color
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Dim a As Integer
MousePointer = 99
Dim sss As POINTAPI
GetCursorPos sss
Select Case clickvalue
Case 1
    Shape1(0).Left = sss.x - 2.5 - Shape1(0).Width / 2
    Shape1(0).Top = sss.y - 22 - Shape1(0).Height / 2
    For aval = 1 To xmax
        xdis = (Shape1(aval - 1).Left - Shape1(aval).Left) / 2 - xdis * 0.5
        ydis = (Shape1(aval - 1).Top - Shape1(aval).Top) / 2 - ydis * 0.5
        Shape1(aval).Left = Shape1(aval).Left + xdis - Shape1(aval).Width / 120
        Shape1(aval).Top = Shape1(aval).Top + ydis - Shape1(aval).Height / 120
    Next
Case 2
    
    Label1(0).Left = sss.x - 2.5
    Label1(0).Top = sss.y - 28
    
    For sval = 1 To Len(tval)
        xdis = (Label1(sval - 1).Left - Label1(sval).Left) / 2 - xdis * 0.7
        ydis = (Label1(sval - 1).Top - Label1(sval).Top) / 2 - ydis * 0.7
        Label1(sval).Left = Label1(sval).Left + xdis + 3
        Label1(sval).Top = Label1(sval).Top + ydis
        
    Next
End Select
End Sub
Public Sub unloadcir()
On Error Resume Next
For y = 1 To objcir
Unload Shape1(y)
Next
End Sub
Public Sub Unloadlab()
On Error Resume Next
For x = 1 To objlab - 1
Unload Label1(x)
Next
End Sub
Public Sub createcir()
Dim r, g, b
Dim i As Integer
xmax = 20
Select Case Option1.Value
    Case 1
            Shape1(0).Shape = 1
    Case Else
        Shape1(0).Shape = 3
End Select
calculate (Picture2.BackColor)
i = 0
startsize = 7
sizeincrement = 2
r = str
g = stg
b = stb
calculate (Picture3.BackColor)
r1 = str
g1 = stg
b1 = stb
Shape1(0).BorderColor = RGB(r, g, b)
For i = 1 To xmax
Load Shape1(i)
Shape1(i).Width = startsize + (i * sizeincrement)
Shape1(i).Height = startsize + (i * sizeincrement)
Shape1(i).Left = (Shape1(i - 1).Left - Shape1(i).Left) / 2
Shape1(i).Top = (Shape1(i - 1).Top - Shape1(i).Top) / 2
Shape1(i).BorderColor = RGB(r, g, b)
If (r1 > r) Then
    r = r + (r1 / xmax)
Else
    r = r - ((r - r1) / xmax)
End If

If (g1 > g) Then
    g = g + (g1 / xmax)
Else
    g = g - ((g - g1) / xmax)
End If
If (b1 > b) Then
    b = b + (b1 / xmax)
Else
    b = b - ((b - b1) / xmax)
End If
Shape1(i).Visible = False
Next
objcir = i
End Sub
Public Sub createlab()
On Error GoTo a
Dim i As Integer
calculate (Picture2.BackColor)
r = str
g = stg
b = stb
calculate (Picture3.BackColor)
r1 = str
g1 = stg
b1 = stb
tval = InputBox("Enter some value in this box", "Enter String", "CREATED & DESIGNED BY SRUJANKUMAR")
Label1(0).ForeColor = RGB(r, g, b)
If Text1.Text = "" Then
Label1(0).FontSize = 10
ElseIf Len(Text1.Text) <= 1 Then
Label1(0).FontSize = 10
Else
Label1(0).FontSize = Text1
End If

Label1(0).FontName = Combo1.List(Combo1.ListIndex)
For ival = 1 To Len(tval)
Load Label1(ival)
Label1(ival).Visible = True
Label1(ival).Caption = Mid(tval, ival + 1, 1)
Label1(ival).ForeColor = RGB(r, g, b)
If (r1 > r) Then
    r = r + (r1 / Len(tval))
Else
    r = r - ((r - r1) / Len(tval))
End If
If (g1 > g) Then
    g = g + (g1 / Len(tval))
Else
    g = g - ((g - g1) / Len(tval))
End If
If (b1 > b) Then
    b = b + (b1 / Len(tval))
Else
    b = b - ((b - b1) / Len(tval))
End If
Next
objlab = ival
Label1(0).Caption = Mid(tval, 1, 1)
Label1(0).Visible = True
Exit Sub
a:
If Err.Number = 13 Then
Text1.Text = 10
End If
End Sub
Public Sub calculate(ByVal ttt)
Dim aaa, first, sec, third, aaa1
aaa1 = Hex(ttt)
For kval = 0 To Len(aaa1) - 1
aaa = aaa + Mid(aaa1, Len(aaa1) - kval, 1)
Next
first = Mid(aaa, 1, 2)
sec = Mid(aaa, 3, 2)
third = Mid(aaa, 5, 2)
str = CHARVAL(first)
stg = CHARVAL(sec)
stb = CHARVAL(third)
End Sub
Public Function CHARVAL(ByVal rrr) As Long
For i = 1 To Len(rrr)
    If Not IsNumeric(Mid(rrr, i, 1)) Then
        chval = Mid(rrr, i, 1)
        If chval = "A" Then
        chval = "10"
        ElseIf chval = "B" Then
        chval = "11"
        ElseIf chval = "C" Then
        chval = "12"
        ElseIf chval = "D" Then
        chval = "13"
        ElseIf chval = "E" Then
        chval = "14"
        ElseIf chval = "F" Then
        chval = "15"
        End If
   Else
    chval = Mid(rrr, i, 1)
   End If
   If (i <= 1) Then
        chval1 = chval
   Else
        chval2 = chval
   End If

Next
CHARVAL = (Val(chval2) * 16 ^ 1 + Val(chval1) * 16 ^ 0)
End Function
