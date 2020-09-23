VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   690
      Width           =   7425
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3420
      Top             =   960
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   3210
      Width           =   7425
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   -30
         Width           =   7515
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   1530
      Top             =   960
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_Load()
       Me.Show
       Timer1.Enabled = True
       
End Sub
Private Sub Frame1_Click()
    Unload Me
End Sub
Private Sub Timer1_Timer()
MousePointer = 11
For i = 0 To Screen.FontCount - 1
    Form1.Combo1.AddItem Screen.Fonts(i)
  Command1.Caption = "Loading Fonts........." & " " & Screen.Fonts(i)
Next
Command1.Caption = "Sorting Fonts List Please Wait....."
Sleep (1500)
Command1.Caption = "Sorting Fonts Completed..."
Sleep (250)
Command1.Caption = "Loading Main Window"
Timer2.Enabled = True
Timer1.Enabled = False
Form1.Combo1.ListIndex = 1
Form1.Text1.Text = 10
MousePointer = 1
End Sub
Private Sub Timer2_Timer()
If i >= Screen.FontCount Then
frmSplash.Hide
Form1.Show
Unload frmSplash
End If
End Sub
