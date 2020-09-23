VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "clsFocus Sample"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   240
      Index           =   1
      Left            =   255
      TabIndex        =   9
      Top             =   4695
      Width           =   1605
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   240
      Index           =   0
      Left            =   255
      TabIndex        =   8
      Top             =   4335
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "form1.frx":0000
      Left            =   2580
      List            =   "form1.frx":000D
      TabIndex        =   15
      Top             =   3090
      Width           =   1785
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "form1.frx":002C
      Left            =   2535
      List            =   "form1.frx":0036
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   2535
      Width           =   1830
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   330
      Left            =   2535
      TabIndex        =   13
      Top             =   2055
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   450
      Left            =   2625
      MaskColor       =   &H00FF0000&
      TabIndex        =   16
      Top             =   4665
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   3
      Top             =   1890
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   2
      Top             =   1455
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2310
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2745
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3180
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Alone controls"
      Height          =   375
      Left            =   2820
      TabIndex        =   18
      Top             =   15
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "textbox Array"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author Manuel Fernandez (mfernan@mail.ono.es)
'based on submissions by
'    Robert Jackson (Basic focus handling) (Rjackson1_99@yahoo.com)
'    T. De Lange (ArielTimer) (tomdl@attglobal.net)

Private WithEvents Focus As clsFocus
Attribute Focus.VB_VarHelpID = -1
Option Explicit

Private Sub Form_Load()
    Set Focus = New clsFocus
    Set Focus.Form = Me
End Sub

Private Sub Form_Unload(cancel As Integer)
    Focus.Terminate
    Set Focus = Nothing
End Sub

Private Sub Form_Activate()
    Focus.Activate
End Sub

Private Sub Command1_Click()
    If Focus.ValidateAll Then MsgBox "Valid Data Entry"
End Sub

Private Sub Focus_ControlValidate(cancel As Boolean, c As Control)
    If c Is Text1(1) Then
        If Not IsNumeric(c.Text) Then
            cancel = True
            Focus.Error c
        End If
    End If
    
    If c Is Text1(2) Then
        If c.Text = "" Then
            cancel = True
            Focus.Error c
        End If
    End If
    
End Sub


