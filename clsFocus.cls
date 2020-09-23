VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFocus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'Class clsFocus
'---------------------------------------------------------------------
'Author      Manuel Fernandez (mfernan@mail.ono.es)
'            based on submissions by
'            Robert Jackson (Basic focus handling) (Rjackson1_99@yahoo.com)
'            T. De Lange (ArielTimer) (tomdl@attglobal.net)
'
'Applies To  VB 6 (Probably to VB5 also)
'
'Usage       See Form1.frm
'---------------------------------------------------------------------

Option Explicit
Private Gcurrent As Object
Private Gpast As Object
Private mForm As Form
Private WithEvents ArTimer As ArielTimer
Attribute ArTimer.VB_VarHelpID = -1

Public Event ControlValidate(ByRef cancel As Boolean, c As Control)

Public Property Get Form() As Form
    Set Form = mForm
End Property

Public Property Set Form(F As Form)
    Set mForm = F
End Property



Private Sub ChangeControl()
Dim cancel As Boolean
Dim min As Integer: min = 32000
Dim c As Control

If Gcurrent Is Nothing Then
        'find minimum taborder
        For Each c In Form
            On Error Resume Next
            If c.TabIndex < min Then
                Set Gcurrent = c
                min = c.taborder
            End If
            On Error GoTo 0
        Next
        ChangeControl
        'Set Gcurrent = Screen.ActiveControl
Else
    If Not Screen.ActiveControl Is Gcurrent Then
        RaiseEvent ControlValidate(cancel, Gcurrent)
        If Not cancel Then
            Set Gpast = Gcurrent
            Normal Gpast
            Set Gcurrent = Screen.ActiveControl
            Hilite Gcurrent
        End If
        Gcurrent.SetFocus
        
    End If
End If
End Sub

Public Sub Hilite(c As Control)
    Select Case UCase(TypeName(c))
        Case "TEXTBOX"
            c.BackColor = vbBlue
            c.ForeColor = vbYellow
            c.FontBold = True
        Case "COMMANDBUTTON"
            c.FontBold = True
        Case "CHECKBOX", "OPTIONBUTTON"
            c.BackColor = vbButtonFace
            c.ForeColor = vbBlue
            c.FontBold = True
        Case "COMBOBOX", "LISTBOX", "FILELISTBOX", "DRIVELISTBOX", "DIRLISTBOX"
            c.BackColor = vbBlue
            c.ForeColor = vbYellow
            c.FontBold = True
    End Select
End Sub

Public Sub Normal(c As Control)
    Select Case UCase(TypeName(c))
        Case "TEXTBOX"
            c.BackColor = vbWhite
            c.ForeColor = vbBlack
            c.FontBold = False
        Case "COMMANDBUTTON"
            c.FontBold = False
        Case "CHECKBOX", "OPTIONBUTTON"
            c.BackColor = vbButtonFace
            c.ForeColor = vbBlack
            c.FontBold = False
        Case "COMBOBOX", "LISTBOX", "FILELISTBOX", "DRIVELISTBOX", "DIRLISTBOX"
            c.BackColor = vbWhite
            c.ForeColor = vbBlack
            c.FontBold = False
    End Select
End Sub

Public Sub Error(c As Control)
    Select Case UCase(TypeName(c))
        Case "TEXTBOX"
            c.BackColor = vbRed
            c.ForeColor = vbYellow
            c.FontBold = True
        Case "CHECKBOX", "OPTIONBUTTON"
            c.BackColor = vbButtonFace
            c.ForeColor = vbRed
            c.FontBold = True
        Case "COMBOBOX", "LISTBOX", "FILELISTBOX", "DRIVELISTBOX", "DIRLISTBOX"
            c.BackColor = vbRed
            c.ForeColor = vbYellow
            c.FontBold = True
    End Select
End Sub



Private Sub ArTimer_OnTimer()
    If Not Form.ActiveControl Is Nothing Then ChangeControl
End Sub

Public Sub Activate()
    ArTimer.Interval = 10
    ArTimer.Enabled = True
End Sub
Public Sub Deactivate()
    ArTimer.Interval = 0
    ArTimer.Enabled = False
End Sub

Public Sub Terminate()
    ArTimer.Interval = 0
    ArTimer.Enabled = False
    Class_Terminate
End Sub

Public Function ValidateAll() As Boolean
    Dim c As Control
    Dim cancel As Boolean
    
    ValidateAll = True
    
    For Each c In Form
        RaiseEvent ControlValidate(cancel, c)
        If cancel Then ValidateAll = False
    Next
    
End Function

Private Sub Class_Initialize()
    Set ArTimer = New ArielTimer
End Sub

Private Sub Class_Terminate()
    Set ArTimer = Nothing
End Sub


