VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "App Priority"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdGetValue 
         Caption         =   "Current Priority Value"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGetName 
         Caption         =   "Current Priority Name"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame fra 
      Height          =   2655
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdSetHigh 
         Caption         =   "Set High"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetNormal 
         Caption         =   "Set Normal"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetIdle 
         Caption         =   "Set Idle"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetRealtime 
         Caption         =   "Set Realtime"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain   frmMain.frm    Form
' DateTime  : 1/22/2002
' Author    : Brian G. Schmitt
' Purpose   : This Form is to Demonstrate the Use of the Priority Level Module
' Compatible: See the Module for Compatibilites
'---------------------------------------------------------------------------------------
Option Explicit
   Dim x As Long

'---------------------------------------------------------------------------------------
' Procedure : cmdClose_Click
' Purpose   :Unloads the Program
'---------------------------------------------------------------------------------------
Private Sub cmdClose_Click()
   Unload Me
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdGetName_Click
' Purpose   : Displays a MessageBox of the Current Priority Name
'---------------------------------------------------------------------------------------
Private Sub cmdGetName_Click()
   MsgBox GetPriorityName
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdGetValue_Click
' Purpose   : Displays a Message Box to Display the Value of the Current Level
'---------------------------------------------------------------------------------------
Private Sub cmdGetValue_Click()
   MsgBox GetPriority
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSetHigh_Click
' Purpose   : Sets the Priority Level to High
'---------------------------------------------------------------------------------------
Private Sub cmdSetHigh_Click()
   x = SetPriority(HIGH_PRIORITY_CLASS)
   Call Verify
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSetIdle_Click
' Purpose   : Sets Priority Level to Idle--The Lowest Level
'---------------------------------------------------------------------------------------
Private Sub cmdSetIdle_Click()
   x = SetPriority(IDLE_PRIORITY_CLASS)
   Call Verify
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSetNormal_Click
' Purpose   : Sets the Priority Level to Normal
'---------------------------------------------------------------------------------------
Private Sub cmdSetNormal_Click()
   x = SetPriority(NORMAL_PRIORITY_CLASS)
   Call Verify
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSetRealtime_Click
' Purpose   : Sets the Priority Level to Realtime--The Highest Level
'---------------------------------------------------------------------------------------
Private Sub cmdSetRealtime_Click()
   x = SetPriority(REALTIME_PRIORITY_CLASS)
   Call Verify
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Verify
' Purpose   : Used to Display a MessageBox stating a Failure or Success
'                 And to Display the Current Priority Level
'---------------------------------------------------------------------------------------
Private Sub Verify()
Dim lngPriority As Long
   
   lngPriority = GetPriority
   
   If x = 0 Then
      MsgBox "Unable to set Priority!" & vbCrLf & "Current Priority is " & GetPriorityName, vbCritical + vbOKOnly, "Error"
   Else
      MsgBox "Priority Set!" & vbCrLf & "Current Priority is " & GetPriorityName, vbInformation + vbOKOnly, "Success"
   End If
End Sub
