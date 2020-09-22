Attribute VB_Name = "mPriority"
Option Explicit

'---------------------------------------------------------------------------------------
' Module: Set Process(Application) Priority
' DateTime: 1/16/2002
' Author: Brian G. Schmitt
' Purpose: Used to Set and Retrieve the Priority of Your Processes
' Returns: If the function succeeds, the return value is nonzero.
'              If the function fails, the return value is zero.
' Requirements: Windows NT/2000/XP: Included in Windows NT 3.1 and later.
'                       Windows 95/98/Me: Included in Windows 95 and later.
'Notes: No current support for Above_Normal and Below_Normal
'        For Explanation of the Levels Look Below all Procedures
'---------------------------------------------------------------------------------------

'Some API Declarations
Private Declare Function GetCurrentProcess Lib "kernel32" _
      () As Long
Private Declare Function SetPriorityClass Lib "kernel32" _
      (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" _
      (ByVal hProcess As Long) As Long

'Enumerates the Priority Classes To Appear as dropdown items(list members)
'when calling the procedure
Enum PriorityClass
   REALTIME_PRIORITY_CLASS = &H100
   HIGH_PRIORITY_CLASS = &H80
   NORMAL_PRIORITY_CLASS = &H20
   IDLE_PRIORITY_CLASS = &H40
End Enum

'---------------------------------------------------------------------------------------
' Procedure : SetPriority
' Purpose   : Sets the Priority Level of the Current Program
'---------------------------------------------------------------------------------------
Function SetPriority(PriorityClass As PriorityClass) As Long
   SetPriority = SetPriorityClass(GetCurrentProcess, PriorityClass)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPriority
' Purpose   : Used to Retrieve the Current Priority Class
' Returns : String
'---------------------------------------------------------------------------------------
Function GetPriority() As Long
   GetPriority = (GetPriorityClass(GetCurrentProcess))
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPriorityName
' Purpose   : Returns the Priority Level Name in Place of a Value as above
'---------------------------------------------------------------------------------------
Function GetPriorityName() As String
   
   Dim lngPriority As Long
   lngPriority = GetPriority
   
   Select Case lngPriority
      Case 256
         GetPriorityName = "Realtime"
      Case 128
         GetPriorityName = "High"
      Case 32
         GetPriorityName = "Normal"
      Case 64
         GetPriorityName = "Idle"
   End Select
End Function

'---------------------------------------------------------------------------------------
' RealTimePriority: Specify this class for a process that has the highest possible priority.
'                 The threads of the process preempt the threads of all other processes,
'                 including operating system processes performing important tasks.
'                 For example, a real-time process that executes for more than a very brief interval
'                 can cause disk caches not to flush or cause the mouse to be unresponsive.
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' HighPriority: Specify this class for a process that performs time-critical tasks that must be executed immediately.
'                 The threads of the process preempt the threads of normal or idle priority class processes.
'                 An example is the Task List, which must respond quickly when called by the user,
'                 regardless of the load on the operating system.
'                 Use extreme care when using the high-priority class,
'                 because a high-priority class application can use nearly all available CPU time.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' NormalPriority: Specify this class for a process with no special scheduling needs.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' IdlePriority: Specify this class for a process whose threads run only when the system is idle.
'                 The threads of the process are preempted by the threads of any process running in a higher priority class.
'                 An example is a screen saver.
'---------------------------------------------------------------------------------------
