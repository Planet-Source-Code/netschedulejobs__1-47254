Attribute VB_Name = "basNetScheduleJobs"
Option Explicit

' Declaration of the function which adds a job which is placed in the queue of the computer which is specified
Declare Function NetScheduleJobAdd Lib "netapi32.dll" (ByVal servername As String, Buffer As Any, Jobid As Long) As Long

' Declaration of the function which deletes the job which is placed in the queue of the computer which is specified
Declare Function NetScheduleJobDel Lib "netapi32.dll" (ByVal servername As String, ByVal MinJobId As Long, ByVal MaxJobId As Long) As Long

' Declaration of the function which enumerates the job which is placed in the queue of the computer which is specified
'Declare Function NetScheduleJobEnum Lib "netapi32.dll" (ByVal servername As String, ByVal PointerToBuffer As String, ByVal PrefferedMaximumLength As Long, ByRef entriesread As Long, ByRef totalentries As Long, ByRef resumehandle As Long) As Long
Declare Function NetScheduleJobEnum Lib "netapi32.Dll " (ByVal servername As String, PointerToBuffer As Any, ByVal PreferredMaximumLength As Long, entriesread As Long, totalentries As Long, resumehandle As Long) As Long
 
' Declaration of the function which releases memory
Declare Function NetApiBufferFree Lib "netapi32.Dll " (ByVal Buffer As Long) As Long
 
' Declaration of the function which moves memory
Declare Sub RtlMoveMemory Lib "Kernel32.Dll " (Destination As Any, Source As Any, ByVal Length As Long)
 
' Declaration of the function which copies the character string
Declare Function lstrcpy Lib "Kernel32.Dll " Alias "lstrcpyW" (LpszString1 As Any, LpszString2 As Any) As Long
 
' Declaration of the function which returns the length in the character string
Declare Function lstrlen Lib "Kernel32.Dll " Alias "lstrlenW" (ByVal lpszString As Long) As Long
 

' Schedule structure
Type AT_INFO
    JobTime     As Long
    DaysOfMonth As Long
    DaysOfWeek  As Byte
    Flags       As Byte
    dummy       As Integer
    Command     As String
End Type

' The structure which stores job information
Type AT_ENUM
    Jobid As Long
    JobTime As Long
    DaysOfMonth As Long
    DaysOfWeek As Byte
    Flags As Byte
    dummy As Integer
    Command As Long
End Type
 
' Schedule constants
Public Const JOB_RUN_PERIODICALLY = &H1
Public Const JOB_NONINTERACTIVE = &H10
Public Const NERR_Success = 0
Public Const JOB_EXEC_ERROR = &H2
Public Const MAX_PREFERRED_LENGTH = -1&
Public Const ERROR_MORE_DATA = 234&

' Converting the pointer to the character string
Function PointerToString(lngPointer As Long) As String
    Dim bytBuffer(255) As Byte
    
    ' Copying the character string which the pointer points to byte array
    lstrcpy bytBuffer(0), ByVal lngPointer
    ' Cutoff after the null character
    PointerToString = Left(bytBuffer, lstrlen(lngPointer))
End Function


