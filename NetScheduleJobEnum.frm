VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListScheduledJobs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List scheduled jobs"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   525
      Left            =   6690
      TabIndex        =   3
      Top             =   2580
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   5310
      TabIndex        =   2
      Top             =   2580
      Width           =   1245
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   8070
      TabIndex        =   0
      Top             =   2580
      Width           =   1245
   End
End
Attribute VB_Name = "frmListScheduledJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Command1_Click()
    Dim lngWin32apiResultCode As Long
    Dim strServerName As String
    Dim lngBufPtr As Long
    Dim lngEntriesRead As Long
    Dim lngTotalEntries As Long
    Dim lngResumeHandle As Long
    Dim udtAtEnum As AT_ENUM
    Dim lngEntry As Long
    
    ' Clearing list view
    ListView1.ListItems.Clear
    
    ' Server name configuration (in case of null local computer)
    strServerName = StrConv("Station1", vbUnicode)
    
    Do
        ' Enumerating the job which is placed in queue
        lngWin32apiResultCode = NetScheduleJobEnum(strServerName, lngBufPtr, MAX_PREFERRED_LENGTH, lngEntriesRead, lngTotalEntries, lngResumeHandle)
        
        ' When  succeeding in the acquisition of user name
        If (lngWin32apiResultCode = NERR_Success) Or (lngWin32apiResultCode = ERROR_MORE_DATA) Then
            For lngEntry = 0 To lngEntriesRead - 1
                ' Buffer in structure copy
                RtlMoveMemory udtAtEnum, ByVal lngBufPtr + Len(udtAtEnum) * lngEntry, Len(udtAtEnum)
                
                ' Displaying the result in list view
                Result2ListView udtAtEnum
            Next
        End If
        
        ' Releasing memory
        If lngBufPtr <> 0 Then
            NetApiBufferFree lngBufPtr
        End If
        
    Loop While lngWin32apiResultCode = ERROR_MORE_DATA
End Sub
' Displaying the member of the structure in list view
Private Sub Result2ListView(udtAtEnum As AT_ENUM)
    Dim lvwUserItem As ListItem
    Dim intCounter As Integer
    Dim strDate As String
    Dim vntWeek As Variant
    
    vntWeek = Array("M", "T", "W", "TH", "F", "S", "SU")
    
    With udtAtEnum
        
        ' Identification number of schedule
        Set lvwUserItem = ListView1.ListItems.Add(, , .Jobid)
        
        
        ' /every and /next switch
        If .Flags And JOB_RUN_PERIODICALLY Then
            strDate = strDate & "every"
        Else
            strDate = strDate & "the next"
        End If
        
        ' Adding in the character string the null blank character lastly
        strDate = strDate & ""
        
        ' The date which runs command
        For intCounter = 1 To 31
            ' When the bit flag is on
            If .DaysOfMonth And 2 ^ (intCounter - 1) Then
                ' Adding the date which corresponds to the character string
                strDate = strDate & intCounter & ""
            End If
        Next
        
        ' The day of the week when command is run
        For intCounter = 0 To 6
            ' When the bit flag is on
            If .DaysOfWeek = 2 ^ (intCounter) Then
                ' Adding the day of the week when it corresponds to the character string
                strDate = strDate & vntWeek(intCounter) & ""
            End If
        Next
        
        ' Displaying day of the week in list view
        lvwUserItem.SubItems(1) = strDate
        
        ' The time which runs command
        lvwUserItem.SubItems(2) = Format(DateAdd("s", .JobTime \ 1000, "00:00"), "hh:Mm")
        
        ' Displaying the command which is run in list view
        lvwUserItem.SubItems(3) = PointerToString(.Command)
        
        ' Error status being configurated, at the time of the ?
        If .Flags And JOB_EXEC_ERROR Then
            ' Displaying error
            lvwUserItem.SubItems(4) = "Error"
        Else
            lvwUserItem.SubItems(4) = ""
        End If
        
        
        
    End With
End Sub
 
Private Sub Command2_Click()
    frmAddScheduledJob.Show 1
    
    ' Refresh the list
    Command1_Click
End Sub

Private Sub Command3_Click()
    Dim lngWin32apiResultCode As Long
    Dim strServerName         As String
    Dim lngJobID              As Long
    Dim strMessage            As String
    
    ' Get the server name
    strServerName = StrConv("", vbUnicode)
    
    ' Get the job id of the selected job
    lngJobID = ListView1.SelectedItem
    
    ' Crreate a message to show the user
    strMessage = "Are you sure you want to delete the Scheduled Job ID" & lngJobID & "?"
    
    ' Confirm that the user wants to delete this job
    If MsgBox(strMessage, vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
        
        ' Delete the job using api
        lngWin32apiResultCode = NetScheduleJobDel(strServerName, lngJobID, lngJobID)
                
        ' Refresh the list
        Command1_Click
    
    End If
    
End Sub


Private Sub Form_Load()
    ' Initializing list view
    With ListView1
        .ColumnHeaders.Add , , "ID"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Command"
        .ColumnHeaders.Add , , "Status"
        .View = lvwReport
    End With
    
    ' Initializing the command button
    Command1.Caption = "List Jobs"
    Command2.Caption = "Add Job"
    Command3.Caption = "Delete Job"
End Sub





