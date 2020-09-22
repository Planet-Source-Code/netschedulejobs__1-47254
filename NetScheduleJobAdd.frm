VERSION 5.00
Begin VB.Form frmAddScheduledJob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetScheduleJobAdd"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "every"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1260
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "next"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1020
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Text            =   "M,T,W,TH,F,S,SU"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Text            =   "calc.exe"
      Top             =   1650
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "interactive"
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "17:00"
      Top             =   420
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "station1"
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Schedule"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "You can enter in the days that you want it to execute on: M,T,SU or the dates i.e. 13,17,28 of each month."
      ForeColor       =   &H000040C0&
      Height          =   795
      Left            =   90
      TabIndex        =   13
      Top             =   2220
      Width           =   2475
   End
   Begin VB.Label Label5 
      Caption         =   $"NetScheduleJobAdd.frx":0000
      Height          =   855
      Left            =   90
      TabIndex        =   10
      Top             =   3210
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Command"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Time (hh:mm)"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Computer Name"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddScheduledJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim lngWin32apiResultCode As Long
    Dim strComputerName As String
    Dim lngJobID As Long
    Dim udtAtInfo As AT_INFO
    
    ' Convert the computer name to unicode
    strComputerName = StrConv(Text1.Text, vbUnicode)
    
    ' Setup the tasks parameters
    SetStructValue udtAtInfo
    
    ' Schedule the task
    lngWin32apiResultCode = NetScheduleJobAdd(strComputerName, udtAtInfo, lngJobID)
    
    ' Check if the task was scheduled
    If lngWin32apiResultCode = NERR_Success Then
        MsgBox "Task " & lngJobID & " has been scheduled."
    End If
        
    ' Close the form
    Unload frmAddScheduledJob
    Set frmAddScheduledJob = Nothing
    
End Sub
Private Sub SetStructValue(udtAtInfo As AT_INFO)
    Dim strTime As String
    Dim strDate() As String
    Dim vntWeek() As Variant
    Dim intCounter As Integer
    Dim intWeekCounter As Integer
    
    vntWeek = Array("M", "T", "W", "TH", "F", "S", "SU")
    
    With udtAtInfo
        
        ' Change the format of the time
        strTime = Format(Text2.Text, "hh:mm")
        
        ' Change the time to one used by the api
        .JobTime = (Hour(strTime) * 3600 + Minute(strTime) * 60) * 1000
        
        ' Set the Date parameters
        If Val(Text3.Text) > 0 Then
            
            ' Set the task to run on specific days of the month i.e. 9th & 22nd of the month
            strDate = Split(Text3.Text, ",")
            For intCounter = 0 To UBound(strDate)
                .DaysOfMonth = .DaysOfMonth + 2 ^ (strDate(intCounter) - 1)
            Next
        
        Else
            
            ' Set the task to run on sepecific days of the week i.e. Monday & Thursday
            strDate = Split(Text3.Text, ",")
            For intCounter = 0 To UBound(strDate)
                For intWeekCounter = 0 To UBound(vntWeek)
                    If UCase(strDate(intCounter)) = vntWeek(intWeekCounter) Then
                        .DaysOfWeek = .DaysOfWeek + 2 ^ intWeekCounter
                        Exit For
                    End If
                Next
            Next
        End If

        ' Set the interactive property
        If Check1.Value = vbUnchecked Then
            .Flags = .Flags Or JOB_NONINTERACTIVE
        End If
        
        ' Set to run periodically
        If Option2.Value = True Then
            .Flags = .Flags Or JOB_RUN_PERIODICALLY
        End If
        
        ' Set the command to run
        .Command = StrConv(Text4.Text, vbUnicode)
    End With
End Sub

