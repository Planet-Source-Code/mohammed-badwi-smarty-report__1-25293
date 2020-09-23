VERSION 5.00
Object = "{16795D44-640A-11D5-8458-D2A0CF80184A}#17.0#0"; "SmartyReport.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smarty Report - ÇáÊÞÇÑíÑ ÇáÐßíÉ - V1.1"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin SmartReport.SmartyReport SmartyReport1 
      Left            =   6510
      Top             =   1815
      _ExtentX        =   1111
      _ExtentY        =   1111
      SQLStatement    =   "select [Au_ID],[Author] from Authors;"
      HTMLFileName    =   "C:\SmartyReport.htm"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export Report to HTML-ÊÕÏíÑ ÇáÊÞÑíÑ ßÕÝÍÉ æíÈ "
      Height          =   630
      Index           =   1
      Left            =   795
      TabIndex        =   4
      Top             =   2940
      Width           =   5670
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Report-ãÔÇåÏÉ ÇáÊÞÑíÑ"
      Height          =   630
      Index           =   0
      Left            =   795
      TabIndex        =   3
      Top             =   2235
      Width           =   5670
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Example for how You Can use Smartry Report Control You must Register Free Send Blank E-Mail to thise Mail"
      Height          =   645
      Left            =   705
      TabIndex        =   2
      Top             =   750
      Width           =   5985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "åÐÇ ãËÇá Úáì ÇÓÊÎÏÇã ÇÏÇÉ ÇáÊÞÇÑíÑ ÇáÐßíÉ æíÌÈ Çä ÊÞæã ÈÊÓÌíá ÇáÇÏÇÉ ÈãÌÑÏ ÇÑÓÇá ÑÓÇáÉ ÝÇÑÛÉ Çáì ÇáÚäæÇä ÇáÈÑíÏí ÇáãæÖÍ ÇÏäÇå"
      Height          =   525
      Left            =   690
      TabIndex        =   1
      Top             =   105
      Width           =   6000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register-SmartyReport@arabteam2000.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   2
      Left            =   990
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1530
      Width           =   5490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)


Dim AppPath As String

    If Right$(App.Path, 1) <> "\" Then
        AppPath = App.Path & "\"
    Else
        AppPath = App.Path
    End If
    
    With SmartyReport1
        .DatabaseName = AppPath & "db.mdb"
        .PicturePath = AppPath & "logo.jpg"
        .SQLStatement = "Select * from Personr;"
        .TextAlignment = ToRight
        .PageSubject = "ExampleãËÇá"
        .PageHeader = "just Try ãÌÑÏ ÊÌÑÈÉ"
        .PageFooter = ""
        Select Case Index
            Case 0
                .ViewReport
            Case 1
                .HTMLFileName = AppPath & "Report.htm"
                .ExpToHTML
        End Select
    End With
    
End Sub


