VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Workaholic PC v1.0"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ATTACHMENTS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   26
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox moveme 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   330
      Left            =   3480
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000012&
      Caption         =   "Enclosed :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000012&
      Caption         =   "Your Workaholic PC"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000012&
      Caption         =   "Thanking You,"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000012&
      Caption         =   "Shut Down Properly."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "The Number Of Times I Was Not "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "Tells You"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "That Is"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "Thus The Difference Between Them"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "Number of Times "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      Caption         =   "And Shut Down  "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000012&
      Caption         =   "Booted "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "Of"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "For Your Information I have Been "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Till"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Since"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "I Have Been Working For A Period  "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "My Workaholic PC "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
moveme.Text = Val(moveme.Text) + 1
'minutes
If Val(moveme.Text) = 1 Then
Text1.Text = Val(Text1.Text) \ 60
Label7.Caption = "Minutes"
moveme.Text = Val(moveme.Text) + 1
End If
'hours
If Val(moveme.Text) = 3 Then
Text1.Text = Val(Text1.Text) \ 60
Label7.Caption = "Hours"
moveme.Text = Val(moveme.Text) + 1
End If
'days
If Val(moveme.Text) = 5 Then
Text1.Text = Val(Text1.Text) \ 24
Label7.Caption = "Days"
moveme.Text = Val(moveme.Text) + 1
End If
'months
If Val(moveme.Text) = 7 Then
Text1.Text = Val(Text1.Text) \ 31
Label7.Caption = "Months"
moveme.Text = Val(moveme.Text) + 1
End If
'years
If Val(moveme.Text) = 9 Then
Text1.Text = Val(Text1.Text) \ 12
Label7.Caption = "Years"
moveme.Text = Val(moveme.Text) + 1
End If
'seconds
If Val(moveme.Text) = 11 Then
Text1.Text = Val(Text2.Text)
Label7.Caption = "Seconds"
moveme.Text = 0
End If

End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Setdefaults:
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "comtime", App.Path & "\comtime.exe"
If App.PrevInstance Then
End
End If

Open App.Path & "\comtime.cot" For Input As #1
  
    Input #1, checktime
    
    Close #1

Text1.Text = checktime
Text2.Text = Text1.Text
Open App.Path & "\startup.cot" For Input As #1
  
    Input #1, start
    
Close #1
Text3.Text = start

Open App.Path & "\shut.cot" For Input As #1
  
    Input #1, shut
    
Close #1
Text4.Text = shut
Open App.Path & "\stdate.cot" For Input As #1
  
    Input #1, shut1
    
Close #1
Text7.Text = shut1
Text6.Text = Format$(Now, "mmm d,yyyy")
Text5.Text = Val(Text3.Text) - Val(Text4.Text)
End Sub


