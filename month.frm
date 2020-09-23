VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "My Workaholic PC v1.0"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   Icon            =   "month.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   27
      Text            =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text42 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   24
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text31 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text32 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text33 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text34 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text35 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text36 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text37 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text38 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text39 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text40 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text41 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Monthly Report"
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
      Left            =   120
      TabIndex        =   26
      Top             =   0
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
      Left            =   840
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Feb"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Apr"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jun"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jul"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Aug"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sep"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Oct"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nov"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 1 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan / 60
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb / 60
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar / 60
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap / 60
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may / 60
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun / 60
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul / 60
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug / 60
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep / 60
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo / 60
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov / 60
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec / 60
Text1.Text = Val(Text1.Text) + 1
Label7.Caption = "Minutes"
End If

'Hour

If Val(Text1.Text) = 3 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan \ 3600
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb \ 3600
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar \ 3600
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap \ 3600
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may \ 3600
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun \ 3600
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul \ 3600
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug \ 3600
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep \ 3600
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo \ 3600
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov \ 3600
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec \ 3600
Text1.Text = Val(Text1.Text) + 1
Label7.Caption = "Hours"
End If

'Seconds

If Val(Text1.Text) = 5 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec
Text1.Text = 0
Label7.Caption = "Seconds"
End If
End Sub

Private Sub Form_Load()
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec
End Sub
