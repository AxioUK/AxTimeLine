VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "*\AAxTimeLine.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   Begin Proyecto1.AxTimeLine AxTimeLine1 
      Height          =   5625
      Left            =   105
      TabIndex        =   41
      Top             =   105
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   9922
      Enabled         =   -1  'True
      Style           =   0
      BorderColor     =   16744576
      BackColor       =   -2147483633
      BorderWidth     =   1
      CornerCurve     =   10
      Caption1Color   =   4210752
      Caption2Color   =   4210752
      IconForeColor   =   16777215
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1AlignV  =   1
      Caption1AlignH  =   0
      Caption2AlignV  =   1
      Caption2AlignH  =   0
      IconAlignV      =   1
      IconAlignH      =   1
      ActiveSection   =   1
      SectionSpace    =   50
      BorderColorActive=   255
      PointBackColor  =   16761024
      LineDistance    =   10
      LineWidth       =   2
      LineColor       =   16761024
      LineStyle       =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Point"
      Height          =   285
      Left            =   4635
      TabIndex        =   40
      Top             =   45
      Width           =   1170
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7290
      TabIndex        =   39
      Text            =   "70"
      Top             =   1140
      Width           =   390
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5985
      TabIndex        =   38
      Text            =   "50"
      Top             =   1140
      Width           =   390
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3150
      TabIndex        =   35
      Text            =   "1"
      Top             =   885
      Width           =   390
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   4845
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   6
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2940
      Width           =   345
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2670
      Width           =   345
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2400
      Width           =   345
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2130
      Width           =   345
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1860
      Width           =   345
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1590
      Width           =   345
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3225
      ScaleHeight     =   195
      ScaleWidth      =   285
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1320
      Width           =   345
   End
   Begin VB.ListBox List6 
      Height          =   645
      Left            =   5370
      TabIndex        =   17
      Top             =   2490
      Width           =   855
   End
   Begin VB.ListBox List5 
      Height          =   645
      Left            =   7020
      TabIndex        =   16
      Top             =   2490
      Width           =   855
   End
   Begin VB.ListBox List4 
      Height          =   450
      Left            =   6255
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   645
      Left            =   7020
      TabIndex        =   10
      Top             =   1815
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   5370
      TabIndex        =   9
      Top             =   1815
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Time Visible?"
      Height          =   225
      Left            =   4845
      TabIndex        =   8
      Top             =   810
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Date Visible?"
      Height          =   225
      Left            =   4845
      TabIndex        =   7
      Top             =   540
      Width           =   1305
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3150
      TabIndex        =   5
      Text            =   "5"
      Top             =   600
      Width           =   390
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3150
      TabIndex        =   3
      Text            =   "10"
      Top             =   315
      Width           =   390
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   7305
      TabIndex        =   2
      Top             =   345
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3150
      TabIndex        =   0
      Text            =   "0"
      Top             =   30
      Width           =   390
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CornerCurve"
      Height          =   195
      Left            =   4995
      TabIndex        =   37
      Top             =   1230
      Width           =   930
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BorderWidth"
      Height          =   165
      Left            =   3675
      TabIndex        =   36
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LineColor"
      Height          =   195
      Left            =   3660
      TabIndex        =   34
      Top             =   2985
      Width           =   660
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Color"
      Height          =   195
      Left            =   3660
      TabIndex        =   32
      Top             =   2715
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2 Color"
      Height          =   195
      Left            =   3660
      TabIndex        =   30
      Top             =   2445
      Width           =   1065
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1 Color"
      Height          =   195
      Left            =   3660
      TabIndex        =   28
      Top             =   2175
      Width           =   1065
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BorderColorActive"
      Height          =   195
      Left            =   3660
      TabIndex        =   26
      Top             =   1905
      Width           =   1305
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BorderColor"
      Height          =   195
      Left            =   3660
      TabIndex        =   24
      Top             =   1635
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor"
      Height          =   195
      Left            =   3660
      TabIndex        =   22
      Top             =   1365
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seccion"
      Height          =   195
      Left            =   6705
      TabIndex        =   20
      Top             =   1230
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2"
      Height          =   195
      Left            =   6285
      TabIndex        =   19
      Top             =   2745
      Width           =   645
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      Height          =   195
      Left            =   6285
      TabIndex        =   18
      Top             =   2025
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      Height          =   195
      Left            =   6270
      TabIndex        =   15
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TextAlign Vertical"
      Height          =   195
      Left            =   7035
      TabIndex        =   13
      Top             =   1605
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TextAlign Horizontal"
      Height          =   195
      Left            =   5385
      TabIndex        =   12
      Top             =   1605
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TimeLineStyle"
      Height          =   195
      Left            =   7305
      TabIndex        =   11
      Top             =   135
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LineWidth"
      Height          =   195
      Left            =   3675
      TabIndex        =   6
      Top             =   660
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LineDistance"
      Height          =   195
      Left            =   3675
      TabIndex        =   4
      Top             =   375
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActivePoint"
      Height          =   165
      Left            =   3675
      TabIndex        =   1
      Top             =   75
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AxTimeLine1_Click()
Text1.Text = AxTimeLine1.ActiveSection
End Sub

Private Sub Check1_Click()
AxTimeLine1.DateVisible = Check1.Value
End Sub

Private Sub Check2_Click()
AxTimeLine1.TimeVisible = Check2.Value
End Sub

Private Sub Command1_Click()
AxTimeLine1.UpdateTimePoint CLng(Text1.Text), True, "User Fired!", "Incompetent Worker", "ef19"
End Sub

Private Sub Form_Load()
Dim i As Long

For i = 0 To 19
    AxTimeLine1.AddTimePoint GetForename(ntRandom) & " " & GetSurname(), _
                             "get " & GetJobName() & " at " & CStr(DateAdd("d", i, Date$)), _
                             "eec" & IIf(i <= 9, i, i - 10), _
                             Format$(Now, "h:m:s"), _
                             CStr(DateAdd("d", i, Date$)), _
                             True

  'AxTimeLine1.AddTimePoint Caption1, Caption2, IconChar, Time, Date, Visible?
Next i
    AxTimeLine1.ActiveSection = 0
    
List1.AddItem "eLine"
List1.AddItem "eDots"
List1.AddItem "eBoxs"
    
List2.AddItem "eLeft"
List2.AddItem "eCenter"
List2.AddItem "eRight"

List3.AddItem "eTop"
List3.AddItem "eMiddle"
List3.AddItem "eBottom"
    
List4.AddItem "Vertical"
List4.AddItem "Horizontal"
    
List6.AddItem "eLeft"
List6.AddItem "eCenter"
List6.AddItem "eRight"

List5.AddItem "eTop"
List5.AddItem "eMiddle"
List5.AddItem "eBottom"
End Sub

Private Sub Form_Resize()
With AxTimeLine1
  If .Style = pVertical Then
    .Move 100, 100, 2800, Form1.ScaleHeight - 300
  Else
    .Move 100, 3250, Form1.ScaleWidth - 300, 2500
  End If
End With
End Sub

Private Sub List1_Click()
AxTimeLine1.LineStyle = List1.ListIndex
End Sub

Private Sub List2_Click()
AxTimeLine1.Caption1AlignH = List2.ListIndex
End Sub

Private Sub List3_Click()
AxTimeLine1.Caption1AlignV = List3.ListIndex
End Sub

Private Sub List4_Click()
AxTimeLine1.Style = List4.ListIndex
Form_Resize
End Sub

Private Sub List5_Click()
AxTimeLine1.Caption2AlignV = List5.ListIndex
End Sub

Private Sub List6_Click()
AxTimeLine1.Caption2AlignH = List6.ListIndex
End Sub

Private Sub pColor_Click(Index As Integer)
With cDialog
  .ShowColor
  pColor(Index).BackColor = .color
End With
With AxTimeLine1
  .BackColor = pColor(0).BackColor
  .BorderColor = pColor(1).BackColor
  .BorderColorActive = pColor(2).BackColor
  .Caption1Color = pColor(3).BackColor
  .Caption2Color = pColor(4).BackColor
  .IconForeColor = pColor(5).BackColor
  .LineColor = pColor(6).BackColor
End With
End Sub

Private Sub Text1_Change()
On Error Resume Next
AxTimeLine1.ActiveSection = Text1.Text
End Sub

Private Sub Text2_Change()
On Error Resume Next
AxTimeLine1.LineDistance = Text2.Text
End Sub

Private Sub Text3_Change()
On Error Resume Next
AxTimeLine1.LineWidth = Text3.Text
End Sub

Private Sub Text4_Change()
On Error Resume Next
AxTimeLine1.BorderWidth = Text4.Text
End Sub

Private Sub Text5_Change()
On Error Resume Next
AxTimeLine1.CornerCurve = Text5.Text
End Sub

Private Sub Text6_Change()
On Error Resume Next
AxTimeLine1.SectionSpace = Text6.Text

End Sub
