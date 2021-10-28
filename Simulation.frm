VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Additional material on the book: Markov Chains from theory to implementation and experimentation"
   ClientHeight    =   11160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15990
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   744
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1066
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame startsta 
      Caption         =   "Plot line for:"
      Height          =   1095
      Index           =   1
      Left            =   7920
      TabIndex        =   34
      Top             =   8280
      Width           =   4335
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   38
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   37
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   36
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   35
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   42
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   41
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   39
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plot !"
      Height          =   1575
      Left            =   4320
      TabIndex        =   31
      Top             =   9480
      Width           =   3375
      Begin VB.CommandButton Button 
         Caption         =   "Do it !"
         Height          =   615
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox EraseLast 
         Caption         =   "Erase last plot"
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame startsta 
      Caption         =   "Draws:"
      Height          =   1095
      Index           =   2
      Left            =   4320
      TabIndex        =   27
      Top             =   8280
      Width           =   3375
      Begin VB.TextBox obs 
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Text            =   "1000"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "No. observations:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transition matrix:"
      Height          =   2775
      Left            =   360
      TabIndex        =   2
      Top             =   8280
      Width           =   3735
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   15
         Left            =   2760
         TabIndex        =   18
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   14
         Left            =   2040
         TabIndex        =   17
         Text            =   "1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   13
         Left            =   1320
         TabIndex        =   16
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   12
         Left            =   600
         TabIndex        =   15
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   14
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   10
         Left            =   2040
         TabIndex        =   13
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   12
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   11
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   7
         Left            =   2760
         TabIndex        =   10
         Text            =   "0.333"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   9
         Text            =   "0.333"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   8
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   7
         Text            =   "0.333"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.PictureBox graf_val 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   360
      ScaleHeight     =   503
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1023
      TabIndex        =   1
      Top             =   360
      Width           =   15375
   End
   Begin VB.Label Label5 
      Caption         =   "Frequency:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   28
      Top             =   120
      Width           =   855
   End
   Begin VB.Label ProcentOut 
      Alignment       =   2  'Center
      Caption         =   "A=0% | B=0% | C=0% | D=0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11760
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
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:      Markov Chains: From Theory To Implementation And Experimentation                #
'# Author:    Dr. Paul A. Gagniuc                                                             #
'# Data:      01/09/2016                                                                      #
'# Category:  open source software                                                            #
'#                                                                                            #
'# Description:                                                                               #
'# Simulator for state diagram behavior (used for the main figures in CHAPTER 9)              #
'##############################################################################################

Dim v(0 To 4, 0 To 3) As Variant
Dim p(1 To 4) As Variant

Dim f(0 To 3) As Variant

Private Sub Button_Click()
Button.Enabled = False
Dim Trow(1 To 4) As String

Trow(1) = Val(MCell(0).Text) + Val(MCell(1).Text) + Val(MCell(2).Text) + Val(MCell(3).Text)
Trow(2) = Val(MCell(4).Text) + Val(MCell(5).Text) + Val(MCell(6).Text) + Val(MCell(7).Text)
Trow(3) = Val(MCell(8).Text) + Val(MCell(9).Text) + Val(MCell(10).Text) + Val(MCell(11).Text)
Trow(4) = Val(MCell(12).Text) + Val(MCell(13).Text) + Val(MCell(14).Text) + Val(MCell(15).Text)

For i = 1 To UBound(Trow)
    If Val(Trow(i)) > 0.98 And Val(Trow(i)) <= 1 Then
        ElseIf Val(Trow(i)) = 0 Then
        Else
            MsgBox "The values from row " & i & " of the transition matrix do not" & vbCrLf & "sum up to 1 (or close: ex. 0.99). Check the values from row " & i
            Exit Sub
    End If
Next i

If IsNumeric(obs.Text) = False Then
    MsgBox "Please insert an integer into the No. observations textbox."
    Exit Sub
ElseIf obs.Text <= 1 Then
    MsgBox "Please use at least 2 observations."
    Exit Sub
Else
    obs.Text = Round(Val(obs.Text))
End If


For jar = 1 To 4
    p(jar) = Empty
    f(jar - 1) = Empty
Next jar
Form_Load

Button.Enabled = True
End Sub

Private Sub Form_Load()

If Val(obs.Text) > 100 Then
    Call draw_scale(100)
Else
    draw_scale (Val(obs.Text))
End If

If EraseLast.Value = 1 Then graf_val.Cls

v(0, 0) = "A"
v(0, 1) = "B"
v(0, 2) = "C"
v(0, 3) = "D"

v(1, 0) = Val(MCell(0).Text)
v(1, 1) = Val(MCell(1).Text)
v(1, 2) = Val(MCell(2).Text)
v(1, 3) = Val(MCell(3).Text)

v(2, 0) = Val(MCell(4).Text)
v(2, 1) = Val(MCell(5).Text)
v(2, 2) = Val(MCell(6).Text)
v(2, 3) = Val(MCell(7).Text)

v(3, 0) = Val(MCell(8).Text)
v(3, 1) = Val(MCell(9).Text)
v(3, 2) = Val(MCell(10).Text)
v(3, 3) = Val(MCell(11).Text)

v(4, 0) = Val(MCell(12).Text)
v(4, 1) = Val(MCell(13).Text)
v(4, 2) = Val(MCell(14).Text)
v(4, 3) = Val(MCell(15).Text)

For jar = 1 To 4
    p(jar) = Fill_Jar(jar)
Next jar

draws = Val(obs.Text)
a = Draw(2)

For i = 1 To draws
    For j = 0 To 3
        If a = v(0, j) Then
            a = Draw(j + 1)
            z = z & a
            
            If a = "A" Then f(0) = f(0) + 1
            If a = "B" Then f(1) = f(1) + 1
            If a = "C" Then f(2) = f(2) + 1
            If a = "D" Then f(3) = f(3) + 1

            ori = graf_val.ScaleWidth
            ver = graf_val.ScaleHeight


            tA = (ver / 100) * ((100 / i) * f(0))
            tB = (ver / 100) * ((100 / i) * f(1))
            tC = (ver / 100) * ((100 / i) * f(2))
            tD = (ver / 100) * ((100 / i) * f(3))


            If PlotL(0).Value = 1 Then graf_val.Line (oldn, ver - oldtA)-((ori / draws) * i, ver - tA), vbRed
            If PlotL(1).Value = 1 Then graf_val.Line (oldn, ver - oldtB)-((ori / draws) * i, ver - tB), vbBlue
            If PlotL(2).Value = 1 Then graf_val.Line (oldn, ver - oldtC)-((ori / draws) * i, ver - tC), &H80FF&
            If PlotL(3).Value = 1 Then graf_val.Line (oldn, ver - oldtD)-((ori / draws) * i, ver - tD), vbBlack

            oldn = (ori / draws) * i

            oldtA = tA
            oldtB = tB
            oldtC = tC
            oldtD = tD

            DoEvents

            GoTo 1
        End If
    Next j
1:

Next i


For i = 0 To 3
    pro = pro & v(0, i) & "=" & Int((100 / draws) * f(i)) & "%" & " | "
Next i

ProcentOut.Caption = pro

End Sub

Function Fill_Jar(ByVal S As Variant) As Variant
Ltot = 100
For i = 0 To 3
    a = Int(Ltot * v(S, i))
        For j = 1 To a
            b = b & v(0, i)
        Next j
Next i
Fill_Jar = b
End Function

Function Draw(ByVal S As Variant) As Variant
    Randomize
    randomly_choose = Int(Rnd * Len(p(S)))
    ball = Mid(p(S), randomly_choose + 1, 1)
    Draw = ball
End Function



Function draw_scale(ByVal k_stat As Integer)
Dim zx, qx, zy, qy As Variant
Dim sp As Variant
Dim i As Integer

Form1.Cls

'X axis on graf_val OBJ
'-------------------------------------
sp = graf_val.ScaleWidth / k_stat
For i = 0 To k_stat

    zx = graf_val.Left + (sp * i)
    qx = zx
    zy = graf_val.Top + graf_val.ScaleHeight
    qy = graf_val.Top + graf_val.ScaleHeight + 6

    If k_stat < 10 Then
        Form1.CurrentX = zx - 6
        Form1.CurrentY = qy
        Form1.Print "S" & i
    End If

    Form1.Line (zx, zy)-(qx, qy), &H808080

Next i
'-------------------------------------

'Y axis on graf_val OBJ
'-------------------------------------
    zx = graf_val.Left - 6
    qx = graf_val.Left
    zy = graf_val.Top
    qy = zy
    Form1.Line (zx, zy)-(qx, qy), &H808080
    Form1.CurrentX = zx - 7
    Form1.CurrentY = qy - 6
    Form1.Print "1"

    zx = graf_val.Left - 6
    qx = graf_val.Left
    zy = graf_val.Top + graf_val.ScaleHeight
    qy = zy
    Form1.Line (zx, zy)-(qx, qy), &H808080
    Form1.CurrentX = zx - 7
    Form1.CurrentY = qy - 6
    Form1.Print "0"
'-------------------------------------

End Function
