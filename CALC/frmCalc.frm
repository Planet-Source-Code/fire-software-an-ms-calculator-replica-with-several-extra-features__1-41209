VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator Plus"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCE 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   540
      Width           =   570
   End
   Begin VB.CommandButton cmdMemClear 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Clear Memory"
      Top             =   1020
      Width           =   570
   End
   Begin VB.CommandButton cmdMR 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Recall Memory"
      Top             =   1500
      Width           =   570
   End
   Begin VB.CommandButton cmdMS 
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Save Current Operation to Memory"
      Top             =   1980
      Width           =   570
   End
   Begin VB.CommandButton cmdMPlus 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Add current operation result to memory"
      Top             =   2460
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "0"
      Height          =   405
      Index           =   0
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2460
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "1"
      Height          =   405
      Index           =   1
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1980
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "2"
      Height          =   405
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1980
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "3"
      Height          =   405
      Index           =   3
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1980
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "4"
      Height          =   405
      Index           =   4
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1500
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "5"
      Height          =   405
      Index           =   5
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1500
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "6"
      Height          =   405
      Index           =   6
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1500
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "7"
      Height          =   405
      Index           =   7
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1020
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "8"
      Height          =   405
      Index           =   8
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1020
      Width           =   570
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "9"
      Height          =   405
      Index           =   9
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1020
      Width           =   570
   End
   Begin VB.CommandButton cmdInv 
      Caption         =   "+ / -"
      Height          =   405
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2460
      Width           =   570
   End
   Begin VB.CommandButton cmdPt 
      Caption         =   "."
      Height          =   405
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2460
      Width           =   570
   End
   Begin VB.CommandButton cmdDiv 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1020
      Width           =   570
   End
   Begin VB.CommandButton cmdMult 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1500
      Width           =   570
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1980
      Width           =   570
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2460
      Width           =   570
   End
   Begin VB.CommandButton cmdSqrt 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1020
      Width           =   570
   End
   Begin VB.CommandButton cmdPct 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1500
      Width           =   570
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1980
      Width           =   570
   End
   Begin VB.CommandButton cmdSolve 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2460
      Width           =   570
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   540
      Width           =   570
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox txtCalc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   60
      Width           =   3855
   End
   Begin VB.Label lblLastSolved 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      MouseIcon       =   "frmCalc.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   29
      ToolTipText     =   "Last operation - Click to edit equation"
      Top             =   2970
      Width           =   2775
   End
   Begin VB.Label lblLast 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Last Solved:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2970
      Width           =   1035
   End
   Begin VB.Label lblMem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   435
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Calculator"
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================
' = Copyright 2002 - Joe Jordan
' = Portions (button forecolor) provided by www.vbthunder.com
' ==============================================================

' ==============================================================
' About This Project:  This project was an assignment in my
' Visual Basic 6.0 class at the University of South Florida.
' Basically, the assignment was to duplicate the basic MS calc.

' It features many beginner VB techniques including resizable
' arrays and string parsing.
' ==============================================================

' ==============================================================
' Features of This Project:
' I had a little extra time to work on the project, so I went
' beyond the original MS functionality.  The calculator performs
' multiple operations on one line, and also follows traditional
' order of operations (multiply and divide before add and subtract.)
'
' Also a "last equation entered" feature was included so you can
' see and optionally edit what you just entered.
' ==============================================================

Private MemVal As Double        ' Stores value in memory
Private Ops() As String         ' Used in solving (holds operands)
Private Vals() As Double        ' Used in solving (holds values)
Private CopyContent As String   ' Stores copy command info.
Private SolveReset As Boolean   ' Reset after solving

Private Enum CalcType
    NIL         '= 0
    Add         '= 1
    Subtract    '= 2
    Multiply    '= 3
    Divide      '= 4
    Point       '= 5
End Enum

Private Sub cmdAdd_Click()
    Call Append(, Add)
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdBack_Click()
    txtCalc.Text = RemoveLast(txtCalc.Text)
End Sub

Private Sub cmdBack_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdCE_Click()
Dim LastNum As String

' This clears the last set of numbers input
' Exits if the last char in txtcalc is an operator

    If IsOperator(Right(txtCalc.Text, 1)) Then
        ' No Current Number
        Exit Sub
    Else

        ' Remove Last Number
        LastNum = Trim(LastNumber(txtCalc.Text))
        If Mid(LastNum, 1, 1) = "." Then
            LastNum = "0" & LastNum
        End If
        
        txtCalc.Text = Mid(txtCalc.Text, 1, Len(txtCalc.Text) - Len(LastNum))
        If txtCalc.Text = "" Then
            txtCalc.Text = "0"
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    txtCalc.Text = "0"
End Sub

Private Sub cmdClear_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdDiv_Click()
    Call Append(, Divide)
End Sub

Private Sub cmdDiv_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdInv_Click()
Dim CurrentNum As String

' This changes the last number to + or - depending on its current state

    If Len(txtCalc.Text) > 1 Then

        If IsOperator(Mid(txtCalc.Text, Len(txtCalc.Text), 1)) = False Then
        
            CurrentNum = LastNumber(txtCalc.Text)
            If IsNegative(CurrentNum) = False Then
                If Val(CurrentNum) <> 0 Then
                    ' Remove current number
                    txtCalc.Text = Mid(txtCalc.Text, 1, Len(txtCalc.Text) - Len(CurrentNum))
                    txtCalc.Text = txtCalc.Text & "-" & CurrentNum
                End If
            Else
                txtCalc.Text = Mid(txtCalc.Text, 1, Len(txtCalc.Text) - Len(CurrentNum))
                txtCalc.Text = txtCalc.Text & Mid(CurrentNum, 2, Len(CurrentNum))
            End If
            
            'txtCalc.Text = txtCalc.Text
            
        End If
    Else
        If txtCalc.Text <> "0" Then
            CurrentNum = LastNumber(txtCalc.Text)
            If IsNegative(CurrentNum) = False Then
                txtCalc.Text = "-" & txtCalc.Text
            Else
                txtCalc.Text = Abs(Val(txtCalc.Text))
            End If
        End If
        
    End If
End Sub

Private Sub cmdInvert_Click()
Dim Result As Double
    Result = Solve(txtCalc.Text)
    If Result = 0 Then
        MsgBox "Cannot divide by zero.", vbInformation, "Divide By Zero"
        Exit Sub
    Else
        lblLastSolved.Caption = " Inverse ( " & txtCalc.Text & " )"
        txtCalc.Text = 1 / Result
    End If
End Sub

Private Sub cmdInv_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub Form_Load()
    Call ColorButtons
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLastSolved.BackColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UNColorButtons
End Sub

Private Sub lblLastSolved_Click()
    If InStr(lblLastSolved.Caption, "(") = 0 And lblLastSolved.Caption <> "" Then
        ' Isn't a Sqrt or other function
        txtCalc.Text = Trim(lblLastSolved.Caption)
    End If
End Sub

Private Sub lblLastSolved_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLastSolved.BackColor = &HFFFF80
End Sub

Private Sub txtCalc_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdInvert_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdMemClear_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdMPlus_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdMR_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdMS_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdCE_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdMult_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdNum_KeyPress(Index As Integer, KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdMPlus_Click()
    ' Solve current equation and add it to current memory value
    MemVal = MemVal + Solve(txtCalc.Text, True)
    lblMem.Caption = "M"
End Sub

Private Sub cmdMS_Click()
    ' Save current equation value to memory
    MemVal = Solve(txtCalc.Text, True)
    lblMem.Caption = "M"
End Sub

Private Sub cmdPct_Click()
    lblLastSolved.Caption = " Percent ( " & txtCalc.Text & " )"
    txtCalc.Text = Solve(txtCalc.Text) / 100
End Sub

Private Sub cmdPct_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdPt_Click()
    Call Append(, Point)
End Sub

Private Sub cmdPt_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdSolve_Click()
    lblLastSolved.Caption = " " & txtCalc.Text
    txtCalc.Text = Solve(txtCalc.Text)
End Sub

Private Sub cmdSolve_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdSqrt_Click()
Dim Solution As Double
    Solution = Solve(txtCalc.Text)
    If IsNegative(Trim(Str(Solution))) = True Then
        MsgBox "Cannot take the square root of a negative number.", vbInformation, "Negative Square Root"
        Exit Sub
    Else
        lblLastSolved.Caption = " Square Root ( " & txtCalc.Text & " )"
        txtCalc.Text = Sqr(Solution)
    End If
End Sub

Private Sub cmdSqrt_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub cmdSubtract_Click()
    Call Append(, Subtract)
End Sub

Private Sub cmdMemClear_Click()
    lblMem.Caption = ""
    MemVal = 0
End Sub

Private Sub cmdMR_Click()
Dim LastNum As String
    ' Set memory value to equation
    If IsOperator(Right(txtCalc.Text, 1)) Then
        txtCalc.Text = txtCalc.Text & MemVal
    Else
        LastNum = LastNumber(txtCalc.Text)
        txtCalc.Text = Mid(txtCalc.Text, 1, Len(txtCalc.Text) - Len(Trim(Str(LastNum)))) & MemVal
    End If
End Sub

Private Sub cmdMult_Click()
    Call Append(, Multiply)
End Sub

Private Sub cmdNum_Click(Index As Integer)
Dim NumStr As String
    NumStr = Trim(Str(Index))
    
    Call Append(NumStr)
    
End Sub

Private Sub Append(Optional Val As String = "", Optional CType As CalcType = NIL)

' This sub manages and error detects all user input

    If SolveReset = True Then
        If CType = NIL Or InStr(txtCalc.Text, "E") Then ' Number press after solve
            txtCalc.Text = "0"
            SolveReset = False
        Else
            SolveReset = False
        End If
    End If
    
    ' Append Numbers
    If Val <> "" Then
        ' Remove initial zero
        If txtCalc.Text = "0" Then
            txtCalc.Text = ""
        End If
        
        If ValidLastOp = False And Val = "0" Then
            ' append point to 0
            txtCalc.Text = txtCalc.Text & Val & "."
        Else
            ' append your number
            txtCalc.Text = txtCalc.Text & Val
        End If
        
        Exit Sub
        
    End If
    
    
    ' Check for calculation type
    If CType = NIL Then
        Exit Sub
    End If
    
    ' Check for Empty calculation
    If txtCalc.Text = "0" Then
        If CType <> Point Then
            Exit Sub
        End If
    End If
    
    ' Check for valid last op
    If ValidLastOp = False Then
        ' Replace last calculation type with new one
        If txtCalc.Text <> "0" And CType <> Point Then
            txtCalc.Text = RemoveLast(txtCalc.Text)
        Else
            If CType <> Point Then ' Decimal right after operator creates a "0."
                Exit Sub
            Else
                txtCalc.Text = txtCalc.Text & "0."
                Exit Sub
            End If
        End If
    End If
    
    ' Trim zeros from last number
    If txtCalc.Text <> "0" Then
        Dim LastNum As String
        ' Only remove zeros if there's a decimal
        If DecimalExists = True Then
            LastNum = LastNumber(txtCalc.Text)
            txtCalc.Text = Mid(txtCalc.Text, 1, Len(txtCalc.Text) - Len(LastNum))
            txtCalc.Text = txtCalc.Text & TrimZeros(LastNum)
        End If
    End If
    
    ' Check for Point right before next op and remove it
    If Mid(txtCalc.Text, Len(txtCalc.Text), 1) = "." Then
        txtCalc.Text = RemoveLast(txtCalc.Text)
    End If
    
    ' Enter your New Operation
    Select Case CType
    
        Case NIL
            Exit Sub
        Case Add
            txtCalc.Text = txtCalc.Text & "+"
        Case Subtract
            txtCalc.Text = txtCalc.Text & "-"
        Case Multiply
            txtCalc.Text = txtCalc.Text & "*"
        Case Divide
            txtCalc.Text = txtCalc.Text & "/"
        Case Point
            If DecimalExists = False Then
                txtCalc.Text = txtCalc.Text & "."
            End If
            
    End Select
    
End Sub

Private Function DecimalExists() As Boolean
Dim i As Long

' Checks last number to see if there is already a decimal
' ( used in Append when attempting to append a decimal )

    If InStr(LastNumber(txtCalc.Text), ".") <> 0 Then
        DecimalExists = True
    Else
        DecimalExists = False
    End If
    
End Function

Private Function IsOperator(CalcItem As String) As Boolean

' Checks to see if the passed item is an operator
    
    If CalcItem = "+" Or CalcItem = "-" Or CalcItem = "*" Or CalcItem = "/" Then
        IsOperator = True
    Else
        IsOperator = False
    End If
    
End Function

Private Function TrimZeros(sNumber As String) As String
Dim i As Long, tmpNum As String
' This trims the zeros off any number
    tmpNum = sNumber
    For i = Len(tmpNum) To 1 Step -1
        If Right(tmpNum, 1) = "0" Then
            tmpNum = Left(tmpNum, Len(tmpNum) - 1)
        End If
    Next
    
    TrimZeros = tmpNum
    
End Function

Private Function IsNegative(Equation As String) As Boolean
Dim firstChar As String

' Checks to see if passed number/equation is negative

    firstChar = Mid(Equation, 1, 1)
    
    If firstChar = "-" Then
        IsNegative = True
    Else
        IsNegative = False
    End If
    
End Function

Private Sub FormPress(KeyAscii As Integer)
    ' Handles all KeyPress events when buttons have focus
    Call Form_KeyPress(KeyAscii)
End Sub

Private Function LastNumber(Equation As String) As String
Dim i As Long, LastNum As String

' This returns the last number operation
' (anything before the last operator)
' -- takes special account for the - sign in negative numbers
'On Error Resume Next

    For i = Len(Equation) To 1 Step -1
        If IsOperator(Mid(Equation, i, 1)) = True Then
            If i > 1 Then ' Not beginning of formula
                If IsOperator(Mid(Equation, i - 1, 1)) = True Then
                    ' Include negative sign in number
                    LastNum = Mid(Equation, i, 1) & LastNum
                End If
                Exit For
            
            Else ' First Char of Equation is an operator
                LastNum = Mid(Equation, i, 1) & LastNum
            End If
        Else
            LastNum = Mid(Equation, i, 1) & LastNum
        End If
    Next
    
    LastNumber = LastNum
    
End Function

Private Function RemoveLast(Equation As String) As String
Dim LastNum As String
Dim LastNum2 As String

' This removes the last item from txtCalc
' Mostly used for Backspace but also called in SOLVE
' -- Takes special account for negative numbers within a multifunction equation
    
    If Equation <> "0" Then
    
        
        If Len(Equation) > 1 Then
            ' Remove last
            Equation = Left(Equation, Len(Equation) - 1)
            
            If Len(Equation) > 1 Then
                ' Now, see if we're deleting the last of a negative number
                LastNum = Mid(Equation, Len(Equation), 1)
                If IsOperator(LastNum) Then
                    If Len(Equation) > 2 Then
                        LastNum2 = Mid(Equation, Len(Equation) - 1, 1)
                        If IsOperator(LastNum2) Then
                            ' We just deleted the last char of a negative number,
                            ' remove the negative sign
                            Equation = Left(Equation, Len(Equation) - 1)
                        End If
                    End If
                End If
            End If
        
        Else
        
            Equation = "0"
            
        End If
        
    End If
        
    
    RemoveLast = Equation
        
End Function

Private Function ValidLastOp() As Boolean
Dim LastOp As String

' Checks to see if last entry was a number
' If not, returns false else returns true

    If txtCalc.Text = "" Then
        ValidLastOp = False
        Exit Function
    End If
    
    LastOp = Right(txtCalc.Text, 1)
    
    If IsNumeric(LastOp) = True Or LastOp = "." Then
        ValidLastOp = True
    Else
        ValidLastOp = False
    End If
    
End Function

Private Function OpExists(Op As String) As Boolean
Dim i As Long, tmpOp As String
' This is used in SOLVE to see if a certain operator (* / etc.)
' exists in the Ops array

    tmpOp = Op

    For i = 1 To UBound(Ops)
        If Ops(i) = tmpOp Then
            OpExists = True
            Exit Function
        End If
    Next
    
    OpExists = False
    
End Function

Private Function ResizeAndReplace(OpPosition As Long, NewVal As Double)
Dim i As Long

' Resizes calculation arrays with new values
' (solves all / and * before + and - )

    For i = OpPosition To UBound(Ops) - 1
        Ops(i) = Ops(i + 1)
    Next
    
    ReDim Preserve Ops(UBound(Ops) - 1)
    Vals(OpPosition) = NewVal
    
    For i = OpPosition + 1 To UBound(Vals) - 1
        Vals(i) = Vals(i + 1)
    Next
    
    ReDim Preserve Vals(UBound(Vals) - 1)
    
End Function

Private Function Solve(Equation As String, Optional SilentSolve As Boolean = False) As Double
Dim Equ As String
Dim LastNum As Double
Dim LastOp As String
Dim TempNum As Double
Dim tmpLastNum As String
Dim TempOp As String
Dim strLastNum As String
Dim i As Long

' This is the main solving function
' SilentSolve is for things such as calculating M+

' Init. Arrays
ReDim Vals(0)
ReDim Ops(0)

    ' Remove if last item is an operator
    If IsOperator(Right(Equation, 1)) = True Then
        txtCalc.Text = RemoveLast(txtCalc.Text)
    End If
    
    ' Trim zeros from last number
    If Equation <> "0" Then
        tmpLastNum = LastNumber(Equation)
        
        ' Only remove zeros if there's a decimal
        
        If DecimalExists = True Then
            Equation = Mid(Equation, 1, Len(Equation) - Len(tmpLastNum))
            Equation = Equation & TrimZeros(tmpLastNum)
            txtCalc.Text = Equation
        End If
        
    End If
    
    ' Remove if last item is a decimal point
    If Mid(Equation, Len(Equation), 1) = "." Then
        txtCalc.Text = RemoveLast(txtCalc.Text)
    End If
    
    If InStr(txtCalc.Text, "E") Then ' Exponent - Reset Calculator
        txtCalc.Text = "0"
        Exit Function
    End If
        
    Equ = txtCalc.Text
    
    ' ==============================================
    ' Parse Equation into seperate arrays
    ' One array for values, the other for operators
    ' ==============================================
    
    Do While Equ <> ""
    
        If Equ <> "" Then
            ' Get Last Number, store in array
            LastNum = Val(LastNumber(Equ))
            ReDim Preserve Vals(UBound(Vals) + 1)
            Vals(UBound(Vals)) = LastNum
            
            If Len(Equ) - Len(Str(LastNum)) > 0 Then
                strLastNum = Trim(Str(LastNum))
                ' Check for "0.x" convert to string
                ' ( 0. decimals ignore the first zero when converting to a string var )
                If Mid(strLastNum, 1, 1) = "." Then
                    strLastNum = "0" & strLastNum
                End If
                
                ' Subtracting Len of LastNumber rather than LastNum
                ' Because huge numbers turn into Exponents, and decimals
                ' are cut off after a certain degree of accuracy
                
                Equ = Mid(Equ, 1, Len(Equ) - Len(LastNumber(Equ)))
                
            Else
                Equ = ""
            End If
            
        End If
        
        If Len(Equ) > 1 Then
            ' Get Last Operator, store in array
            LastOp = Right(Equ, 1)
            ReDim Preserve Ops(UBound(Ops) + 1)
            Ops(UBound(Ops)) = LastOp
            
            If Len(Equ) - 1 > 0 Then
                Equ = Mid(Equ, 1, Len(Equ) - 1)
            Else
                Equ = ""
            End If
            
        End If
        
    Loop
    
    
    '================================================
    ' Parse of Equation Complete -
    ' Follow standard order of operations to solve
    '================================================
    
    
    ' Solve all * and / first
    For i = UBound(Ops) To 0 Step -1
    
        ' Check to see if / or * is in there
        If OpExists("/") = False And OpExists("*") = False Then
            Exit For
        End If
        
        ' Solve /
        If Ops(i) = "/" Then
            ' Check for Divide by Zero
            If Vals(i) <> 0 Then
                TempNum = Vals(i + 1) / Vals(i)
            Else
                MsgBox "Cannot divide by zero.", vbInformation, "Divide By Zero"
                Exit Function
            End If
            
            Call ResizeAndReplace(i, TempNum)
            i = UBound(Ops)
        End If
        
        
        If UBound(Ops) = 0 Then ' Done Solving
            GoTo Solved
        ElseIf OpExists("*") = False Then
            ' Continue Loop
            If Ops(i) <> "/" Then ' Wasnt first operation
                ' Do Nothing so we can find the rest of * and /
            Else
                i = UBound(Ops) + 1
            End If
            
        Else
            ' Solve *
            If Ops(i) = "*" Then
                TempNum = Vals(i + 1) * Vals(i)
                Call ResizeAndReplace(i, TempNum)
                i = UBound(Ops) + 1
            End If
        End If
        
    Next
    
    
    ' Compute any remaining + or -
    For i = UBound(Ops) To 0 Step -1
    
        ' Check to see if / or * is in there
        If OpExists("-") = False And OpExists("+") = False Then
            Exit For
        End If
    
        If Ops(i) = "-" Then
            TempNum = Vals(i + 1) - Vals(i)
            Call ResizeAndReplace(i, TempNum)
            i = UBound(Ops)
        End If
        
        If UBound(Ops) = 0 Then ' Done Solving
        
            GoTo Solved
            
        ElseIf OpExists("+") = False Then
        
            If Ops(i) <> "-" Then ' Wasnt first operation
                ' Do Nothing so we can find the rest of + and -
            Else
                i = UBound(Ops) + 1
            End If
            
        Else
            If Ops(i) = "+" Then
                TempNum = Vals(i + 1) + Vals(i)
                Call ResizeAndReplace(i, TempNum)
                i = UBound(Ops) + 1
            End If
        End If
        
        ' Start from Beginning
        i = UBound(Ops) + 1
        
    Next
    
    
    
Solved:
    ' Check for solve reset
    If SilentSolve = False Then
        SolveReset = True
    End If
    
    Solve = Vals(1)
    
End Function

Private Sub cmdSubtract_KeyPress(KeyAscii As Integer)
    Call FormPress(KeyAscii)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim Char As String
    Char = LCase(Chr(KeyAscii))
    
    ' This Sub manages all keypress functions
    If IsNumeric(Char) Then
        Call cmdNum_Click(Val(Char))
        Exit Sub
    End If
    
    If Char = "." Then
        Call cmdPt_Click        ' Simulate Point
    ElseIf Char = "+" Then
        Call cmdAdd_Click       ' Simulate Add
    ElseIf Char = "-" Then
        Call cmdSubtract_Click  ' Simulate Subtract
    ElseIf Char = "*" Then
        Call cmdMult_Click      ' Simulate Multiply
    ElseIf Char = "/" Or Char = "\" Then
        Call cmdDiv_Click       ' Simulate Divide
    ElseIf Char = "r" Then
        Call cmdSqrt_Click      ' Simulate Square root
    ElseIf Char = "p" Then
        Call cmdPct_Click       ' Simulate %
    ElseIf Char = "i" Then
        Call cmdInvert_Click    ' Simulate  +/-
    ElseIf Char = "c" Then
        Call cmdClear_Click     ' Simulate Clear
    ElseIf Char = Chr(8) Then
        Call cmdBack_Click      ' Simulate Backspace
    ElseIf Char = Chr(13) Or Char = "=" Then
        Call cmdSolve_Click     ' Simulate SOLVE
    
    
    Else ' Ignore all other keypresses
    
    End If
        
End Sub

Private Sub ColorButtons()
Dim i As Integer

' Disable this function to allow for pausing in IDE

    ' Blue numbers
    For i = 0 To 9
        SetButtonForeColor cmdNum(i), vbBlue
    Next
    
    ' Blue buttons
    SetButtonForeColor cmdInvert, vbBlue
    SetButtonForeColor cmdPt, vbBlue
    SetButtonForeColor cmdSqrt, vbBlue
    SetButtonForeColor cmdPct, vbBlue
    SetButtonForeColor cmdInv, vbBlue
    
    ' Red buttons
    SetButtonForeColor cmdMemClear, vbRed
    SetButtonForeColor cmdMR, vbRed
    SetButtonForeColor cmdMS, vbRed
    SetButtonForeColor cmdMPlus, vbRed
    SetButtonForeColor cmdBack, vbRed
    SetButtonForeColor cmdCE, vbRed
    SetButtonForeColor cmdClear, vbRed
    SetButtonForeColor cmdDiv, vbRed
    SetButtonForeColor cmdMult, vbRed
    SetButtonForeColor cmdSubtract, vbRed
    SetButtonForeColor cmdAdd, vbRed
    SetButtonForeColor cmdSolve, vbRed
    
End Sub

Private Sub UNColorButtons()
Dim cmd As Control
    For Each cmd In Me
        If TypeOf cmd Is CommandButton Then
            UnsetButtonForeColor cmd
        End If
    Next
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Calculator Plus:  Copyright - Joe Jordan 2002.", vbInformation, "About.."
End Sub

Private Sub mnuCopy_Click()
    ' Clear clipboard and set equation to it
    Clipboard.Clear
    CopyContent = txtCalc.Text
    Clipboard.SetText txtCalc.Text
End Sub

Private Sub mnuPaste_Click()
Dim ClipText As String
    ClipText = Clipboard.GetText
    If ClipText = CopyContent Then
        If Trim(CopyContent) <> "" Then
            ' Equation copied to clipboard is still there
            txtCalc.Text = CopyContent
        End If
    End If
End Sub

Private Sub mnuHelp_Click()
    MsgBox "Information:  This calculator performs multiple operations on one line.  It will follow traditional ""order of operations"" -- multiply and divide are processed before add and subtract.  See features file for more feature info.", vbInformation, "Helpful Information"
End Sub


