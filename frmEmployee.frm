VERSION 5.00
Begin VB.Form frmEmployee 
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Employee"
      Height          =   495
      Left            =   4935
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox lstEmp 
      Height          =   3180
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   7215
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List Employees"
      Height          =   495
      Left            =   2535
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtEmpSalary 
      Height          =   375
      Left            =   5535
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtEmpName 
      Height          =   375
      Left            =   1935
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtEmpNo 
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New Employee"
      Height          =   495
      Left            =   135
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblEmp 
      AutoSize        =   -1  'True
      Caption         =   "Salary (Rs.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label lblEmp 
      AutoSize        =   -1  'True
      Caption         =   "Emp Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblEmp 
      AutoSize        =   -1  'True
      Caption         =   "Emp No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEmployees As Employees

Private Sub cmdAdd_Click()
  
  If Not IsNumeric(Trim$(txtEmpNo.Text)) Then Exit Sub
  If Trim$(txtEmpName.Text) = "" Then Exit Sub
  If Not IsNumeric(Trim$(txtEmpSalary.Text)) Then Exit Sub

  objEmployees.Add CLng(Trim$(txtEmpNo.Text)), Trim$(txtEmpName.Text), CDbl(Trim$(txtEmpSalary.Text))
  
  txtEmpNo.Text = "": txtEmpName.Text = "": txtEmpSalary.Text = ""
  
  txtEmpNo.SetFocus
  
End Sub

Private Sub cmdList_Click()
Dim objEmp As clsEmployee
Dim i As Integer

  lstEmp.Clear
  
  For Each objEmp In objEmployees
    lstEmp.AddItem objEmp.EmployeeNo & " " & objEmp.EmployeeName & " " & objEmp.EmployeeSalary
  Next objEmp
  
'  For i = 1 To objEmployees.Count
'    lstEmp.AddItem objEmployees(i).EmployeeNo & " " & objEmployees(i).EmployeeName & " " & objEmployees(i).EmployeeSalary
'  Next i
  
End Sub

Private Sub cmdRemove_Click()

  If lstEmp.ListIndex = -1 Then Exit Sub

  objEmployees.Remove lstEmp.ListIndex + 1
  
  lstEmp.RemoveItem lstEmp.ListIndex
  
End Sub

Private Sub Form_Load()
  
  Set objEmployees = New Employees
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set objEmployees = Nothing
  
End Sub

Private Sub txtEmpNo_Change()
  
  txtEmpNo.Text = Trim$(Me.txtEmpNo.Text)
  
  Me.txtEmpNo.SelStart = Len(Me.txtEmpNo.Text)
  
End Sub

Private Sub txtEmpSalary_Change()

  Me.txtEmpSalary.Text = Trim$(Me.txtEmpSalary.Text)
  
  Me.txtEmpSalary.SelStart = Len(Me.txtEmpSalary.Text)
  
End Sub
