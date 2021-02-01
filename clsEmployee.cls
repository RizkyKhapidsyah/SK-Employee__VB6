VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_EmpName As String
Private m_EmpNo As Long
Private m_EmpSalray As Double

Public Property Get EmployeeName() As String

  EmployeeName = m_EmpName
  
End Property

Public Property Let EmployeeName(ByVal New_EmpName As String)

  m_EmpName = New_EmpName
  
End Property

Public Property Get EmployeeNo() As Long

  EmployeeNo = m_EmpNo
  
End Property

Public Property Let EmployeeNo(ByVal New_EmpNo As Long)
  
  m_EmpNo = New_EmpNo
  
End Property

Public Property Get EmployeeSalary() As Double

  EmployeeSalary = m_EmpSalray
  
End Property

Public Property Let EmployeeSalary(ByVal New_EmpSalary As Double)

  m_EmpSalray = New_EmpSalary
  
End Property
