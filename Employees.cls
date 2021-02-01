VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private collEmployees As Collection

Public Function Add(ByVal EmpNo As Long, ByVal EmpName As String, ByVal EmpSalary As Double) As clsEmployee
Dim objEmployee As New clsEmployee

  objEmployee.EmployeeNo = EmpNo
  objEmployee.EmployeeName = EmpName
  objEmployee.EmployeeSalary = EmpSalary
  
  collEmployees.Add objEmployee, CStr(EmpNo)
  
  Set Add = objEmployee
  
End Function

Public Sub Remove(ByVal Index As Variant)

  
  collEmployees.Remove Index
  
End Sub

Public Function Count() As Long
  
  Count = collEmployees.Count
  
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

  Set NewEnum = collEmployees.[_NewEnum]
  
End Function

Public Function Item(ByVal Index As Variant) As clsEmployee
Attribute Item.VB_UserMemId = 0

  Set Item = collEmployees.Item(Index)
  
End Function

Private Sub Class_Initialize()

    Set collEmployees = New Collection
    
End Sub
