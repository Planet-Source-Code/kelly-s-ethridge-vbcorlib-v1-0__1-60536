VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArgumentOutOfRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE  As String = "User Error."


Private ex As ArgumentOutOfRangeException
Private Inner As Exception

Public Sub Setup()
    Set Inner = New Exception
    Set ex = Cor.NewArgumentOutOfRangeException(EXPECTED_MESSAGE, "Argument", "Value", Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.IsTrue Len(ex.Message) > 0, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Assert.Equals Inner, ex.InnerException, "InnerException should return Inner instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As ArgumentOutOfRangeException
    Set firstLevelEx = Cor.NewArgumentOutOfRangeException(EXPECTED_MESSAGE, , , ex)
    Assert.Equals Inner, firstLevelEx.GetBaseException, "GetBaseException should return base instance."
End Sub

Public Sub TestUserDefinedToString()
    Assert.IsTrue Len(ex.ToString) > 0, "ToString return wrong value."
End Sub

Public Sub TestPersitence()
    Dim Copy As ArgumentOutOfRangeException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.IsTrue Len(Copy.Message) > 0, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."
    Assert.Equals "Argument", Copy.ParamName, "ParamName should return 'Argument'."
    Assert.Equals "Value", Copy.ActualValue, "ActualValue should return 'Value'."
End Sub

Public Sub TestParamName()
    Assert.Equals "Argument", ex.ParamName, "ParamName should return 'Argument'."
End Sub

Public Sub TestActualValue()
    Assert.Equals "Value", ex.ActualValue, "ActualValue should return 'Value'."
End Sub
