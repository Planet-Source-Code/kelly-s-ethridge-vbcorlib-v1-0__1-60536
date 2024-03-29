VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArgumentNullException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE As String = "User Error."

Private ex As ArgumentNullException
Private Inner As Exception

Public Sub Setup()
    Set Inner = New Exception
    Set ex = Cor.NewArgumentNullException(EXPECTED_MESSAGE, , Inner)
End Sub


Public Sub TestUserDefinedMessage()
    Assert.Equals EXPECTED_MESSAGE, ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Assert.Equals Inner, ex.InnerException, "InnerException should return UserException instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As ArgumentNullException
    Set firstLevelEx = Cor.NewArgumentNullException(EXPECTED_MESSAGE, , ex)
    Assert.Equals Inner, firstLevelEx.GetBaseException, "GetBaseException should return base instance."
End Sub

Public Sub TestUserDefinedToString()
    Assert.Equals "VBCorLib.ArgumentNullException: " & EXPECTED_MESSAGE, ex.ToString, "ToString return wrong value."
End Sub

Public Sub TestPersitence()
    Dim Copy As ArgumentNullException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", ex
    Set Copy = Bag.ReadProperty("ex")
    
    Assert.Equals EXPECTED_MESSAGE, Copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing Copy.InnerException, "InnerException should not be Nothing."
    'Assert.Equals "Argument", Copy.ParamName, "ParamName should return 'Argument'."
End Sub

'Public Sub TestParamName()
'    Assert.Equals "Argument", ex.ParamName, "ParamName should return 'Argument'."
'End Sub

