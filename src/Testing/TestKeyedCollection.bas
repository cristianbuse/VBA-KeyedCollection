Attribute VB_Name = "TestKeyedCollection"
Option Explicit
Option Private Module

'Rubberduck COM add-in is needed to run the tests in this module:
'https://rubberduckvba.com/

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private Type EXPECTED_ERROR
    code_ As Long
    wasRaised As Boolean
End Type

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

Private Function NewExpectedError(code_) As EXPECTED_ERROR
    NewExpectedError.code_ = code_
    NewExpectedError.wasRaised = False
End Function

'@TestMethod("Add")
Private Sub TestAddValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long

    'Act:
    'Add at the end
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Add before
    c.Add "key2.5", 2.5, 3
    c.Add "key2.6", 2.6, "key3"
    c.Add "key0.5", 0.5, 1
    c.Add "key0.6", 0.6, "key1"
    c.Add "key0.4", 0.4, "key0.5"
    
    'Add after
    c.Add "key5.5", 5.5, , "key5"
    c.Add "key5.6", 5.6, , 11
    c.Add "key0.7", 0.7, , 3
    
    'Assert:
    Assert.Succeed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Add")
Private Sub TestAddInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    'Act:
    'Wrong index_ for before_ (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    c.Add "key", 1, 1
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Wrong index_ for after_ (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    c.Add "key", 1, , 1
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Arrange:
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Both before_ and after_ specified
    ExpectedError = NewExpectedError(5)
    c.Add "key6", 6, 2, 1
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Wrong key_ for before_
    ExpectedError = NewExpectedError(5)
    c.Add "key6", 6, "key9"
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Wrong key_ for after_
    ExpectedError = NewExpectedError(5)
    c.Add "key6", 6, , "key9"
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Wrong index_ for before_ (while collection has elements)
    ExpectedError = NewExpectedError(9)
    c.Add "key6", 6, 9
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Wrong index_ for after_ (while collection has elements)
    ExpectedError = NewExpectedError(9)
    c.Add "key6", 6, , -5
    If Not ExpectedError.wasRaised Then GoTo AssertFail

    'Wrong data type for before_
    ExpectedError = NewExpectedError(13)
    c.Add "key6", 6, Nothing
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Wrong data type for after_
    ExpectedError = NewExpectedError(13)
    c.Add "key6", 6, , Array()
    If Not ExpectedError.wasRaised Then GoTo AssertFail

    'Duplicated key_
    ExpectedError = NewExpectedError(457)
    c.Add "key5", 5
    If Not ExpectedError.wasRaised Then GoTo AssertFail
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("CompareMode")
Private Sub TestCompareMode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection

    'Assert:
    Assert.IsTrue c.CompareMode = VbCompareMethod.vbTextCompare
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Count")
Private Sub TestCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    'Assert:
    Assert.IsTrue c.Count = 0

    'Arrange:
    c.Add 1, 1
    c.Add 2, 2
    'Assert:
    Assert.IsTrue c.Count = 2
    
    'Arrange:
    c.RemoveAll
    'Assert:
    Assert.IsTrue c.Count = 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Exists")
Private Sub TestExists()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    c.Add "key1", 1
    c.Add "key3", 1

    'Assert:
    Assert.IsTrue c.Exists("key1")
    Assert.IsTrue c.Exists("key3")
    Assert.IsFalse c.Exists("key2")
    Assert.IsFalse c.Exists("test")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Item")
Private Sub TestGetItemValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim u As IUnknown
    Dim i As Long

    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    c.Add "coll", New Collection
    Set u = New Collection
    c.Add "unk", u
    
    'Assert:
    Assert.IsTrue c.Item("key2") = 2
    Assert.IsTrue c.Item("key4") = 4
    Assert.IsFalse c.Item("key5") = 0
    Assert.IsTrue c.Item("coll").Count = 0
    Assert.IsTrue c.Item(1) = 1
    Assert.IsTrue c.Item(3) = 3
    Assert.IsFalse c.Item(2) = 3
    Assert.IsTrue TypeName(c.Item(c.Count)) = "Collection"
    Assert.IsTrue c(1) = 1 'Default member - equivalent with c.Item(1)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Item")
Private Sub TestGetItemInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim v As Variant
    Dim i As Long
    
    'Act:
    'Invalid Index (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    v = c.Item(9)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Arrange:
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Invalid Key
    ExpectedError = NewExpectedError(5)
    v = c.Item("keyNone")
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid Index (while collection has elements)
    ExpectedError = NewExpectedError(9)
    v = c.Item(9)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid data type for index/key
    ExpectedError = NewExpectedError(13)
    v = c.Item(Nothing)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    v = c.Item(Null)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    v = c.Item(Array(1))
    If Not ExpectedError.wasRaised Then GoTo AssertFail
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("Item")
Private Sub TestLetItemValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long

    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    c.Item("key1") = "val1"
    'Assert:
    Assert.IsTrue c.Item("key1") = "val1"
    
    'Act:
    c.Item(2) = "val2"
    'Assert:
    Assert.IsTrue c.Item(2) = "val2"
    
    'Act:
    c(5) = "val5"
    'Assert:
    Assert.IsTrue c.Item("key5") = "val5"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Item")
Private Sub TestLetItemInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    'Act:
    'Invalid Index (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    c.Item(6) = 15
    If Not ExpectedError.wasRaised Then GoTo AssertFail

    'Arrange:
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Invalid Key
    ExpectedError = NewExpectedError(5)
    c.Item("keyNone") = "none"
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid Index (while collection has elements)
    ExpectedError = NewExpectedError(9)
    c.Item(9) = 15
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid data type for index/key
    ExpectedError = NewExpectedError(13)
    c.Item(Nothing) = 3
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    c.Item(Null) = 5
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    c.Item(Array(1)) = 7
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Missing 'Set'
    ExpectedError = NewExpectedError(450)
    c.Item(1) = New Collection
    If Not ExpectedError.wasRaised Then GoTo AssertFail
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("Item")
Private Sub TestSetItemValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    Dim coll As New Collection

    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    Set c.Item("key1") = coll
    'Assert:
    Assert.IsTrue c.Item("key1") Is coll
    
    'Act:
    Set c.Item(5) = Application
    'Assert:
    Assert.IsTrue c.Item(5) Is Application
    
    'Act:
    Set c.Item(5) = Nothing
    'Assert:
    Assert.IsTrue c.Item(5) Is Nothing
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Item")
Private Sub TestSetItemInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    Dim coll As New Collection
    
    'Act:
    'Invalid Index (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    Set c.Item(6) = Nothing
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Arrange:
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Invalid Key
    ExpectedError = NewExpectedError(5)
    Set c.Item("keyNone") = coll
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid Index (while collection has elements)
    ExpectedError = NewExpectedError(9)
    Set c.Item(9) = Nothing
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid data type for index/key
    ExpectedError = NewExpectedError(13)
    Set c.Item(Nothing) = Nothing
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    Set c.Item(Null) = Application
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    Set c.Item(Array(1)) = coll
    If Not ExpectedError.wasRaised Then GoTo AssertFail
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("Items")
Private Sub TestItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim arr() As Variant
    Dim i As Long
    
    arr = c.Items
    
    'Assert:
    Assert.IsTrue UBound(arr) - LBound(arr) + 1 = 0
    
    'Arrange:
    For i = 1 To 5
        c.Add CStr(i), i
    Next i
    arr = c.Items
    
    'Assert:
    Assert.IsTrue UBound(arr) - LBound(arr) + 1 = c.Count
    For i = LBound(arr) To UBound(arr)
        Assert.IsTrue arr(i) = i
    Next i
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("NewEnum")
Private Sub TestItemsEnum()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim v As Variant
    Dim i As Long
    
    'Assert:
    For Each v In c.ItemsEnum
        Assert.Fail
    Next v
    
    'Arrange:
    For i = 1 To 5
        c.Add CStr(i), i
    Next i
    
    'Assert:
    i = 1
    For Each v In c.ItemsEnum
        Assert.IsTrue v = i
        i = i + 1
    Next v
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("GetKey")
Private Sub TestGetKeyAtIndexValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Assert:
    Assert.IsTrue c.GetKeyAtIndex(2) = "key2"
    Assert.IsTrue c.GetKeyAtIndex(4) = "key4"
    Assert.IsFalse c.GetKeyAtIndex(5) = "key4"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("GetKey")
Private Sub TestGetKeyAtIndexInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim s As String
    Dim i As Long
    
    'Act:
    'Invalid index (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    s = c.GetKeyAtIndex(1)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Arrange:
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Invalid index (while collection has elements)
    ExpectedError = NewExpectedError(9)
    s = c.GetKeyAtIndex(6)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    '
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("Key")
Private Sub TestKeyValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    c.Key("key1") = "k1"
    
    'Assert:
    Assert.IsTrue c.GetKeyAtIndex(1) = "k1"
    
    'Act:
    c.Key("key5") = "key1"
    
    'Assert:
    Assert.IsTrue c.GetKeyAtIndex(5) = "key1"
    
    'Act:
    c.Key("key3") = "key3"
    
    'Assert:
    Assert.IsTrue c.GetKeyAtIndex(3) = "key3"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Key")
Private Sub TestKeyInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Invalid old key
    ExpectedError = NewExpectedError(5)
    c.Key("key9") = "newKey"
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid new key
    ExpectedError = NewExpectedError(457)
    c.Key("key1") = "key2"
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    '
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("Keys")
Private Sub TestKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim arr() As Variant
    Dim i As Long
    
    arr = c.Keys
    
    'Assert:
    Assert.IsTrue UBound(arr) - LBound(arr) + 1 = 0
    
    'Arrange:
    For i = 1 To 5
        c.Add CStr(i), i
    Next i
    arr = c.Keys
    
    'Assert:
    Assert.IsTrue UBound(arr) - LBound(arr) + 1 = c.Count
    For i = LBound(arr) To UBound(arr)
        Assert.IsTrue arr(i) = CStr(i)
    Next i
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("NewEnum")
Private Sub TestKeysEnum()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim v As Variant
    Dim i As Long
    
    'Assert:
    For Each v In c.KeysEnum
        Assert.Fail
    Next v
    
    'Arrange:
    For i = 1 To 5
        c.Add CStr(i), i
    Next i
    
    'Assert:
    i = 1
    For Each v In c.KeysEnum
        Assert.IsTrue v = CStr(i)
        i = i + 1
    Next v
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("KeyItemPairs")
Private Sub KeyItemPairs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim arr() As Variant
    Dim i As Long
    
    arr = c.KeyItemPairs
    
    'Assert:
    Assert.IsTrue UBound(arr, 1) - LBound(arr, 1) + 1 = 0
    
    'Arrange:
    For i = 1 To 5
        c.Add CStr(i), i
    Next i
    arr = c.KeyItemPairs
    
    'Assert:
    Assert.IsTrue UBound(arr, 1) - LBound(arr, 1) + 1 = c.Count
    Assert.IsTrue UBound(arr, 2) - LBound(arr, 2) + 1 = 2
    For i = LBound(arr, 1) To UBound(arr, 1)
        Assert.IsTrue arr(i, 1) = CStr(i)
        Assert.IsTrue arr(i, 2) = i
    Next i
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Remove")
Private Sub TestRemoveValidArgs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act
    c.Remove 1
    
    'Assert:
    Assert.IsTrue c.Count = 4
    Assert.IsTrue c.Item(1) = 2
    
    'Act
    c.Remove "key3"
    
    'Assert:
    Assert.IsTrue c.Count = 3
    Assert.IsTrue c.Item(2) = 4
    
    'Act
    c.Remove 1
    c.Remove 1
    c.Remove 1
    
    'Assert:
    Assert.IsTrue c.Count = 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Remove")
Private Sub TestRemoveInvalidArgs()
    Dim ExpectedError As EXPECTED_ERROR
    
    On Error GoTo TestFail

    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    'Act:
    'Invalid index (while collection has no elements)
    ExpectedError = NewExpectedError(5)
    c.Remove 1
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Arrange:
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act:
    'Invalid Key
    ExpectedError = NewExpectedError(5)
    c.Remove "keyNone"
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid index (while collection has elements)
    ExpectedError = NewExpectedError(9)
    c.Remove 10
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    'Invalid data type for index/key
    ExpectedError = NewExpectedError(13)
    c.Remove Nothing
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    c.Remove Null
    If Not ExpectedError.wasRaised Then GoTo AssertFail
    
    ExpectedError = NewExpectedError(13)
    c.Remove Array(1)
    If Not ExpectedError.wasRaised Then GoTo AssertFail
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError.code_ Then
        ExpectedError.wasRaised = True
        Resume Next
    End If
AssertFail:
    Assert.Fail "Expected error was not raised"
End Sub

'@TestMethod("RemoveAll")
Private Sub TestRemoveAll()
    On Error GoTo TestFail
    
    'Arrange:
    Dim c As New KeyedCollection
    Dim i As Long
    
    For i = 1 To 5
        c.Add "key" & i, i
    Next i
    
    'Act
    c.RemoveAll
    
    'Assert:
    Assert.IsTrue c.Count = 0

    'Arrange:
    c.Add "1", 1
    
    'Act
    c.RemoveAll
    
    'Assert:
    Assert.IsTrue c.Count = 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
