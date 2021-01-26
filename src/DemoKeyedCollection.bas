Attribute VB_Name = "DemoKeyedCollection"
Option Explicit

Public Sub DemoMain()
    Dim c As New KeyedCollection
    Dim v As Variant
    Dim arr() As Variant
    Dim i As Long
    
    'Add
    c.Add "key1", "item1"
    c.Add "key4", "item4"
    
    'Add After/Before position or key
    c.Add key_:="key2", item_:="item2", after_:=1
    c.Add key_:="key3", item_:="item3", before_:="key4"
    
    'Count
    Debug.Print "Keyed Collection has " & c.Count & " key-value pairs"
    
    'Exists
    Debug.Print "Does collection have ""key3""?: " & c.Exists("Key3") 'Case-insensitive
    Debug.Print "Does collection have ""key5""?: " & c.Exists("Key5") 'Case-insensitive
    
    'GetKeyAtIndex
    Debug.Print "Key at index 1 is: " & c.GetKeyAtIndex(1)
    Debug.Print "Key at index 3 is: " & c.GetKeyAtIndex(3)
    
    'Item <Get>
    Debug.Print "Item at position 2 is: " & c.Item(2)
    Debug.Print "Item with ""key4"" is: " & c.Item("key4")
    
    'Item <Set>
    Set c.Item(3) = Nothing
    Debug.Print "Item at position 3 was changed to: " & TypeName(c.Item(3))
    Set c.Item("key4") = New Collection
    Debug.Print "Item with ""key4"" was changed to: " & TypeName(c.Item("key4"))
    
    'Item <Let>
    c.Item(3) = 3.5
    Debug.Print "Item at position 3 was changed to: " & c.Item(3)
    c.Item("key4") = 3.7
    Debug.Print "Item with ""key4"" was changed to: " & c.Item("key4")

    'Items - returns 1D array
    Debug.Print "Items are: " & VBA.Join(c.Items, ", ")
    
    'ItemsEnum
    Debug.Print "Items can be iterated using a For Each... loop: "
    For Each v In c.ItemsEnum
        Debug.Print v
    Next v
    
    'Key
    c.Key("key2") = "key2.5"
    Debug.Print "Key ""key2"" was changed to: """ & c.GetKeyAtIndex(2) & """"
    
    'KeyItemPairs - returns 2D array
    arr = c.KeyItemPairs
    Debug.Print "Key-item pairs are:"
    For i = LBound(arr, 1) To UBound(arr, 1)
        Debug.Print arr(i, 0) & ", " & arr(i, 1)
    Next i
    
    'Keys - returns 1D array
    Debug.Print "Keys are: " & VBA.Join(c.Keys, ", ")
    
    'KeysEnum
    Debug.Print "keys can be iterated using a For Each... loop: "
    For Each v In c.KeysEnum
        Debug.Print v
    Next v

    'Remove
    c.Remove 1
    c.Remove "key3"
    
    'RemoveAll
    c.RemoveAll
    Debug.Print "All items were removed. Item/Key count is now: " & c.Count
End Sub
