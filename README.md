# VBA-KeyedCollection

An enhanced version of a classic VBA.Collection, Keyed Collection is a combination of functionality found in a regular VBA.Collection and a Scripting.Dictionary plus some extra utilities. It uses 2 syncronized internal collections in order to achieve the functionality (one for items and one for keys). 

Forces the use of keys for all items and also allows the retrieval of keys like a Dictionary does. Although keys can only be of ```String``` type and are case insensitive (like in a regular VBA.Collection), this custom collection is suitable to plenty of tasks where a simple collection does not provide some of the much-needed functionality: editing keys and items, retrieving keys and checking if a specific key exists. For example, parsing JSON would be a perfect fit as keys are always of ```String``` type in the JSON data-interchange format.

The items/keys are ordered allowing the use of Before/After parameters just like in a VBA.Collection. Removes the need for extra libraries (i.e. Microsoft Scripting Runtime) in most cases and obviously works on Mac as well. Items and keys can both be seen in the Locals Window of the VBE, unlike a Dictionary.

## Installation
Just import the following code modules in your VBA Project:
* **EnumHelper.cls**
* **KeyedCollection.cls**

## Demo

```VBA
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
```

## Notes
* The ```EnumHelper``` class helps with avoiding a [x64 bug](https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes) but also allows the iteration of both keys and items using a ```For Each...``` loop on the ```.KeysEnum``` and ```.ItemsEnum``` methods.

## License
MIT License

Copyright (c) 2017 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.