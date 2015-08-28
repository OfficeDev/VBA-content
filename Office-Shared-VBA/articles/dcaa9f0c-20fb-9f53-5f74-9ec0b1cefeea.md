
# COMAddIn Object (Office)

 **Last modified:** July 28, 2015

Represents a COM add-in in the Microsoft Office host application. The  **COMAddIn** object is a member of the **COMAddIns**collection.

## Example

Use  **COMAddIns.Item(index)**, where  _index_ is either an ordinal value that returns the COM add-in at that position in the **COMAddIns** collection, or a **String** value that represents the ProgID of the specified COM add-in. The following example displays a COM add-in's description text in a message box.


```
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```

Use the  **ProgID**property of the  **COMAddin** object to return the programmatic identifier for a COM add-in, and use the **Guid**property to return the globally unique identifier (GUID) for the add-in. The following example displays the ProgID and GUID for COM add-in one in a message box.




```
MsgBox "My ProgID is " &amp; _ 
 Application.COMAddIns(1).ProgID &amp; _ 
 " and my GUID is " &amp; _ 
 Application.COMAddIns(1).Guid
```

Use the  **Connect**property to set or return the state of the connection to a specified COM add-in. The following example displays a message box that indicates whether COM add-in one is registered and currently connected.




```
If Application.COMAddIns(1).Connect Then 
 MsgBox "The add-in is connected." 
Else 
MsgBox "The add-in is not connected." 
End If
```


## See also


#### Concepts


 [Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Other resources


 [COMAddIn Object Members](698d4d8e-6071-acd3-a39b-ab01fd878452.md)
