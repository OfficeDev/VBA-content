
# AddIns2.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that returns a **AddIns2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the  **AddIns2** method is the **Title** property of the **AddIn** object.


```vb
If ThisWorkbook.Application.AddIns2.Item("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```


## See also


#### Concepts


[AddIns2 Object](ca4bff78-8ddb-6bc3-b95a-a06a9f75dd88.md)
#### Other resources


[AddIns2 Object Members](6f9dfc17-648d-a004-2321-d3ed86cd438f.md)
