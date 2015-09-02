
# Application.WorkbookPivotTableCloseConnection Event (Excel)

 **Last modified:** July 28, 2015

Occurs after a PivotTable report connection has been closed.

## Syntax

 _expression_. **WorkbookPivotTableCloseConnection**( **_Wb_**,  **_Target_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Wb|Required| ** [Workbook](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)**|The selected workbook.|
|Target|Required| ** [PivotTable](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)**|The selected PivotTable report.|

### Return Value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been closed. This example assumes you have declared an object of type  **Workbook** with events in a class module.


```
Private Sub ConnectionApp_WorkbookPivotTableCloseConnection(ByVal wbOne As Workbook, Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```


## See also


#### Concepts


 [Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
 [Application Object Members](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)
