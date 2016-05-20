
# ListObject.ListRows Property (Excel)

Returns a  **[ListRows](e4035209-00a2-ea16-a3b9-2d23afe0b88a.md)** object that represents all the rows of data in the **[ListObject](46de6c4f-8ce0-0c7d-da59-6e52f5eab612.md)** object. Read-only.


## Syntax

 _expression_ . **ListRows**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

The  **ListRows** object returned does not include the header, total, or Insert rows.


## Example

The following example deletes a row specified by number in the  **ListRows** collection that is created by a call to the **ListRows** property.


```vb
Sub DeleteListRow(iRowNumber As Integer) 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListRows As ListRows 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListRows = objListObj.ListRows 
 
 If (iRowNumber <> 0) And (iRowNumber < objListRows.Count - 1) Then 
 objListRows(iRowNumber).Delete 
 End If 
End Sub
```


## See also


#### Concepts


[ListObject Object](46de6c4f-8ce0-0c7d-da59-6e52f5eab612.md)
#### Other resources


[ListObject Object Members](d34f895c-cf60-f644-866b-7b757716e7a6.md)
