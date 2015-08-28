
# Areas.Count Property (Excel)

 **Last modified:** July 28, 2015

Returns a  **Long** value that represents the number of objects in the collection.

## Syntax

 _expression_. **Count**

 _expression_A variable that represents an  **Areas** object.


## Example

This example displays the number of columns in the selection on Sheet1. The code also tests for a multiple-area selection; if one exists, the code loops on the areas of the multiple-area selection.


```
Sub DisplayColumnCount() 
 Dim iAreaCount As Integer 
 Dim i As Integer 
 
 Worksheets("Sheet1").Activate 
 iAreaCount = Selection.Areas.Count 
 
 If iAreaCount <= 1 Then 
 MsgBox "The selection contains " &amp; Selection.Columns.Count &amp; " columns." 
 Else 
 For i = 1 To iAreaCount 
 MsgBox "Area " &amp; i &amp; " of the selection contains " &amp; _ 
 Selection.Areas(i).Columns.Count &amp; " columns." 
 Next i 
 End If 
End Sub
```


## See also


#### Concepts


 [Areas Collection](43d05ef3-7ae2-2881-dec2-6fec8281f045.md)
#### Other resources


 [Areas Object Members](5df53e64-1fe5-66cb-0777-438a80f399cc.md)
