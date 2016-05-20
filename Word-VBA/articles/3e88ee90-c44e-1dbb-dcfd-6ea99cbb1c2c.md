
# MailMergeDataSource.FieldNames Property (Word)

Returns a  **[MailMergeFieldNames](5a3752da-63b2-f0f9-7456-01a31bac5f62.md)** collection that represents the names of all the fields in the specified mail merge data source. Read-only.


## Syntax

 _expression_ . **FieldNames**

 _expression_ A variable that represents a **[MailMergeDataSource](f86f7d3c-d7ab-45e8-21e7-fd5a426e0391.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of the first field in the data source attached to the active mail merge main document.


```vb
MsgBox ActiveDocument.MailMerge.DataSource.FieldNames(1).Name
```

This example uses the mNames() array to store the names of each merge field contained in the data source attached to the active document.




```vb
Dim mNames As Variant 
Dim mmTemp As MailMerge 
Dim intCount As Integer 
Dim intIncrement As Integer 
Dim mmfnLoop As MailMergeFieldName 
 
Set mmTemp = ActiveDocument.MailMerge 
intCount = _ 
 ActiveDocument.MailMerge.DataSource.FieldNames.Count - 1 
 
ReDim mNames(intCount) 
intIncrement = 0 
 
For Each mmfnLoop In mmTemp.DataSource.FieldNames 
 mNames(intIncrement) = mmfnLoop.Name 
 intIncrement = intIncrement + 1 
Next mmfnLoop
```


## See also


#### Concepts


[MailMergeDataSource Object](f86f7d3c-d7ab-45e8-21e7-fd5a426e0391.md)
#### Other resources


[MailMergeDataSource Object Members](a52f088c-2507-8f39-17b9-9b97c8a8ed7e.md)
