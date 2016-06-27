
# MailMergeDataSource.DataFields Property (Word)

Returns a  **[MailMergeDataFields](a660288d-1a2c-53ec-20d2-c52353be90c8.md)** collection that represents the fields in the specified mail merge data source. Read-only.


## Syntax

 _expression_ . **DataFields**

 _expression_ A variable that represents a **[MailMergeDataSource](f86f7d3c-d7ab-45e8-21e7-fd5a426e0391.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of each field in the data source attached to the active mail merge main document.


```vb
Dim mmdfTemp As MailMergeDataField 
 
For Each mmdfTemp In _ 
 ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox mmdfTemp.Name 
Next mmdfTemp
```

This example displays the value of the LastName field from the first record in the data source attached to "Main.doc."




```vb
With Documents("Main.doc").MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 MsgBox .DataFields("LastName").Value 
End With
```


## See also


#### Concepts


[MailMergeDataSource Object](f86f7d3c-d7ab-45e8-21e7-fd5a426e0391.md)
#### Other resources


[MailMergeDataSource Object Members](a52f088c-2507-8f39-17b9-9b97c8a8ed7e.md)
