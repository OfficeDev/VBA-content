
# MailMergeDataFields Object (Publisher)

A collection of  **[MailMergeDataField](46768b72-482c-06c5-5e77-27a95109f610.md)** objects that represent the data fields in a mail merge or catalog merge data source.
 


## Remarks

You cannot add fields to the  **MailMergeDataFields** collection. When a data field is added to a data source, the field is automatically included in the **MailMergeDataFields** collection.
 

 

## Example

Use the  **[DataFields](820af882-d54c-a205-2925-e7110fc0c02b.md)** property to return the **MailMergeDataFields** collection.
 

 

 

 
The following example displays the field names in the data source attached to the active publication.
 

 



```
Sub ShowFieldNames() 
 Dim intCount As Integer 
 With ActiveDocument.MailMerge.DataSource.DataFields 
 For intCount = 1 To .Count 
 MsgBox .Item(intCount).Name 
 Next 
 End With 
End Sub
```

Use  **DataFields** (index), where index is the data field name or the index number, to return a single **MailMergeDataField** object. The index number represents the position of the data field in the mail merge data source. This example retrieves the name of the first field and value of the first record of the FirstName field in the data source attached to the active publication.
 

 



```
Sub GetDataFromSource() 
 With ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox "First field name: " &amp; .Item(1).Name &amp; vbLf &amp; _ 
 "Value of the first record of the FirstName field: " &amp; _ 
 .Item("FirstName").Value 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](d0305807-90a2-87eb-60cb-cd45925c5ed0.md)|

## Properties



|**Name**|
|:-----|
|[Application](8ee03f9e-1996-58e7-c6e9-3a1ccd0d2963.md)|
|[Count](f46da7b1-acd8-f2d2-a6aa-71cc3c8eca99.md)|
|[Creator](3d3adcd4-6746-7c0b-58be-bba04f9ba1e2.md)|
|[Parent](4ee811c9-1eca-6ed7-a60a-6541acac32f3.md)|
