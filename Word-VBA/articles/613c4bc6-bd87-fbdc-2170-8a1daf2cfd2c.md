
# MailMergeDataSource.DataFields Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [MailMergeDataFields](a660288d-1a2c-53ec-20d2-c52353be90c8.md)**collection that represents the fields in the specified mail merge data source. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DataFields**

 _expression_A variable that represents a  ** [MailMergeDataSource](f86f7d3c-d7ab-45e8-21e7-fd5a426e0391.md)** object.


## Remarks
<a name="sectionSection1"> </a>

For information about returning a single member of a collection, see  [Returning an Object from a Collection](28f76384-f495-9640-a7c8-10ada3fac727.md).


## Example
<a name="sectionSection2"> </a>

This example displays the name of each field in the data source attached to the active mail merge main document.


```
Dim mmdfTemp As MailMergeDataField 
 
For Each mmdfTemp In _ 
 ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox mmdfTemp.Name 
Next mmdfTemp
```

This example displays the value of the LastName field from the first record in the data source attached to "Main.doc."




```
With Documents("Main.doc").MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 MsgBox .DataFields("LastName").Value 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [MailMergeDataSource Object](f86f7d3c-d7ab-45e8-21e7-fd5a426e0391.md)
#### Other resources


 [MailMergeDataSource Object Members](a52f088c-2507-8f39-17b9-9b97c8a8ed7e.md)
