---
title: MailMergeDataFields Object (Publisher)
keywords: vbapb10.chm6422527
f1_keywords:
- vbapb10.chm6422527
ms.prod: publisher
api_name:
- Publisher.MailMergeDataFields
ms.assetid: 44ae8a3c-b8a8-fc57-9d02-d71dcffc21ef
ms.date: 06/08/2017
---


# MailMergeDataFields Object (Publisher)

A collection of  **[MailMergeDataField](mailmergedatafield-object-publisher.md)** objects that represent the data fields in a mail merge or catalog merge data source.
 


## Remarks

You cannot add fields to the  **MailMergeDataFields** collection. When a data field is added to a data source, the field is automatically included in the **MailMergeDataFields** collection.
 

 

## Example

Use the  **[DataFields](mailmergedatasource-datafields-property-publisher.md)** property to return the **MailMergeDataFields** collection.
 

 

 

 
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
|[Item](mailmergedatafields-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](mailmergedatafields-application-property-publisher.md)|
|[Count](mailmergedatafields-count-property-publisher.md)|
|[Creator](mailmergedatafields-creator-property-publisher.md)|
|[Parent](mailmergedatafields-parent-property-publisher.md)|

