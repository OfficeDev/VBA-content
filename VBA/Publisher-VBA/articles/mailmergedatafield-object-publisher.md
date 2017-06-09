---
title: MailMergeDataField Object (Publisher)
keywords: vbapb10.chm6488063
f1_keywords:
- vbapb10.chm6488063
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField
ms.assetid: 46768b72-482c-06c5-5e77-27a95109f610
ms.date: 06/08/2017
---


# MailMergeDataField Object (Publisher)

Represents a single merge field in a data source. The  **MailMergeDataField** object is a member of the **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** collection. The **MailMergeDataFields** collection includes all the data fields in a mail merge or catalog merge data source (for example, Name, Address, and City).
 


## Remarks

You cannot add fields to the  **MailMergeDataFields** collection. All data fields in a data source are automatically included in the **MailMergeDataFields** collection.
 

 

## Example

Use  **[DataFields](mailmergedatasource-datafields-property-publisher.md)** (index), where index is the data field name or index number, to return a single **MailMergeDataField** object. The index number represents the position of the data field in the mail merge data source. This example retrieves the name of the first field and value of the first record of the FirstName field in the data source attached to the active publication.
 

 

```
Sub GetDataFromSource() 
 With ActiveDocument.MailMerge.DataSource 
 MsgBox "Field Name: " &amp; .DataFields.Item(1).Name &amp; _ 
 "Value: " &amp; .DataFields.Item("FirstName").Value 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddToRecipientFields](mailmergedatafield-addtorecipientfields-method-publisher.md)|
|[Insert](mailmergedatafield-insert-method-publisher.md)|
|[MapToRecipientField](mailmergedatafield-maptorecipientfield-method-publisher.md)|
|[UnMapRecipientField](mailmergedatafield-unmaprecipientfield-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](mailmergedatafield-application-property-publisher.md)|
|[Creator](mailmergedatafield-creator-property-publisher.md)|
|[FieldType](mailmergedatafield-fieldtype-property-publisher.md)|
|[Index](mailmergedatafield-index-property-publisher.md)|
|[IsMapped](mailmergedatafield-ismapped-property-publisher.md)|
|[MappedTo](mailmergedatafield-mappedto-property-publisher.md)|
|[Name](mailmergedatafield-name-property-publisher.md)|
|[Parent](mailmergedatafield-parent-property-publisher.md)|
|[Value](mailmergedatafield-value-property-publisher.md)|

