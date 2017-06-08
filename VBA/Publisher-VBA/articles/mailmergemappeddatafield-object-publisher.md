---
title: MailMergeMappedDataField Object (Publisher)
keywords: vbapb10.chm6619135
f1_keywords:
- vbapb10.chm6619135
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataField
ms.assetid: 3711d28e-f005-27fb-88b5-8674d4ece887
ms.date: 06/08/2017
---


# MailMergeMappedDataField Object (Publisher)

Represents a single mapped data field. The  **MailMergeMappedDataField** object is a member of the **[MailMergeMappedDataFields](mailmergemappeddatafields-object-publisher.md)** collection. A mapped data field is a field contained within Microsoft Publisher that represents commonly used name or address information, such as First Name. If a data source contains a First Name field or a variation (such as First_Name, FirstName, First, or FName), the field in the data source will automatically map to the corresponding mapped data field. If a publication is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the publication to agree with the field names in the database.
 


## Example

Use  **MappedDataFields** (index) to return a **MailMergeMappedDataField** object. This example returns the data source field name for the **pbFirstName** mapped data field. This example assumes the current publication is a mail merge publication. A blank string value returned for the **DataFieldName** property indicates that the mapped data field is not mapped to a field in the data source.
 

 

```
Sub MappedFieldName() 
 Dim strMappedDataField As String 
 With ActiveDocument.MailMerge.DataSource 
 strMappedDataField = .MappedDataFields(pbFirstName).DataFieldName 
 If strMappedDataField <> "" Then 
 MsgBox "The mapped data field 'FirstName' is mapped to " _ 
 &amp; .MappedDataFields(pbFirstName).DataFieldName &amp; "." 
 Else 
 MsgBox "The mapped data field 'FirstName' is not " &amp; _ 
 "mapped to any of the data fields in your " &amp; _ 
 "data source." 
 End If 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](mailmergemappeddatafield-application-property-publisher.md)|
|[DataFieldName](mailmergemappeddatafield-datafieldname-property-publisher.md)|
|[Index](mailmergemappeddatafield-index-property-publisher.md)|
|[Name](mailmergemappeddatafield-name-property-publisher.md)|
|[Parent](mailmergemappeddatafield-parent-property-publisher.md)|
|[Value](mailmergemappeddatafield-value-property-publisher.md)|

