---
title: MappedDataField Object (Word)
keywords: vbawd10.chm1641
f1_keywords:
- vbawd10.chm1641
ms.prod: word
api_name:
- Word.MappedDataField
ms.assetid: 35b9b770-bf18-8922-7c3a-431f454561e9
ms.date: 06/08/2017
---


# MappedDataField Object (Word)

A mapped data field is a field contained within Microsoft Word that represents commonly used name or address information, such as "First Name." If a data source contains a "First Name" field or a variation (such as "First_Name," "FirstName," "First," or "FName"), the field in the data source will automatically map to the corresponding mapped data field in Word. If a document or template is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the document to agree with the field names in the database.


## Remarks

Use the  **MappedDataFields** property to return a **MappedDataField** object. This example returns the data source field name for the **wdFirstName** mapped data field. This example assumes the current document is a mail merge document. A blank string value returned for the **DataFieldName** property indicates that the mapped data field is not mapped to a field in the data source.


```vb
Sub MappedFieldName() 
 
 With ActiveDocument.MailMerge.DataSource 
 If .MappedDataFields.Item(wdFirstName).DataFieldName <> "" Then 
 MsgBox "The mapped data field 'FirstName' is mapped to " _ 
 &; .MappedDataFields(Index:=wdFirstName) _ 
 .DataFieldName &; "." 
 Else 
 MsgBox "The mapped data field 'FirstName' is not " &; _ 
 "mapped to any of the data fields in your " &; _ 
 "data source." 
 End If 
 
 End With 
 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


