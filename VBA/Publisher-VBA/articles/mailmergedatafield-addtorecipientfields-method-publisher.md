---
title: MailMergeDataField.AddToRecipientFields Method (Publisher)
keywords: vbapb10.chm6422562
f1_keywords:
- vbapb10.chm6422562
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.AddToRecipientFields
ms.assetid: eaf365f0-a9f4-c6e2-1267-d0a31b5934ce
ms.date: 06/08/2017
---


# MailMergeDataField.AddToRecipientFields Method (Publisher)

Adds the parent  **MailMergeDataField** object from a particular data source to the master data source (collection of data fields) for a mail-merge publication.


## Syntax

 _expression_. **AddToRecipientFields**

 _expression_A variable that represents a  **MailMergeDataField** object.


## Remarks

This method works only if the parent  **MailMergeDataField** object has not already been mapped to a recipient field. You can use the **[IsMapped](mailmergedatafield-ismapped-property-publisher.md)** property of the **MailMergeDataField** object to determine if the object has already been mapped.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AddToRecipientFields** method to add a data field (column) in a particular data source to the master data source (combined recipient list) for the publication.

Before running this macro, replace  _datasourceindex_ with the index number of a valid data source in the data source collection of the active document, and replace _fieldname_ with the name of the field in the data source that you want to add to the combined list of recipient fields.

See the  **[Item](mailmergedatasources-item-method-publisher.md)** method topic for an example of how you can use the **Name** property of the **DataSource** object to determine the index number of the data source you want.




```vb
Public Sub AddToRecipientFields_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataField As Publisher.MailMergeDataField 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 Set pubMailMergeDataField = pubMailMergeDataSources.Item(datasourceindex).DataFields.Item("fieldname") 
 
 If pubMailMergeDataField.IsMapped Then 
 
 Debug.Print "This field is already mapped!" 
 
 Else 
 
 pubMailMergeDataField.AddToRecipientFields 
 Debug.Print "Field added successfully. (You can verify this by looking at the recipient or product list in the UI.)" 
 
 End If 
 
End Sub
```


