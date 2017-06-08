---
title: MailMergeDataSource.EditRecord Method (Publisher)
keywords: vbapb10.chm6291504
f1_keywords:
- vbapb10.chm6291504
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.EditRecord
ms.assetid: 1fa31b25-b00a-9478-b341-094c2cdb2d9e
ms.date: 06/08/2017
---


# MailMergeDataSource.EditRecord Method (Publisher)

Changes one of the data fields in one of the records in the master data source (the combined mail-merge recipient list).


## Syntax

 _expression_. **EditRecord**( **_lRec_**,  **_varField_**,  **_Value_**)

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|lRec|Required| **Long**|The ID of the record you want to edit.|
|varField|Required| **Variant**|The data field (column) that contains the value you want to change.|
|Value|Required| **Variant**|The value to be changed.|

## Remarks

You can use the  **EditRecord** method to correct data source information that is in error, such as an outdated recipient address.

The  **EditRecord** method does not make any changes to the individual data sources that together make up the master data source.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **EditRecord** method to change the value of a single column in a particular record in the master data source (the combined mail-merge recipient list).

Before running this macro, replace  _recordID_ with the index number of the record in the data source that you want to edit; replace _fieldname_ with the name of the field (column) in the record that you want to edit; and replace _value_ with the new value you want to set for the field.




```vb
Public Sub EditRecord_Example() 
 
 Dim pubMailMergeDataSource As Publisher.MailMergeDataSource 
 
 Set pubMailMergeDataSource = ThisDocument.MailMerge.DataSource 
 
 pubMailMergeDataSource.EditRecord recordID, "fieldname", "value" 
 
End Sub
```


