---
title: TextRange.InsertMailMergeField Method (Publisher)
keywords: vbapb10.chm5308483
f1_keywords:
- vbapb10.chm5308483
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertMailMergeField
ms.assetid: 97bce07d-b831-3ad6-2436-f85590c3bcd8
ms.date: 06/08/2017
---


# TextRange.InsertMailMergeField Method (Publisher)

Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a text data field for a mail merge or catalog merge.


## Syntax

 _expression_. **InsertMailMergeField**( **_varIndex_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|varIndex|Required| **Variant**|The name or index of the data field in the datasource.|

### Return Value

TextRange


## Remarks

For a publication's catalog merge area to contain text data fields, it must first contain at least one text box to contain the text data fields. 


## Example

This example inserts a  **LastName** field at the cursor position. This example assumes that the active publication is a mail merge publication and that the cursor position is somewhere inside a text box.


```vb
Sub InsertMergeField() 
 Selection.TextRange.InsertMailMergeField varIndex:="LastName" 
End Sub
```

This example adds a text box to the specified publication's catalog merge area, and then inserts a text data field into the text box. This example assumes that the specified publication is connected to a data source, and that it contains a catalog merge area.




```vb
Set pbTextBox1 = ThisDocument.Pages(1).Shapes.AddTextbox(1, 100, 100, 175, 25) 
pbTextBox1.AddToCatalogMergeArea 
 
With pbTextBox1.TextFrame.TextRange 
 .Text = "List Price: " 
 .InsertMailMergeField "List Price" 
End With 

```


