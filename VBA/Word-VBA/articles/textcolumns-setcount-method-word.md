---
title: TextColumns.SetCount Method (Word)
keywords: vbawd10.chm158531786
f1_keywords:
- vbawd10.chm158531786
ms.prod: word
api_name:
- Word.TextColumns.SetCount
ms.assetid: 59ff1b21-5bec-982d-a2b5-7a8d7dc08f9a
ms.date: 06/08/2017
---


# TextColumns.SetCount Method (Word)

Arranges text into the specified number of text columns.


## Syntax

 _expression_ . **SetCount**( **_NumColumns_** )

 _expression_ Required. A variable that represents a **[TextColumns](textcolumns-objectword.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumColumns_|Required| **Long**|The number of columns the text is to be arranged into.|

## Remarks

You can also use the  **[Add](textcolumns-add-method-word.md)** method to add a single column to the **TextColumns** collection.


## Example

This example arranges the text in the active document into two columns of equal width.


```vb
ActiveDocument.PageSetup.TextColumns.SetCount NumColumns:=2
```

This example arranges the text in the first section of Brochure.doc into three columns of equal width.




```
Documents("Brochure.doc").Sections(1) _ 
 .PageSetup.TextColumns.SetCount NumColumns:=3
```


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

