---
title: TextColumns.Add Method (Word)
keywords: vbawd10.chm158531785
f1_keywords:
- vbawd10.chm158531785
ms.prod: word
api_name:
- Word.TextColumns.Add
ms.assetid: 09e01558-9efc-ac84-684b-63ce459705fd
ms.date: 06/08/2017
---


# TextColumns.Add Method (Word)

Returns a  **TextColumn** object that represents a new text column added to a section or document.


## Syntax

 _expression_ . **Add**( **_Width_** , **_Spacing_** , **_EvenlySpaced_** )

 _expression_ Required. A variable that represents a **[TextColumns](textcolumns-objectword.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Width_|Optional| **Variant**|The width of the new text column in the document, in points.|
| _Spacing_|Optional| **Variant**|The spacing between the text columns in the document, in points.|
| _EvenlySpaced_|Optional| **Variant**| **True** to evenly space all the text columns be in the document.|

### Return Value

TextColumn


## Example

This example creates a new document and then adds another 2.5-inch-wide text column to it.


```vb
Set myDoc = Documents.Add 
myDoc.PageSetup.TextColumns.Add Width:=InchesToPoints(2.5), _ 
 Spacing:=InchesToPoints(0.5), EvenlySpaced:=False
```

This example adds a new text column to the active document and then evenly spaces all the text columns in the document.




```vb
ActiveDocument.PageSetup.TextColumns.Add _ 
 Width:=InchesToPoints(1.5), _ 
 EvenlySpaced:=True
```


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

