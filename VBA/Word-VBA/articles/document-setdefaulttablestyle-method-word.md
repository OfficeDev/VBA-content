---
title: Document.SetDefaultTableStyle Method (Word)
keywords: vbawd10.chm158007662
f1_keywords:
- vbawd10.chm158007662
ms.prod: word
api_name:
- Word.Document.SetDefaultTableStyle
ms.assetid: 6e932b12-6af8-af0a-5c3b-c74cefaf0d35
ms.date: 06/08/2017
---


# Document.SetDefaultTableStyle Method (Word)

Specifies the table style to use for newly created tables in a document.


## Syntax

 _expression_ . **SetDefaultTableStyle**( **_Style_** , **_SetInTemplate_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **Variant**|A string specifying the name of the style.|
| _SetInTemplate_|Required| **Boolean**| **True** to save the table style in the template attached to the document.|

## Example

This example checks to see if the default table style used in the active document is named Table Normal and, if it is, changes the default table style to TableStyle1. This example assumes that you have a table style named TableStyle1.


```vb
Sub TableDefaultStyle() 
 With ActiveDocument 
 If .DefaultTableStyle = "Table Normal" Then 
 .SetDefaultTableStyle Style:="TableStyle1", _ 
 SetInTemplate:=True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

