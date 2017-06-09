---
title: XMLSchemaReferences.HideValidationErrors Property (Word)
keywords: vbawd10.chm116129797
f1_keywords:
- vbawd10.chm116129797
ms.prod: word
api_name:
- Word.XMLSchemaReferences.HideValidationErrors
ms.assetid: a31185b6-f179-acf8-d5ee-57311dca4c34
ms.date: 06/08/2017
---


# XMLSchemaReferences.HideValidationErrors Property (Word)

Returns a  **Boolean** indicating whether Word displays schema validation errors for the current XML document. Read/write.


## Syntax

 _expression_ . **HideValidationErrors**

 _expression_ An expression that returns an **[XMLSchemaReferences](xmlschemareferences-object-word.md)** collection.


## Remarks

 **True** causes Word to hide schema validation errors for the current XML document. **False** causes schema validation errors to be displayed in the **XML Structure** task pane.


## Example

The following example disables the display of schema validation errors in the current XML document.


```vb
ActiveDocument.XMLSchemaReferences _ 
 .HideValidationErrors = True
```


## See also


#### Concepts


[XMLSchemaReferences Collection](xmlschemareferences-object-word.md)

