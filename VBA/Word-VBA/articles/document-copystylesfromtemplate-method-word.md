---
title: Document.CopyStylesFromTemplate Method (Word)
keywords: vbawd10.chm158007422
f1_keywords:
- vbawd10.chm158007422
ms.prod: word
api_name:
- Word.Document.CopyStylesFromTemplate
ms.assetid: f02fbce7-f5aa-d71d-9043-f151f26bc9ec
ms.date: 06/08/2017
---


# Document.CopyStylesFromTemplate Method (Word)

Copies styles from the specified template to a document.


## Syntax

 _expression_ . **CopyStylesFromTemplate**( **_Template_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Template_|Required| **String**|The template file name.|

## Remarks

When styles are copied from a template to a document, like-named styles in the document are redefined to match the style descriptions in the template. Unique styles from the template are copied to the document. Unique styles in the document remain intact.


## Example

This example copies the styles from the active document's template to the document.


```vb
ActiveDocument.CopyStylesFromTemplate _ 
 Template:=ActiveDocument.AttachedTemplate.FullName
```

This example copies the styles from the Sales96.dot template to Sales.doc.




```
Documents("Sales.doc").CopyStylesFromTemplate _ 
 Template:="C:\MSOffice\Templates\Sales96.dot"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

