---
title: ListTemplate.OutlineNumbered Property (Word)
keywords: vbawd10.chm160366593
f1_keywords:
- vbawd10.chm160366593
ms.prod: word
api_name:
- Word.ListTemplate.OutlineNumbered
ms.assetid: 0d728c52-b33d-7764-a0ef-6573040ed1ef
ms.date: 06/08/2017
---


# ListTemplate.OutlineNumbered Property (Word)

 **True** if the specified **ListTemplate** object is outline numbered. Read/write **Boolean** .


## Syntax

 _expression_ . **OutlineNumbered**

 _expression_ An expression that returns a **[ListTemplate](listtemplate-object-word.md)** object.


## Remarks

Setting this property to  **False** converts the list template to a single-level list that uses the formatting of the first level.

You cannot set this property for a  **ListTemplate** object returned from a **[ListGallery](listgallery-object-word.md)** object.


## Example

This example changes the selected outline-numbered list to a single-level numbered list.


```vb
Selection.Range.ListFormat.ListTemplate.OutlineNumbered = False
```

This example checks to see whether the third list in MyDoc.doc is an outline-numbered list. If it is, the third outline-numbered list template is applied to it.




```vb
Set myltemp = Documents("MyDoc.doc").Lists(3).Range _ 
 .ListFormat.ListTemplate 
num = myltemp.OutlineNumbered 
If num = True Then ActiveDocument.Lists(3).ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(3)
```


## See also


#### Concepts


[ListTemplate Object](listtemplate-object-word.md)

