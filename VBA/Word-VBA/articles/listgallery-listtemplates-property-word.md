---
title: ListGallery.ListTemplates Property (Word)
keywords: vbawd10.chm160694273
f1_keywords:
- vbawd10.chm160694273
ms.prod: word
api_name:
- Word.ListGallery.ListTemplates
ms.assetid: 459297de-c2b6-23f8-8670-7c81d8f577c8
ms.date: 06/08/2017
---


# ListGallery.ListTemplates Property (Word)

Returns a  **ListTemplates** collection that represents all the list formats for the specified list gallery. Read-only.


## Syntax

 _expression_ . **ListTemplates**

 _expression_ A variable that represents a **[ListGallery](listgallery-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx). The ListTemplates property is a member of the [Document](document-object-word.md), [ListGallery](listgallery-object-word.md), and [Template](template-object-word.md) objects.


## Example

This example sets the variable  _mytemp_ to the first list template on the **Outline Numbered** tab in the **Bullets and Numbering** dialog box. The template is modified to use lowercase letters for each level, and it is applied to the second list in the active document.


```vb
Set mytemp = ListGalleries(wdOutlineNumberGallery).ListTemplates(1) 
For each lev in mytemp.ListLevels 
 lev.NumberStyle = wdListNumberStyleLowercaseLetter 
Next lev 
ActiveDocument.Lists(2).ApplyListTemplate ListTemplate:=mytemp
```


## See also


#### Concepts


[ListGallery Object](listgallery-object-word.md)

