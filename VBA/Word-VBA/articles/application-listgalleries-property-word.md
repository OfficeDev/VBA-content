---
title: Application.ListGalleries Property (Word)
keywords: vbawd10.chm158335041
f1_keywords:
- vbawd10.chm158335041
ms.prod: word
api_name:
- Word.Application.ListGalleries
ms.assetid: 769d3494-3fc3-5a4b-e6d1-a3910107c8bd
ms.date: 06/08/2017
---


# Application.ListGalleries Property (Word)

Returns a  **[ListGalleries](listgalleries-object-word.md)** collection that represents the three list template galleries. .


## Syntax

 _expression_ . **ListGalleries**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Each template gallery (Bulleted, Numbered, and Outline Numbered) corresponds to a tab in the  **Bullets and Numbering** dialog box ( **Format** menu). For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the variable mylsttmp to the second list template on the  **Outline Numbered** tab in the **Bullets and Numbering** dialog box. The example then applies that template to the first list in the active document.


```vb
Set mylsttmp = _ 
 ListGalleries(wdOutlineNumberGallery).ListTemplates(2) 
ActiveDocument.Lists(1).ApplyListTemplate ListTemplate:=mylsttmp
```

This example cycles through the  **ListGalleries** collection and changes the templates in each list template gallery back to the built-in template.




```vb
For Each listgal In ListGalleries 
 For i = 1 To 7 
 listgal.Reset(i) 
 Next i 
Next listgal
```


## See also


#### Concepts


[Application Object](application-object-word.md)

