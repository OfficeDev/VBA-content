---
title: Global.ListGalleries Property (Word)
keywords: vbawd10.chm163119169
f1_keywords:
- vbawd10.chm163119169
ms.prod: word
api_name:
- Word.Global.ListGalleries
ms.assetid: 56ac5cc2-552a-cff6-95cb-40eebd904eb7
ms.date: 06/08/2017
---


# Global.ListGalleries Property (Word)

Returns a  **ListGalleries** collection that represents the three list template galleries ( **Bulleted**,  **Numbered**, and  **Outline Numbered**).


## Syntax

 _expression_ . **ListGalleries**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

Each gallery corresponds to a tab in the  **Bullets and Numbering** dialog box. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the variable  _mylsttmp_ to the second list template on the **Outline Numbered** tab in the **Bullets and Numbering** dialog box. The example then applies that template to the first list in the active document.


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


[Global Object](global-object-word.md)

