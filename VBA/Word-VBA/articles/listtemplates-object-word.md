---
title: ListTemplates Object (Word)
ms.prod: word
ms.assetid: 5b5f3ed8-4522-f52e-5ae8-9df26a7da154
ms.date: 06/08/2017
---


# ListTemplates Object (Word)

A collection of  **[ListTemplate](listtemplate-object-word.md)** objects in a document, list gallery, or template.


## Remarks

Use the  **ListTemplates** property with a [Document](document-object-word.md), [ListGallery](listgallery-object-word.md), or [Template](template-object-word.md) object to return a **ListTemplates** collection. With a ListGallery object, the ListTemplates collection is the seven list formats for bulleted lists, numbered lists, and outline numbered lists. 
 
The following example displays a message with the level status (single or multiple-level) for each list template in the active document.


```vb
For Each lt In ActiveDocument.ListTemplates 
 MsgBox "This is a multiple-level list template - " _ 
 & lt.OutlineNumbered 
Next lt
```

Use the  **Add** method to add a list template to the collection in the specified document or template. The following example adds a new list template to the active document and applies it to the selection.




```vb
Set myLT = ActiveDocument.ListTemplates.Add 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=myLT
```

Use  **ListTemplates** (Index), where Index is the name of a list template or an index number, to return a single list template in a document or template. The following example sets an object variable equal to a list template named "ListBullets" in the active document, and then formats the selection as the first level of that list template. 


```vb
Set mylt = ActiveDocument.ListTemplates("ListBullets")
Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=mylt, ApplyLevel:=1
```

Use  **ListTemplates** (Index), where Index is a number 1 through 7, to return a single list template in a list gallery. The following example sets an object variable equal to the first list template in the bullet list gallery, and then it applies that list template to the selection.




```vb
Set mylt = ListGalleries(wdBulletGallery).ListTemplates(1) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=mylt
```


> **Note**  Some properties and methods —  **Convert** and **Add** , for example — won't work with the list templates in a list gallery. You can modify those list templates, but you cannot change their list gallery type ( **wdBulletGallery** , **wdNumberGallery** , or **wdOutlineNumberGallery** ).

To see whether a list template in a list gallery contains the formatting built into Word, use the  **[Modified](listgallery-modified-property-word.md)** property with the **ListGallery** object. To reset formatting to the original list format, use the **[Reset](listgallery-reset-method-word.md)** method for the **ListGallery** object.

After you have returned a  **[ListTemplate](listtemplate-object-word.md)** object, use **ListLevels** (Index), where Index is a number from 1 through 9, to return a single **ListLevel** object. With a **ListLevel** object, you have access to all the formatting properties for the specified list level, such as **Alignment** , **Font** , **NumberFormat** , **NumberPosition** , **NumberStyle** , and **TrailingCharacter** .

Use the  **Convert** method to convert a multiple-level list template to a single-level template.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

