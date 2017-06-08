---
title: ListGalleries Object (Word)
ms.prod: word
ms.assetid: 3ae91fbf-fb7c-e96f-fd13-e4e4e9c4f09e
ms.date: 06/08/2017
---


# ListGalleries Object (Word)

A collection of  **[ListGallery](listgallery-object-word.md)** objects that represent the three tabs in the **Bullets and Numbering** dialog box.


## Remarks

Use the  **ListGalleries** property to return the **ListGalleries** collection. The following code example enumerates the collection of list galleries and sets each of the seven list templates (formats) back to the list template format built into Word.


```
For Each lg In ListGalleries 
 For x = 1 To 7 
 lg.Reset(x) 
 Next x 
Next lg
```

Use  **ListGalleries** (Index), where Index is **wdBulletGallery**, **wdNumberGallery**, or **wdOutlineNumberGallery**, to return a single **ListGallery** object.

The following code example returns the third list format (excluding  **None**) on the  **Bulleted** tab in the **Bullets and Numbering** dialog box and then applies it to the selection.




```
Set temp3 = ListGalleries(wdBulletGallery).ListTemplates(3) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:= temp3
```

To see whether the specified list template contains the formatting built into Word, use the  **Modified** property with the **ListGallery** object. To reset formatting to the original list format, use the **Reset** method for the **ListGallery** object.


## Methods



|**Name**|
|:-----|
|[Item](listgalleries-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](listgalleries-application-property-word.md)|
|[Count](listgalleries-count-property-word.md)|
|[Creator](listgalleries-creator-property-word.md)|
|[Parent](listgalleries-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
