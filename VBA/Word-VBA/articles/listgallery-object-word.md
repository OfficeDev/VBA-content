---
title: ListGallery Object (Word)
keywords: vbawd10.chm2452
f1_keywords:
- vbawd10.chm2452
ms.prod: word
api_name:
- Word.ListGallery
ms.assetid: 4fa3af33-becd-0dfc-5c7a-a0e70714e045
ms.date: 06/08/2017
---


# ListGallery Object (Word)

Represents a single gallery of list formats. The  **ListGallery** object is a member of the **ListGalleries** collection.


## Remarks

Each  **ListGallery** object represents one of the three tabs in the **Bullets and Numbering** dialog box.

Use  **ListGalleries** (Index), where Index is **wdBulletGallery** , **wdNumberGallery** , or **wdOutlineNumberGallery** , to return a single **ListGallery** object.

The following example returns the third list format (excluding  **None**) on the  **Bulleted** tab in the **Bullets and Numbering** dialog box and then applies it to the selection.




```vb
Set temp3 = ListGalleries(wdBulletGallery).ListTemplates(3) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:= temp3
```

To see whether the specified list template contains the formatting built into Word, use the  **Modified** property for the **ListGallery** object. To reset formatting to the original list format, use the **Reset** method for the **ListGallery** object.


## Methods



|**Name**|
|:-----|
|[Reset](listgallery-reset-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](listgallery-application-property-word.md)|
|[Creator](listgallery-creator-property-word.md)|
|[ListTemplates](listgallery-listtemplates-property-word.md)|
|[Modified](listgallery-modified-property-word.md)|
|[Parent](listgallery-parent-property-word.md)|


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


