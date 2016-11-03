
# ListGalleries Object (Word)

A collection of  **[ListGallery](4fa3af33-becd-0dfc-5c7a-a0e70714e045.md)** objects that represent the three tabs in the **Bullets and Numbering** dialog box.


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
|[Item](df43ee1c-5834-c002-9e53-458f404f8b53.md)|

## Properties



|**Name**|
|:-----|
|[Application](1e6c3078-3024-ebad-be2a-9d1c7ea8b497.md)|
|[Count](bab7df3a-51f7-79fe-6d3d-f665dd23b7cf.md)|
|[Creator](2c24a4a7-b109-0b50-483a-b118b76ed731.md)|
|[Parent](71c4e3b7-0aa3-2f3c-7dd7-978f183b525b.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)