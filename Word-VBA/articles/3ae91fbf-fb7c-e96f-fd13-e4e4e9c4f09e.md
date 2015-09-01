
# ListGalleries Object (Word)

 **Last modified:** July 28, 2015

A collection of  ** [ListGallery](4fa3af33-becd-0dfc-5c7a-a0e70714e045.md)** objects that represent the three tabs in the **Bullets and Numbering** dialog box.

## Remarks

Use the  **ListGalleries** property to return the **ListGalleries** collection. The following code example enumerates the collection of list galleries and sets each of the seven list templates (formats) back to the list template format built into Word.


```
For Each lg In ListGalleries 
 For x = 1 To 7 
 lg.Reset(x) 
 Next x 
Next lg
```

Use  **ListGalleries**(Index), where Index is  **wdBulletGallery**,  **wdNumberGallery**, or  **wdOutlineNumberGallery**, to return a single  **ListGallery** object.

The following code example returns the third list format (excluding  **None**) on the  **Bulleted** tab in the **Bullets and Numbering** dialog box and then applies it to the selection.




```
Set temp3 = ListGalleries(wdBulletGallery).ListTemplates(3) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:= temp3
```

To see whether the specified list template contains the formatting built into Word, use the  **Modified** property with the **ListGallery** object. To reset formatting to the original list format, use the **Reset** method for the **ListGallery** object.


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [ListGalleries Object Members](c68a29b8-af7f-9863-8501-829d18511a61.md)
