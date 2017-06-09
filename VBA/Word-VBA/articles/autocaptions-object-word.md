---
title: AutoCaptions Object (Word)
keywords: vbawd10.chm2426
f1_keywords:
- vbawd10.chm2426
ms.prod: word
ms.assetid: da4bd001-8f4c-28c9-4f46-a5a6499000a8
ms.date: 06/08/2017
---


# AutoCaptions Object (Word)

A collection of  **[AutoCaption](autocaption-object-word.md)** objects that represent the captions that can be automatically added when items such as tables, pictures, or OLE objects are inserted into a document.


## Remarks

Use the  **[AutoCaptions](application-autocaptions-property-word.md)** property to return the **AutoCaptions** collection. The following example displays the names of the selected items in the **AutoCaption** dialog box.


```vb
For Each autoCap In AutoCaptions 
 If autoCap.AutoInsert = True Then 
 MsgBox autoCap.Name &; " is configured for auto insert" 
 End If 
Next autoCap
```

The  **AutoCaptions** collection contains all the captions listed in the **AutoCaption** dialog box. **[AutoCaption](autocaption-object-word.md)** objects cannot be programmatically added to or deleted from the **AutoCaptions** collection.

Use  **AutoCaptions** (index), where index is the caption name or index number, to return a single **[AutoCaption](autocaption-object-word.md)** object. The caption names correspond to the items listed in the **AutoCaption** dialog box. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the **AutoCaption** dialog box. The following example displays the caption text "Microsoft Word Table."




```vb
MsgBox AutoCaptions("Microsoft Word Table").CaptionLabel.Name
```

The index number represents the position of the  **[AutoCaption](autocaption-object-word.md)** object in the list of captions in the **AutoCaption** dialog box. The following example displays the name of the first item selected in the **AutoCaption** dialog box.




```vb
MsgBox AutoCaptions(1).Name
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


