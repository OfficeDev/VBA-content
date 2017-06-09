---
title: AutoCaption Object (Word)
keywords: vbawd10.chm2427
f1_keywords:
- vbawd10.chm2427
ms.prod: word
api_name:
- Word.AutoCaption
ms.assetid: 895b5181-d36f-7f63-572a-c2d37c878e17
ms.date: 06/08/2017
---


# AutoCaption Object (Word)

Represents a single caption that can be automatically added when items such as tables, pictures, or OLE objects are inserted into a document. The  **AutoCaption** object is a member of the **[AutoCaptions](autocaptions-object-word.md)** collection. The **AutoCaptions** collection contains all the captions listed in the **AutoCaption** dialog box.


## Remarks

Use  **[AutoCaptions](application-autocaptions-property-word.md)** (index), where index is the caption name or index number, to return a single **AutoCaption** object. The caption names correspond to the items listed in the **AutoCaption** dialog box. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the **AutoCaption** dialog box. The following example enables autocaptions for Word tables.


```vb
AutoCaptions("Microsoft Word Table").AutoInsert = True
```

The index number represents the position of the  **AutoCaption** object in the list of items in the **AutoCaption** dialog box. The following example displays the name of the first item listed in the **AutoCaption** dialog box.




```vb
MsgBox AutoCaptions(1).Name
```

 **AutoCaption** objects cannot be programmatically added to or deleted from the **AutoCaptions** collection.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


