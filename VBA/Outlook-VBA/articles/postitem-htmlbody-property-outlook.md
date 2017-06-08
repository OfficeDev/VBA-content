---
title: PostItem.HTMLBody Property (Outlook)
keywords: vbaol11.chm1548
f1_keywords:
- vbaol11.chm1548
ms.prod: outlook
api_name:
- Outlook.PostItem.HTMLBody
ms.assetid: 5db93b3c-96b0-ce14-4d53-cbc113c2c14c
ms.date: 06/08/2017
---


# PostItem.HTMLBody Property (Outlook)

Returns or sets a  **String** representing the HTML body of the specified item. Read/write.


## Syntax

 _expression_ . **HTMLBody**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

The  **HTMLBody** property should be an HTML syntax string.

Setting the  **HTMLBody** property sets the **[EditorType](inspector-editortype-property-outlook.md)** property of the item's **[Inspector](inspector-object-outlook.md)** to **olEditorHTML** .

Setting the  **HTMLBody** property will always update the **[Body](postitem-body-property-outlook.md)** property immediately.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

