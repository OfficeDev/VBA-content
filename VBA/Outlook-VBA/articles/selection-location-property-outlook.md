---
title: Selection.Location Property (Outlook)
keywords: vbaol11.chm3481
f1_keywords:
- vbaol11.chm3481
ms.prod: outlook
api_name:
- Outlook.Selection.Location
ms.assetid: 8a2db72a-8db0-840e-349e-5d9d22f3affb
ms.date: 06/08/2017
---


# Selection.Location Property (Outlook)

Returns an  **[OlSelectionLocation](olselectionlocation-enumeration-outlook.md)** constant that specifies where in the Microsoft Outlook user interface the current selection is. Read-only


## Syntax

 _expression_ . **Location**

 _expression_ A variable that represents a **[Selection](selection-object-outlook.md)** object.


## Remarks

A  **Location** property with the value **olViewList** means that the current selection is in a list of items in an explorer. Calling **[Selection.GetSelection](selection-getselection-method-outlook.md)** with **olConversationHeaders** as the argument returns a **Selection** object with **[Selection.Count](selection-count-property-outlook.md)** equal to the number of conversation headers in the current selection.

If the  **Location** property is not equal to **olViewList** , calling **GetSelection** with **olConversationHeaders** as the argument returns a **Selection** object with **Selection.Count** equal to 0.


## See also


#### Concepts


[Selection Object](selection-object-outlook.md)

