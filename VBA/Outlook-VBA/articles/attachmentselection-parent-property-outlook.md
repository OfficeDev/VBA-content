---
title: AttachmentSelection.Parent Property (Outlook)
keywords: vbaol11.chm2942
f1_keywords:
- vbaol11.chm2942
ms.prod: outlook
api_name:
- Outlook.AttachmentSelection.Parent
ms.assetid: 1c80c1fd-b7bd-288c-d017-8159ddcbd037
ms.date: 06/08/2017
---


# AttachmentSelection.Parent Property (Outlook)

Returns the parent  **Object** of the specified object. Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents an **[AttachmentSelection](attachmentselection-object-outlook.md)** object.


## Remarks

The  **Parent** property of an **AttachmentSelection** object represents the Microsoft Outlook item that contains the selected attachments.

If the item is in an explorer, the value of the  **Parent** property is the same as the first item in the selection that is returned by the **[Explorer.Selection](explorer-selection-property-outlook.md)** property, which is `Explorer.Selection.Item(1)`. 

If the item is in an inspector, the value of the  **Parent** property is the same as the value of the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property.


## See also


#### Concepts


[AttachmentSelection Object](attachmentselection-object-outlook.md)

