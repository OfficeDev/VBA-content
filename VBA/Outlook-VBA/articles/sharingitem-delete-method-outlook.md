---
title: SharingItem.Delete Method (Outlook)
keywords: vbaol11.chm625
f1_keywords:
- vbaol11.chm625
ms.prod: outlook
api_name:
- Outlook.SharingItem.Delete
ms.assetid: 9848fe0e-b32f-8796-f37d-7b7795309e1a
ms.date: 06/08/2017
---


# SharingItem.Delete Method (Outlook)

Removes a  **[SharingItem](sharingitem-object-outlook.md)** item from the folder that contains the item.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

The  **Delete** method deletes a single item in a collection. To delete all items in the **[Items](folder-items-property-outlook.md)** collection of a folder, you must delete each item starting with the last item in the folder. For example, in the items collection of a folder, `AllItems`, if there are  `n` number of items in the folder, start deleting the item at `AllItems.Item(n)`, decrementing the index each time until you delete  `AllItems.Item(1)`.

The  **Delete** method moves the item from the containing folder to the **Deleted Items** folder. If the containing folder is the **Deleted Items** folder, the **Delete** method removes the item permanently.


## See also


#### Concepts


[SharingItemObject](sharingitem-object-outlook.md)
#### Other resources



[Delete All Items and Subfolders in the Deleted Items Folder](http://msdn.microsoft.com/library/359a416b-43d4-396e-e348-5624c4ca3599%28Office.15%29.aspx)

