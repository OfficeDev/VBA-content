---
title: Folder.Delete Method (Outlook)
keywords: vbaol11.chm1995
f1_keywords:
- vbaol11.chm1995
ms.prod: outlook
api_name:
- Outlook.Folder.Delete
ms.assetid: 3df0f063-3f41-e3b7-d1e3-7ea08970c56d
ms.date: 06/08/2017
---


# Folder.Delete Method (Outlook)

Deletes an object from the collection.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **Folder** object.


## Remarks

The  **Delete** method deletes a single folder.

In general, deleting a folder does not require first deleting the items in the folder. Deleting the folder also deletes all items in the folder. An exception would be if the folder is an Outlook folder that cannot be deleted, such as the Inbox and the Deleted Items folder. In such cases, you can delete only the items of the folder but not the folder itself. To delete all items in the  **[Items](folder-items-property-outlook.md)** collection of the folder, you must delete each item starting with the last item in the folder. For example, in the items collection of a folder, `AllItems`, if there are  `n` number of items in the folder, start deleting the item at `AllItems.Item(n)`, decrementing the index each time until you delete  `AllItems.Item(1)`.


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

