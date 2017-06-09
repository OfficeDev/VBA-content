---
title: Storing Data for Solutions
ms.prod: outlook
ms.assetid: 58e69983-5718-4dde-64fc-858abd80c9e5
ms.date: 06/08/2017
---


# Storing Data for Solutions

This topic describes using the  **[StorageItem](storageitem-object-outlook.md)** object as a means for developers to store private solution data.

Outlook solution developers often require a place to store and maintain private application data. For example, you may need to store an incrementing order number. The Outlook object model provides the  **StorageItem** object to store this private data.

The  **StorageItem** object represents a thin wrapper on a message object in MAPI (the **IMessage** object). It is always saved to the associated portion of its parent MAPI folder so that the item is hidden in the folder. It is a child object of the **[Folder](folder-object-outlook.md)** object. This means that solution private data is actually stored at the folder level, allowing the data to roam with the mailbox and be available online and offline.

You can identify a  **StorageItem** object using its subject, message class, or Entry ID. A **StorageItem** is not tightly bound to only one solution. This allows you to create one or more **StorageItem** objects in one folder or in multiple folders. Instances of the same solution, or multiple collaborating solutions, can also share the data stored in the private storage.
You can create a  **StorageItem** or get an existing **StorageItem** to store solution data. You can store the data as an attachment or a value to an item property. To clean up the storage for an application, you can delete the **StorageItem** objects that it uses, which removes these objects permanently.
The Outlook object model does not provide any collection object for  **StorageItem** objects. However, you can use **[Folder.GetTable](folder-gettable-method-outlook.md)** to obtain a **[Table](table-object-outlook.md)** with all the hidden items in a **Folder**, when you specify the  _TableContents_ parameter as **olHiddenItems**. If keeping your data private is of a high concern, you should encrypt the data before storing it.

