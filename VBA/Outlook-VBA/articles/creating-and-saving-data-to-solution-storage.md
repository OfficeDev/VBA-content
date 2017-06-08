---
title: Creating and Saving Data to Solution Storage
ms.prod: outlook
ms.assetid: 5a417191-ed36-be5c-5d63-1ab618bd06cf
ms.date: 06/08/2017
---


# Creating and Saving Data to Solution Storage

This topic describes creating or using existing storage to store private solution data.

The Outlook object model supports creating and storing solution data as hidden items in a folder. You can use  **[Folder.GetStorage](folder-getstorage-method-outlook.md)** to create a **[StorageItem](storageitem-object-outlook.md)** object in a specified folder. You can identify this object by the subject, message class, or Entry ID. Solutions can create **StorageItem** objects in all folders except when:

- The folder is a Microsoft Exchange public folder, an Internet Message Access Protocol (IMAP), MSN Hotmail, or a Microsoft SharePoint Foundation folder.
    
- The user permission for the folder is read-only.
    
- The store provider does not support hidden items.
    

In these cases,  **Folder.GetStorage** will return an error: "Cannot create StorageItem in this folder."
When you call  **Folder.GetStorage** specifying a subject or a message class and the specified item does not exist in the folder, the call creates and returns a **StorageItem** object with the message class **IPM.Storage**; if you specified an Entry ID, howwever, the call will return the error, "The operation failed. An object could not be found."

## Obtaining an Existing StorageItem

You can call  **Folder.GetStorage** for an item that already exists in a folder. For example, the item can be one that the solution has previously created; it can be an item with a well-known message class such as **IPC.MS.Outlook.AgingProperties**, or an item that existed as a hidden message in the folder in a previous version of Outlook. In these cases, the call will return a  **StorageItem** object representing the item. The message class of the item however will not change.

 If you call **Folder.GetStorage** specifying a subject or message class and more than one item exists in the folder, then the call returns the item that was last modified (that is, the item with the most recent **PidTagLastModificationTime**).


## Storing Data in a StorageItem

After obtaining a  **StorageItem** object, you can store private data as an attachment to the item, or as a value to the **Body** property or a custom property of the item. The initial size of the item is 0. As you store data to the item, the **[StorageItem.Size](storageitem-size-property-outlook.md)** is updated. Call **[StorageItem.Save](storageitem-save-method-outlook.md)** to update the contents of the item in the folder.


