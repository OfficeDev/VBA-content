---
title: Updating and Deleting Solution Storage
ms.prod: outlook
ms.assetid: ac1b1e9f-25d2-4157-c237-318e2e7c5f6b
ms.date: 06/08/2017
---


# Updating and Deleting Solution Storage

This topic describes updating and deleting solution storage.


## Updating Solution Storage

You can store private solution data as an attachment or as the value of a property of a  **[StorageItem](storageitem-object-outlook.md)** object in most folders. Since it is possible for multiple solutions to share the same solution storage, after updating an attachment or property, a solution should call ** [StorageItem.Save](storageitem-save-method-outlook.md)** to update the item in the folder. In cases where there is more than one solution accessing the same object, the object would always show the updates through the most recent **StorageItem.Save**.

Attempting to save to a  **StorageItem** object that has been deleted will result in the error, "Unable to perform the operation."


## Deleting Solution Storage

Solutions can remove a  **StorageItem** object by calling ** [StorageItem.Delete](storageitem-delete-method-outlook.md)**. This call permanently removes the object from the folder; it does not move it to the  **Deleted Items** folder. This allows a solution to clean up or reset the storage for its private data.

Attempting to delete a  **StorageItem** that has been removed by a prior **StorageItem.Delete** call will result in the error, "Could not complete the deletion."


 **Note**  Solution storage can only be removed through the  **Delete** method of the corresponding **StorageItem** object. If the creator solution has been uninstalled and there is no other solution that can access the object to delete it, the object will remain as a hidden item in the folder for as long as the folder exists.


