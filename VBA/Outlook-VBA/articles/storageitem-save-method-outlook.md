---
title: StorageItem.Save Method (Outlook)
keywords: vbaol11.chm2144
f1_keywords:
- vbaol11.chm2144
ms.prod: outlook
api_name:
- Outlook.StorageItem.Save
ms.assetid: 9462a342-294a-175e-7e8f-d416f0959f69
ms.date: 06/08/2017
---


# StorageItem.Save Method (Outlook)

Saves the  **[StorageItem](storageitem-object-outlook.md)** .


## Syntax

 _expression_ . **Save**

 _expression_ A variable that represents a **StorageItem** object.


## Remarks

If the  **StorageItem** has never been saved before, **Save** saves the item as a hidden item in the **[Folder](folder-object-outlook.md)** on which **[Folder.GetStorage](folder-getstorage-method-outlook.md)** was called. If the **StorageItem** has been saved previously and the item has since been changed, **Save** saves the changes to the item. If the **StorageItem** has been saved previously and the item has not been changed since then, the **Save** method does nothing.

For more information on saving solution data to a  **StorageItem** object, see[Creating and Saving Data to Solution Storage](http://msdn.microsoft.com/library/5a417191-ed36-be5c-5d63-1ab618bd06cf%28Office.15%29.aspx).


## Example

The following code sample in Visual Basic for Applications shows how to use the  **StorageItem** object to store private solution data. It saves the data in a custom property of a **StorageItem** object in the Inbox folder. The following describes the steps:


1. The code sample calls  **[Folder.GetStorage](folder-getstorage-method-outlook.md)** to obtain an existing **StorageItem** object that has the subject "My Private Storage" in the Inbox; if no **StorageItem** with that subject already exists, **GetStorage** creates a **StorageItem** object with that subject.
    
2. If the  **StorageItem** is newly created, the code sample creates a custom property "Order Number" for the object. Note that "Order Number" is a property of a hidden item in the Inbox.
    
3. The code sample then assigns a value to "Order Number" and saves the  **StorageItem** object.
    





```vb
Sub AssignStorageData() 
 
 Dim oInbox As Outlook.Folder 
 
 Dim myStorage As Outlook.StorageItem 
 
 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 ' Get an existing instance of StorageItem, or create new if it doesn't exist 
 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 ' If StorageItem is new, add a custom property for Order Number 
 
 If myStorage.Size = 0 Then 
 
 myStorage.UserProperties.Add "Order Number", olNumber 
 
 End If 
 
 ' Assign a value to the custom property 
 
 myStorage.UserProperties("Order Number").Value = 100 
 
 myStorage.Save 
 
End Sub
```


## See also


#### Concepts


[StorageItem Object](storageitem-object-outlook.md)

