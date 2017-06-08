---
title: Folder.GetStorage Method (Outlook)
keywords: vbaol11.chm2017
f1_keywords:
- vbaol11.chm2017
ms.prod: outlook
api_name:
- Outlook.Folder.GetStorage
ms.assetid: cc5ee63b-7d11-6340-8392-8b35a689a28c
ms.date: 06/08/2017
---


# Folder.GetStorage Method (Outlook)

Gets a  **[StorageItem](storageitem-object-outlook.md)** object on the parent **[Folder](folder-object-outlook.md)** to store data for an Outlook solution.


## Syntax

 _expression_ . **GetStorage**( **_StorageIdentifier_** , **_StorageIdentifierType_** )

 _expression_ A variable that represents a **Folder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StorageIdentifier_|Required| **String**|An identifier for the  **StorageItem** object; depending on the identifier type, the value can represent an Entry ID, a message class, or a subject.|
| _StorageIdentifierType_|Required| **[OlStorageIdentifierType](olstorageidentifiertype-enumeration-outlook.md)**|Specifies the type of identifier for the  **StorageItem** object.|

### Return Value

A  **StorageItem** object that is used to store data for a solution.


## Remarks

The  **GetStorage** method obtains a **StorageItem** on a **Folder** object using the identifier specified by _StorageIdentifier_ and has the identifier type specified by _StorageIdentifierType_ . The **StorageItem** is a hidden item in the **Folder** , which roams with the account and is available online and offline.

If you specify the  **[EntryID](storageitem-entryid-property-outlook.md)** for the **StorageItem** by using the **olIdentifyByEntryID** value for _StorageIdentifierType_ , then the **GetStorage** method will return the **StorageItem** with the specified **EntryID** . If no **StorageItem** can be found using that **EntryID** or if the **StorageItem** does not exist, then the **GetStorage** method will raise an error.

If you specify the message class for the  **StorageItem** by using the **olIdentifyByMessageClass** value for _StorageIdentifierType_ , then the **GetStorage** method will return the **StorageItem** with the specified message class. If there are multiple items with the same message class, then the **GetStorage** method returns the item with the most recent **PR_LAST_MODIFICATION_TIME** . If no **StorageItem** exists with the specified message class, then the **GetStorage** method creates a new **StorageItem** with the message class specified by _StorageIdentifier_ .

If you specify the  **[Subject](storageitem-subject-property-outlook.md)** of the **StorageItem** , then the **GetStorage** method will return the **StorageItem** with the **Subject** specified in the **GetStorage** call. If there are multiple items with the same **Subject** , then the **GetStorage** method will return the item with the most recent **PR_LAST_MODIFICATION_TIME** . If no **StorageItem** exists with the specified **Subject** , then the **GetStorage** method will create a new **StorageItem** with the **Subject** specified by _StorageIdentifier_ .

 **GetStorage** returns an error if the store type of the folder is not supported. The following stores return an error when **GetStorage** is called:


- Hotmail store
    
- Internet Message Access Protocol (IMAP) stores
    
- Delegate stores
    
- Public folder stores
    


The  **[Size](storageitem-size-property-outlook.md)** of a **StorageItem** that is newly created is zero (0) until you make an explicit call on the **[Save](storageitem-save-method-outlook.md)** method of the item.

For more information on storing data for a solution, see [Storing Data for Solutions](http://msdn.microsoft.com/library/58e69983-5718-4dde-64fc-858abd80c9e5%28Office.15%29.aspx).


## Example

The following code sample in Visual Basic for Applications shows how to use the  **StorageItem** object to store private solution data. It saves the data in a custom property of a **StorageItem** object in the Inbox folder. The following describes the steps:


1. The code sample calls  **GetStorage** to obtain an existing **StorageItem** object that has the subject "My Private Storage" in the Inbox. If no **StorageItem** with that subject already exists, **GetStorage** creates a **StorageItem** object with that subject.
    
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


[Folder Object](folder-object-outlook.md)

