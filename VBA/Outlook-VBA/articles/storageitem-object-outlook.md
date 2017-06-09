---
title: StorageItem Object (Outlook)
keywords: vbaol11.chm3017
f1_keywords:
- vbaol11.chm3017
ms.prod: outlook
api_name:
- Outlook.StorageItem
ms.assetid: 41776bc3-b838-2755-fd6b-3b5012fb9ae5
ms.date: 06/08/2017
---


# StorageItem Object (Outlook)

A message object in MAPI that is always saved as a hidden item in the parent folder and stores private data for Outlook solutions.


## Remarks

A  **StorageItem** object is stored at the folder level, allowing it to roam with the account and be available online or offline.

The Outlook object model does not provide any collection object for  **StorageItem** objects. However, you can use **[Folder.GetTable](folder-gettable-method-outlook.md)** to obtain a **[Table](table-object-outlook.md)** with all the hidden items in a **[Folder](folder-object-outlook.md)**, when you specify the _TableContents_ parameter as **olHiddenItems**. If keeping your data private is of a high concern, you should encrypt the data before storing it.

Once you have obtained a  **StorageItem** object, you can do the following to store solution data:


- Add attachments to the item for storage.
    
- Use explicit built-in properties of the item such as  **[Body](storageitem-body-property-outlook.md)** to store custom data.
    
- Add custom properties to the item using  **[UserProperties.Add](userproperties-add-method-outlook.md)** method. Note that in this case, the optional _AddToFolderFields_ and _DisplayFormat_ arguments of the **UserProperties.Add** method will be ignored.
    
- Use the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object to get or set custom properties.
    


The default message class for a new  **StorageItem** is **IPM.Storage**. If the **StorageItem** existed as a hidden message in a version of Outlook prior to Microsoft Office Outlook 2007, the message class will remain unchanged. In order to prevent modification of the message class, **StorageItem** does not expose an explicit **MessageClass** property.

For more information on storing solution data using the  **StorageItem** object, see[Storing Data for Solutions](http://msdn.microsoft.com/library/58e69983-5718-4dde-64fc-858abd80c9e5%28Office.15%29.aspx).


## Example

The following code sample in Visual Basic for Applications shows how to use the  **StorageItem** object to store private solution data. It saves the data in a custom property of a **StorageItem** object in the Inbox folder. The following describes the steps.


1. The code sample calls  **[Folder.GetStorage](folder-getstorage-method-outlook.md)** to obtain an existing **StorageItem** object that has the subject "My Private Storage" in the Inbox. If no **StorageItem** with that subject already exists, **GetStorage** creates a **StorageItem** object with that subject.
    
2. If the  **StorageItem** is newly created, the code sample creates a custom property "Order Number" for the object. Note that "Order Number" is a property of a hidden item in the Inbox.
    
3. The code sample then assigns a value to "Order Number" and saves the  **StorageItem** object.
    





```
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


## Methods



|**Name**|
|:-----|
|[Delete](storageitem-delete-method-outlook.md)|
|[Save](storageitem-save-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](storageitem-application-property-outlook.md)|
|[Attachments](storageitem-attachments-property-outlook.md)|
|[Body](storageitem-body-property-outlook.md)|
|[Class](storageitem-class-property-outlook.md)|
|[CreationTime](storageitem-creationtime-property-outlook.md)|
|[Creator](storageitem-creator-property-outlook.md)|
|[EntryID](storageitem-entryid-property-outlook.md)|
|[LastModificationTime](storageitem-lastmodificationtime-property-outlook.md)|
|[Parent](storageitem-parent-property-outlook.md)|
|[PropertyAccessor](storageitem-propertyaccessor-property-outlook.md)|
|[Session](storageitem-session-property-outlook.md)|
|[Size](storageitem-size-property-outlook.md)|
|[Subject](storageitem-subject-property-outlook.md)|
|[UserProperties](storageitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
