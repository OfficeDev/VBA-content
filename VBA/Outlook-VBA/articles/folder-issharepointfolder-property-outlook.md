---
title: Folder.IsSharePointFolder Property (Outlook)
keywords: vbaol11.chm2014
f1_keywords:
- vbaol11.chm2014
ms.prod: outlook
api_name:
- Outlook.Folder.IsSharePointFolder
ms.assetid: fc2e2645-d6e0-0bc0-29a2-8cc17f456225
ms.date: 06/08/2017
---


# Folder.IsSharePointFolder Property (Outlook)

Returns a  **Boolean** that determines if the folder is a Microsoft SharePoint Foundation folder. Read-only.


## Syntax

 _expression_ . **IsSharePointFolder**

 _expression_ A variable that represents a **Folder** object.


## Remarks

A SharePoint Foundation folder is a custom folder in Outlook that contains a live copy of the contact list or event list that lives on a SharePoint Foundation Web site. The contact list maps to a Contacts folder in Outlook and the event list maps to a Calendar folder. 

SharePoint Foundation folders are automatically created under the  **SharePoint Folders** node in the Navigation Pane when a contact list or an event list is exported from the SharePoint Foundation Web site.

Though SharePoint Foundation folders work the same way as other folders, there are a few exceptions. SharePoint Foundation folders are read-only and any attempt to edit folder properties or add, edit, or remove existing items will fail. 

A folder in the user's Microsoft Exchange server folder will never be a SharePoint Foundation folder, and no folder in the user's default Personal Folders file (.pst) will ever be a SharePoint Foundation folder. Typically the SharePoint Foundation folders will be under the node  **SharePoint Folders** in the Navigation Pane.


## Example

The following Microsoft Visual Basic for Applications (VBA) example changes the Subject line of the appointment item displayed in the active inspector and saves the item. If the item is contained in a SharePoint Foundation folder, it displays a message to the user that the item cannot be modified. To run this example, make sure that an appointment item is displayed in the active inspector window. This example will modify the subject of the appointment item.


```vb
Sub ChangeItem() 
 
'Checks if the item is contained in a SharePoint folder. If it is not, it changes the Subject line, and then saves the item. 
 
 Dim myItem As Outlook.AppointmentItem 
 
 Dim fldFolder As Outlook.Folder 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 Set fldFolder = myItem.Parent 
 
 If fldFolder.IsSharePointFolder = True Then 
 
 MsgBox _ 
 
 "The item is contained in a SharePoint Foundation folder and cannot be modified." 
 
 Else 
 
 myItem.Subject = myItem.Subject + " Changed by VBA" 
 
 myItem.Save 
 
 MsgBox "The item has been changed." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

