---
title: Import Saved Items using OpenSharedItem
ms.prod: outlook
ms.assetid: e3e770c4-a4fd-6484-dbee-0d5e5141d9f9
ms.date: 06/08/2017
---


# Import Saved Items using OpenSharedItem

Microsoft Outlook provides the  **[OpenSharedItem](namespace-openshareditem-method-outlook.md)** method, for the **[NameSpace](namespace-object-outlook.md)** object, to open iCalendar appointment (.ics) files, vCard (.vcf) files, and Outlook message (.msg) files and return the Outlook item appropriate for the file. The type of object returned by this method depends on the type of shared item opened, as described in the following table.


| **Shared item type**| **Outlook item**|
|:-----|:-----|
|iCalendar appointment (.ics) file| **[AppointmentItem](appointmentitem-object-outlook.md)**|
|vCard (.vcf) file| **[ContactItem](contactitem-object-outlook.md)**|
|Outlook message (.msg) file|Type corresponds to the type of the item that was saved as the .msg file|

Once the shared item is opened, you can then import the item by using the  **Save** method of the returned object to save it to the default folder appropriate to that Outlook item.

This sample opens and imports a vCard file into the  **Contacts** default folder for the current user.

1. The sample obtains a reference to a  **NameSpace** object, then calls the **GetSharedItem** method of the **NameSpace** object to load the vCard file and return a **ContactItem** reference.
    
2. It then calls the  **Save** method of the **ContactItem** to save it to the **Contacts** default folder.
    
3. Finally, it obtains a  **[Folder](folder-object-outlook.md)** object reference to the **Contacts** default folder for the current user by using the **[GetDefaultFolder](namespace-getdefaultfolder-method-outlook.md)** method of the **NameSpace** object, and then displays the folder.
    



```vb
Public Sub OpenSharedContact() 
 
 Dim oNamespace As NameSpace 
 Dim oSharedItem As ContactItem 
 Dim oFolder As Folder 
 
 On Error GoTo ErrRoutine 
 
 ' Get a reference to a NameSpace object. 
 Set oNamespace = Application.GetNamespace("MAPI") 
 
 ' Open the vCard (.vcf) file containing the shared item. 
 Set oSharedItem = oNamespace.OpenSharedItem( _ 
 "C:/SampleContact.vcf") 
 
 ' Save the item to the Contacts default folder. 
 oSharedItem.Save 
 
 ' Get a reference to and display the Contacts default folder. 
 Set oFolder = oNamespace.GetDefaultFolder( _ 
 olFolderContacts) 
 oFolder.Display 
 
EndRoutine: 
 On Error GoTo 0 
 Set oSharedItem = Nothing 
 Set oFolder = Nothing 
 Set oNamespace = Nothing 
Exit Sub 
 
ErrRoutine: 
 Select Case Err.Number 
 Case 287 ' &;H0000011F 
 ' This error occurs if the code is run by an 
 ' untrusted application, and the user chose not to 
 ' allow access. 
 MsgBox "Access to Outlook was denied by the user.", _ 
 vbOKOnly, _ 
 Err.Number &; " - " &; Err.Source 
 Case -2147024894 ' &;H80070002 
 ' Occurs if the specified file or URL could not 
 ' be found, or the file or URL cannot be 
 ' processed by the OpenSharedItem method. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number &; " - " &; Err.Source 
 Case -2147352567 ' &;H80020009 
 ' Occurs if the specified file or URL is not valid, 
 ' or you attempt to use the Move method on 
 ' an Outlook item that represents a shared item. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number &; " - " &; Err.Source 
 Case Else 
 ' Any other error that may occur. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number &; " - " &; Err.Source 
 End Select 
 
 GoTo EndRoutine 
End Sub
```


