---
title: ContactItem.AddPicture Method (Outlook)
keywords: vbaol11.chm1090
f1_keywords:
- vbaol11.chm1090
ms.prod: outlook
api_name:
- Outlook.ContactItem.AddPicture
ms.assetid: aa02c3b2-bb4c-fde9-3dbf-f871cbc200b1
ms.date: 06/08/2017
---


# ContactItem.AddPicture Method (Outlook)

Adds a picture to a contact item.


## Syntax

 _expression_ . **AddPicture**( **_Path_** )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|A string containing the complete path and filename of the picture to be added to the contact item.|

## Remarks

 If the contact item already has a picture attached to it, this method will overwrite the existing picture.

The picture can be an icon, GIF, JPEG, BMP, TIFF, WMF, EMF, or PNG file. Microsoft Outlook will automatically perform the necessary resizing of the picture.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user to specify the name of a contact and the file name containing a picture of the contact, and then adds the picture to the contact item. If a picture already exists for the contact item, the example prompts the user to specify if the existing picture should be overwritten by the new file.


```vb
Sub AddPictureToAContact() 
 
 Dim myNms As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myContactItem As Outlook.ContactItem 
 
 Dim strName As String 
 
 Dim strPath As String 
 
 Dim strPrompt As String 
 
 
 
 Set myNms = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNms.GetDefaultFolder(olFolderContacts) 
 
 strName = InputBox("Type the name of the contact: ") 
 
 Set myContactItem = myFolder.Items(strName) 
 
 If myContactItem.HasPicture = True Then 
 
 strPrompt = MsgBox("The contact already has a picture associated with it. Do you want to overwrite the existing picture?", vbYesNo) 
 
 If strPrompt = vbNo Then 
 
 Exit Sub 
 
 End If 
 
 End If 
 
 strPath = InputBox("Type the file name for the contact: ") 
 
 myContactItem.AddPicture (strPath) 
 
 myContactItem.Save 
 
 myContactItem.Display 
 
 End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

