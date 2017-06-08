---
title: Folder.GetCustomIcon Method (Outlook)
keywords: vbaol11.chm3316
f1_keywords:
- vbaol11.chm3316
ms.prod: outlook
api_name:
- Outlook.Folder.GetCustomIcon
ms.assetid: 49a3da64-2b2f-76db-0053-88e35141cca0
ms.date: 06/08/2017
---


# Folder.GetCustomIcon Method (Outlook)

Returns an  **[IPictureDisp](http://msdn.microsoft.com/en-us/library/ms680762%28VS.85%29.aspx)** object that represents the custom icon for the folder.


## Syntax

 _expression_ . **GetCustomIcon**

 _expression_ A variable that represents a **[Folder](folder-object-outlook.md)** object.


### Return Value

An  **IPictureDisp** object that represents a custom icon for the folder.


## Remarks

The returned  **IPictureDisp** object has its **Type** property equal to **PICTYPE_ICON** or **PICTYPE_BITMAP** .

 **GetCustomIcon** returns **Null** ( **Nothing** in Visual Basic) if the folder does not have a custom folder icon, or if the folder belongs to one of the following groups of folders:


- Default folders (as listed by the  **[OlDefaultFolders](oldefaultfolders-enumeration-outlook.md)** enumeration)
    
- Special folders (as listed by the  **[OlSpecialFolders](olspecialfolders-enumeration-outlook.md)** enumeration)
    
- Exchange public folders
    
-  Root folder of any Exchange mailbox
    
- Hidden folders
    
You can only call  **GetCustomIcon** from code that runs in-process as Outlook. An **IPictureDisp** object cannot be marshaled across process boundaries. If you attempt to call **GetCustomIcon** from out-of-process code, an exception occurs. For more information, see[An automation server cannot pass a pointer to the picture object's IPictureDisp implementation across process boundaries](http://support.microsoft.com/kb/150034).


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

