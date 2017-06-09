---
title: Folder.WebViewURL Property (Outlook)
keywords: vbaol11.chm2001
f1_keywords:
- vbaol11.chm2001
ms.prod: outlook
api_name:
- Outlook.Folder.WebViewURL
ms.assetid: 6a7630c2-5c16-f63f-a435-987c7c1251b8
ms.date: 06/08/2017
---


# Folder.WebViewURL Property (Outlook)

Returns or sets a  **String** indicating the URL of the Web page that is assigned to a folder. Read/write.


## Syntax

 _expression_ . **WebViewURL**

 _expression_ A variable that represents a **Folder** object.


## Example

The following Visual Basic for Applications (VBA) example creates a subfolder under the Inbox folder and assigns a home page to it.


```vb
Sub SetupFolderHomePage() 
 
 Dim nsp As Outlook.NameSpace 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim mpfNew As Outlook.Folder 
 
 
 
 Set nsp = Application.GetNamespace("MAPI") 
 
 Set mpfInbox = nsp.GetDefaultFolder(olFolderInbox) 
 
 Set mpfNew = mpfInbox.Folders.Add("MyFolderHomePage") 
 
 mpfNew.WebViewURL = "http://www.microsoft.com" 
 
 mpfNew.WebViewOn = True 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

