---
title: Folder.WebViewOn Property (Outlook)
keywords: vbaol11.chm2000
f1_keywords:
- vbaol11.chm2000
ms.prod: outlook
api_name:
- Outlook.Folder.WebViewOn
ms.assetid: 9b483d0e-dea0-9b3e-8ce9-fc136857a428
ms.date: 06/08/2017
---


# Folder.WebViewOn Property (Outlook)

Returns or sets a  **Boolean** indicating the Web view state for a folder. Read/write.


## Syntax

 _expression_ . **WebViewOn**

 _expression_ A variable that represents a **Folder** object.


## Remarks

Returns  **True** to display the Web page specified by the **[WebViewURL](folder-webviewurl-property-outlook.md)** property of the **[Folder](folder-object-outlook.md)** object.

Microsoft Outlook uses the rendering engine of the version Windows Internet Explorer installed on the client computer to display the web page. If Internet Explorer is not installed on the client computer, Outlook will not display the web page.

This property is always  **False** if the value of the **WebViewURL** property is empty.

Also, setting the  **WebViewOn** property to **True** before setting the **WebViewURL** property will not display the home page specified in the **WebViewURL** property.


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

