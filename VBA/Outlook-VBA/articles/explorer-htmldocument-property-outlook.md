---
title: Explorer.HTMLDocument Property (Outlook)
keywords: vbaol11.chm2778
f1_keywords:
- vbaol11.chm2778
ms.prod: outlook
api_name:
- Outlook.Explorer.HTMLDocument
ms.assetid: dd9ff575-37f5-1b64-5ebf-f17998586d28
ms.date: 06/08/2017
---


# Explorer.HTMLDocument Property (Outlook)

Returns an  **HTMLDocument** object that specifies the HTML object model associated with the HTML document in the current view (assuming one exists). Read-only.


## Syntax

 _expression_ . **HTMLDocument**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

In order to use this property, a folder must be using a folder home page, or you can set the  **[WebViewURL](folder-webviewurl-property-outlook.md)** property of the **[Folder](folder-object-outlook.md)** object to a Web page.


## Example

The following Microsoft Visual Basic for Applications (VBA) example accesses the Microsoft Outlook View Control.


```vb
 Sub GetHTML() 
 
'Returns the Outlook View Control 
 
 
 
 Dim objVC As OLXLib.ViewCtl 
 
 Dim objExp As Outlook.Explorer 
 
 Dim HTMLDoc As MSHTML.HTMLDocument 
 
 
 
 'Reference the current folder 
 
 Set objExp = Application.ActiveExplorer 
 
 
 
 'Reference the HTML file that is the home page 
 
 Set HTMLDoc = objExp.HTMLDocument 
 
 
 
 'Reference an Outlook View Control that is on the HTML page 
 
 Set objVC = HTMLDoc.all.tags("object").Item(0).Object 
 
 
 
 'Have the control display an address book window 
 
 objVC.AddressBook 
 
 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

