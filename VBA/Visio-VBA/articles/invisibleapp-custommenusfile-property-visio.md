---
title: InvisibleApp.CustomMenusFile Property (Visio)
keywords: vis_sdr.chm17513350
f1_keywords:
- vis_sdr.chm17513350
ms.prod: visio
api_name:
- Visio.InvisibleApp.CustomMenusFile
ms.assetid: 189faa67-41bb-2b87-9761-365c0c0433ba
ms.date: 06/08/2017
---


# InvisibleApp.CustomMenusFile Property (Visio)

Gets or sets the name of the file that defines custom menus and accelerators for an  **InvisibleApp** object. Read/write.


## Syntax

 _expression_ . **CustomMenusFile**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the object is not using custom menus, the  **CustomMenusFile** property returns **Nothing** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the currently active UI for your document without replacing the application-level custom UI. It also saves any existing custom menus to a file and specifies that the current document use those menus. You must write additional code to add your custom UI items.




 **Note**  This macro uses the VBA keyword  **Kill** to delete a file on disk. Use this keyword carefully, because you cannot undo a **Kill** command once it has been run, and you will not get a prior warning message.




```vb
 
Sub CustomMenusFile_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim strPath As String 
 
 'Check whether there are custom menus bound to the document. 
 If ThisDocument.CustomMenus Is Nothing Then 
 
 'If not, check whether there are custom menus bound to the application. 
 If Visio.Application.CustomMenus Is Nothing Then 
 
 'If not, use the Visio built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 MsgBox "Using Built-In Menus", 0 
 
 Else 
 
 'If there are existing Visio custom menus, use them. 
 Set vsoUIObject = Visio.Application.CustomMenus 
 
 'Save these custom menus to a file. 
 strPath = Visio.Application.Path &; "\CustomUI.vsu" 
 vsoUIObject.SaveToFile (strPath) 
 
 'Set the document to use the existing custom UI. 
 ThisDocument.CustomMenusFile = strPath 
 
 'Get this document's UIObject object. 
 Set vsoUIObject = ThisDocument.CustomMenus 
 
 'Delete the newly created temp file. 
 Kill Visio.Application.Path &; "\CustomUI.vsu" 
 ThisDocument.ClearCustomMenus 
 MsgBox "Using Custom Menus", 0 
 
 End If 
 
 Else 
 
 'Use the existing custom menus. 
 Set vsoUIObject = ThisDocument.CustomMenus 
 
 End If 
 
End Sub
```


