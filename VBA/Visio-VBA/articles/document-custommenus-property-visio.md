---
title: Document.CustomMenus Property (Visio)
keywords: vis_sdr.chm10513345
f1_keywords:
- vis_sdr.chm10513345
ms.prod: visio
api_name:
- Visio.Document.CustomMenus
ms.assetid: f7d3ec25-a62e-ffe3-affe-c80ed46f58a6
ms.date: 06/08/2017
---


# Document.CustomMenus Property (Visio)

Gets a  **UIObject** object that represents the current custom menus and accelerators of a **Document** object. Read-only.


## Syntax

 _expression_ . **CustomMenus**

 _expression_ A variable that represents a **Document** object.


### Return Value

UIObject


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If Microsoft Visio menus and accelerators have not been customized, either programmatically, by a Visio solution, or in the user interface, the  **CustomMenus** property returns **Nothing** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the currently active user interface (UI) for your document without replacing the application-level custom UI. You must write additional code to add your custom UI items.


```vb
 
Sub CustomMenus_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 
 'Check whether there are custom menus bound to the document. 
 If ThisDocument.CustomMenus Is Nothing Then 
 
 'If not, check whether there are custom menus bound to the application. 
 If Visio.Application.CustomMenus Is Nothing Then 
 
 'If not, use the Visio built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 MsgBox "Using Built-In Menus", 0 
 
 Else 
 
 'If there are existing Visio application-level custom menus, use them. 
 Set vsoUIObject = Visio.Application.CustomMenus 
 MsgBox "Using Custom Menus", 0 
 
 End If 
 
 Else 
 
 'Use the existing custom menus. 
 Set vsoUIObject = ThisDocument.CustomMenus 
 MsgBox "Using Custom Menus", 0 
 
 End If 
 
End Sub
```


