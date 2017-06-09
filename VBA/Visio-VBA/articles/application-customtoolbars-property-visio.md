---
title: Application.CustomToolbars Property (Visio)
keywords: vis_sdr.chm10013355
f1_keywords:
- vis_sdr.chm10013355
ms.prod: visio
api_name:
- Visio.Application.CustomToolbars
ms.assetid: 1c945955-af48-5dd1-f186-d7d0cf02e6d2
ms.date: 06/08/2017
---


# Application.CustomToolbars Property (Visio)

Gets a  **UIObject** object that represents the current custom toolbars and status bars of an **Application** object. Read-only.


## Syntax

 _expression_ . **CustomToolbars**

 _expression_ A variable that represents an **Application** object.


### Return Value

UIObject


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If Microsoft Visio toolbars and status bars have not been customized, either programmatically, by a Visio solution, or in the user interface, the  **CustomToolbars** property returns **Nothing** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the currently active user interface (UI) for your document without replacing the application-level custom UI. You must write additional code to add your custom UI items.


```vb
Sub CustomToolbars_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 
 'Check whether there are custom toolbars bound to the document. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'If not, check whether there are custom toolbars bound to the application. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'If not, use the Visio built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 MsgBox "Using Built-In Toolbars", 0 
 
 Else 
 
 'If there are existing Visio application-level custom toolbars, use them. 
 Set vsoUIObject = Visio.Application.CustomToolbars 
 MsgBox "Using Custom Toolbars", 0 
 
 End If 
 
 Else 
 
 'Use the existing custom toolbars. 
 Set vsoUIObject = ThisDocument.CustomToolbars 
 MsgBox "Using Custom Toolbars", 0 
 
 End If 
 
End Sub
```


