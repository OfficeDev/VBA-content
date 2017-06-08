---
title: InvisibleApp.CustomToolbars Property (Visio)
keywords: vis_sdr.chm17513355
f1_keywords:
- vis_sdr.chm17513355
ms.prod: visio
api_name:
- Visio.InvisibleApp.CustomToolbars
ms.assetid: 0f237eec-3d1b-329c-8f39-3dcc49f98570
ms.date: 06/08/2017
---


# InvisibleApp.CustomToolbars Property (Visio)

Gets a  **UIObject** object that represents the current custom toolbars and status bars of an **InvisibleApp** object. Read-only.


## Syntax

 _expression_ . **CustomToolbars**

 _expression_ A variable that represents an **InvisibleApp** object.


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


