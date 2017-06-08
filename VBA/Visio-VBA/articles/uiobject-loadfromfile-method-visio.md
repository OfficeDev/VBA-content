---
title: UIObject.LoadFromFile Method (Visio)
keywords: vis_sdr.chm14916390
f1_keywords:
- vis_sdr.chm14916390
ms.prod: visio
api_name:
- Visio.UIObject.LoadFromFile
ms.assetid: 6a4ef6d5-9a3a-771b-be87-bc5f21bce4e7
ms.date: 06/08/2017
---


# UIObject.LoadFromFile Method (Visio)

Loads a Microsoft Visio application  **UIObject** object from a file.


## Syntax

 _expression_ . **LoadFromFile**( **_FileName_** )

 _expression_ A variable that represents a **UIObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to load.|

### Return Value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You must use the  **SaveToFile** method to save a **UIObject** object in a file that the **LoadFromFile** method can load.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to save and then load a custom user interface file (.vsu). It does not manipulate any menus or menu items.

Before running this code, replace  _path\filename_ with the full path to and name of a valid .vsu file on your computer.




```vb
 
Public Sub LoadFromFile_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim strPath As String 
 
 'Get Menus object from Visio. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Save Menus object to a file. 
 strPath = "path\filename.vsu " 
 vsoUIObject.SaveToFile (strPath) 
 MsgBox ("Menus successfully saved to " &; strPath) 
 
 'Load menus from the file. 
 vsoUIObject.LoadFromFile (strPath) 
 Visio.Application.SetCustomMenus vsoUIObject 
 MsgBox ("Menus successfully loaded from " &; strPath) 
 
End Sub
```


