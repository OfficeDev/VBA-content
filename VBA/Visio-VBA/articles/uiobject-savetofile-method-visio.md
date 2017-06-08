---
title: UIObject.SaveToFile Method (Visio)
keywords: vis_sdr.chm14916510
f1_keywords:
- vis_sdr.chm14916510
ms.prod: visio
api_name:
- Visio.UIObject.SaveToFile
ms.assetid: 0e734a30-08be-e3e8-590f-88e399e699fd
ms.date: 06/08/2017
---


# UIObject.SaveToFile Method (Visio)

Saves the user interface represented by a  **UIObject** object in a file.


## Syntax

 _expression_ . **SaveToFile**( **_FileName_** )

 _expression_ A variable that represents a **UIObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file in which to save the  **UIObject** object.|

### Return Value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The file can be loaded into the application by using the  **LoadFromFile** method of a **UIObject** object.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to save a custom user interface file (.vsu). It does not manipulate any Visio menus or menu items. Before running this macro, change  _path_ to the location where you want to save the file, and change _filename_ to the name you'd like to assign the file.


```vb
 
Public Sub SaveMenusToFile_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim strPath As String 
 
 'Get the Menus object from Visio. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Save the Menus object to a file. 
 strPath = "path\filename.vsu " 
 vsoUIObject.SaveToFile (strPath) 
 MsgBox ("Menus successfully saved to " &; strPath) 
 
End Sub
```


