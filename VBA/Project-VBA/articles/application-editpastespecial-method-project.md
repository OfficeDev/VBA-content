---
title: Application.EditPasteSpecial Method (Project)
keywords: vbapj.chm232
f1_keywords:
- vbapj.chm232
ms.prod: project-server
api_name:
- Project.Application.EditPasteSpecial
ms.assetid: afbe96f1-a4f6-e879-cacc-115761f5e1c4
ms.date: 06/08/2017
---


# Application.EditPasteSpecial Method (Project)

Copies or links data from the Clipboard into the active selection.


## Syntax

 _expression_. **EditPasteSpecial**( ** _Link_**, ** _Type_**, ** _DisplayAsIcon_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Link_|Optional|**Boolean**|**True** if the data is linked to its source application.|
| _Type_|Optional|**Integer**|A numeric value specifying the type of object to paste or link. The  **Type** argument can be one of the **[PjPasteSpecialType](pjpastespecialtype-enumeration-project.md)** constants.|
| _DisplayAsIcon_|Optional|**Boolean**|**True** if the object appears as an icon.|

### Return Value

 **Boolean**


## Example

The following example pastes the Clipboard content as a picture.


```vb
Sub Edit_PasteSpecial() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="&;Gantt Chart" 
 
 SelectRow Row:=2, RowRelative:=False 
 EditPasteSpecial Link:=False, Type:=pjPicture, DisplayAsIcon:=False 
 
End Sub
```


