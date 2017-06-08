---
title: Application.DetailStylesProperties Method (Project)
keywords: vbapj.chm952
f1_keywords:
- vbapj.chm952
ms.prod: project-server
api_name:
- Project.Application.DetailStylesProperties
ms.assetid: f066f826-eef2-7f97-dafa-998f7bd70f42
ms.date: 06/08/2017
---


# Application.DetailStylesProperties Method (Project)

Sets the format of details in a usage view.


## Syntax

 _expression_. **DetailStylesProperties**( ** _AlignCellData_**, ** _RepeatRowLabel_**, ** _ShortLabels_**, ** _DisplayDetailsColumn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlignCellData_|Optional|**Long**|Specifies the alignment of data in cells. Can be one of the following  **PjAlignment** constants: **pjCenter**, **pjLeft**, or **pjRight**. The default value is **pjRight**.|
| _RepeatRowLabel_|Optional|**Boolean**|**True** if details headers are repeated on all assignment rows. The default value is **True**.|
| _ShortLabels_|Optional|**Boolean**|**True** if Project displays short details header names. The default value is **True**.|
| _DisplayDetailsColumn_|Optional|**Long**|Specifies whether a details column displays. Can be one of the following  **PjYesNoAutomatic** constants: **pjAuto**, **pjNo**, or **pjYes**. The default value is **pjYes**.|

### Return Value

 **Boolean**


## Remarks

Using the  **DetailStylesProperties** method without specifying any arguments displays the **Detail Styles** dialog box with the **Usage Properties** tab selected.


## Example

The following example hides the details column displays.


```vb
Sub DetailStyles_Remove() 
 
    ' Activate the Usage view 
    ViewApply Name:="Task Usage" 
    DetailStylesRemove Item:=pjWork 
End Sub
```


