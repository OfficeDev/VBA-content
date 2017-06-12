---
title: Application.DetailStylesAdd Method (Project)
keywords: vbapj.chm963
f1_keywords:
- vbapj.chm963
ms.prod: project-server
api_name:
- Project.Application.DetailStylesAdd
ms.assetid: 40a1dfa4-ef57-835d-4e42-9631c906ac0b
ms.date: 06/08/2017
---


# Application.DetailStylesAdd Method (Project)

Adds another timescale data field to a usage view.


## Syntax

 _expression_. **DetailStylesAdd**( ** _Item_**, ** _Position_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Long**|The timescale data field to add. The default value is  **pjWork**.If the active view is the Resource Usage view, can be one of the **[PjTimescaledData](pjtimescaleddata-enumeration-project.md)** constants.|
| _Position_|Optional|**Integer**|The position to add the field, relative to other fields. If  **Position** is n + 2 or greater, where n is the number of fields displayed, the field is added at n + 1. The default value is n + 1.|

### Return Value

 **Boolean**


## Example

The following example makes overallocations stand out from other information in a usage view.


```vb
Sub HighlightOverallocations() 
 
 DetailStylesAdd pjOverallocation 
 DetailStylesFormat Item:=pjOverallocation, Font:="Arial", Size:=12, _ 
 Bold:=True, Color:=pjRed, CellColor:=pjBlack, Pattern:=pjSolidFill 
 
End Sub
```


