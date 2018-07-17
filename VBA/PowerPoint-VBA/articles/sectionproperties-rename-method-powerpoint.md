---
title: SectionProperties.Rename Method (PowerPoint)
keywords: vbapp10.chm725005
f1_keywords:
- vbapp10.chm725005
ms.prod: powerpoint
api_name:
- PowerPoint.SectionProperties.Rename
ms.assetid: b1e7428e-c7ee-04b8-1f09-246fe3e7fe6f
ms.date: 06/08/2017
---


# SectionProperties.Rename Method (PowerPoint)

Renames the specified section with the specified name.


## Syntax

 _expression_. **Rename**( **_sectionIndex_**, **_sectionName_** )

 _expression_ A variable that represents a **SectionProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _sectionIndex_|Required|**Integer**|The index of the section to rename.|
| _sectionName_|Required|**String**|The new name of the section.|

## Remarks

If sectionName is empty, the section is assigned the default section name.


## See also


#### Concepts


[SectionProperties Object](sectionproperties-object-powerpoint.md)

