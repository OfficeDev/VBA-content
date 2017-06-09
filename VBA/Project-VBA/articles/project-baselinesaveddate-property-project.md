---
title: Project.BaselineSavedDate Property (Project)
ms.prod: project-server
api_name:
- Project.Project.BaselineSavedDate
ms.assetid: 780c5190-68bb-1c10-0dbb-612e5606184e
ms.date: 06/08/2017
---


# Project.BaselineSavedDate Property (Project)

Gets date the specified baseline was last saved. Read-only  **Variant**.


## Syntax

 _expression_. **BaselineSavedDate**( ** _Baseline_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Baseline_|Required|**Long**|Can be one of the  **[PjBaselines](pjbaselines-enumeration-project.md)** constants.|

## Remarks

If the specified baseline has not been saved,  **BaselineSavedDate** returns "NA".


