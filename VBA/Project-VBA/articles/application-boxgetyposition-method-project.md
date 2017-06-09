---
title: Application.BoxGetYPosition Method (Project)
keywords: vbapj.chm131247
f1_keywords:
- vbapj.chm131247
ms.prod: project-server
api_name:
- Project.Application.BoxGetYPosition
ms.assetid: 8284181f-b677-8cc4-8311-23d50987239c
ms.date: 06/08/2017
---


# Application.BoxGetYPosition Method (Project)

Returns the vertical position of the upper-left corner of a box in the active Network Diagram. At a zoom setting of 100%, the number of nonscaleable units returned by  **BoxGetYPosition** is equivalent to an equal number of pixels.


## Syntax

 _expression_. **BoxGetYPosition**( ** _TaskID_**, ** _ProjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Required|**Long**|The identification number of the task.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the identification number specified by  **TaskID**. The default value is the name of the active project.|

### Return Value

 **Long**


