---
title: Application.BoxGetXPosition Method (Project)
keywords: vbapj.chm131246
f1_keywords:
- vbapj.chm131246
ms.prod: project-server
api_name:
- Project.Application.BoxGetXPosition
ms.assetid: df7a41c8-01df-bd60-0ae1-0fb60cbc3347
ms.date: 06/08/2017
---


# Application.BoxGetXPosition Method (Project)

Returns the horizontal position of the upper-left corner of a box in the active Network Diagram. At a zoom setting of 100%, the number of nonscaleable units returned by  **BoxGetXPosition** is equivalent to an equal number of pixels.


## Syntax

 _expression_. **BoxGetXPosition**( ** _TaskID_**, ** _ProjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Required|**Long**|The identification number of the task.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the identification number specified by  **TaskID**. The default value is the name of the active project.|

### Return Value

 **Long**


