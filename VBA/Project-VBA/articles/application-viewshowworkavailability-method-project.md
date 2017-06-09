---
title: Application.ViewShowWorkAvailability Method (Project)
keywords: vbapj.chm930
f1_keywords:
- vbapj.chm930
ms.prod: project-server
api_name:
- Project.Application.ViewShowWorkAvailability
ms.assetid: 909fbc1a-fe49-8121-c103-e287d10a49fa
ms.date: 06/08/2017
---


# Application.ViewShowWorkAvailability Method (Project)

Displays work availability information in the active Resource Graph view.


## Syntax

 _expression_. **ViewShowWorkAvailability**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **ViewShowWorkAvailability** method has no effect unless the active window contains the Resource Graph view. The **ViewShowWorkAvailability** method is not available for material resources and returns a trappable error (error code 1100) when applied to material resources.


