---
title: Chart.Creator Property (Project)
keywords: vbapj.chm131613
f1_keywords:
- vbapj.chm131613
ms.prod: project-server
ms.assetid: d2ef5502-f55f-73ff-3df1-04aa22cbc9c0
ms.date: 06/08/2017
---


# Chart.Creator Property (Project)
Gets a 32-bit integer that indicates the application in which the chart was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_ A variable that represents a **Chart** object.


## Remarks

If the chart was created in Microsoft Project, the  **Creator** property returns the decimal number 1347571530, which is hexadecimal 0x50524F4A, which is equivalent to the string **PROJ**. For example, run the following command in the  **Immediate** window of the VBE, with the name of the active report.


```
Print ActiveProject.Reports("Simple scalar chart").Shapes(1).Chart.Creator
```


## Property value

 **INT32**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
