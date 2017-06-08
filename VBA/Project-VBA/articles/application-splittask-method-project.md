---
title: Application.SplitTask Method (Project)
keywords: vbapj.chm1011
f1_keywords:
- vbapj.chm1011
ms.prod: project-server
api_name:
- Project.Application.SplitTask
ms.assetid: 490dcca9-66c5-9284-44ff-a92aa30fadf4
ms.date: 06/08/2017
---


# Application.SplitTask Method (Project)

Enters the interactive task split mode, enabling the user to manually create task splits.


## Syntax

 _expression_. **SplitTask**( ** _Lock_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Lock_|Optional|**Boolean**|**True** if the task split pointer stays active after a split is made, enabling more task splits to be made. **False** if the pointer returns to normal after making a split. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SplitTask** method requires user interaction before additional code can be executed. The **SplitTask** method is only available in Gantt views; it corresponds to the **Split Task** icon on the **Task** tab of the Ribbon.


