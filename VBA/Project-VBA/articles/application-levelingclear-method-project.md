---
title: Application.LevelingClear Method (Project)
keywords: vbapj.chm612
f1_keywords:
- vbapj.chm612
ms.prod: project-server
api_name:
- Project.Application.LevelingClear
ms.assetid: fdd537eb-f9c2-c8d9-ec26-0f4af9a63c33
ms.date: 06/08/2017
---


# Application.LevelingClear Method (Project)

Removes the effects of leveling.


## Syntax

 _expression_. **LevelingClear**( ** _All_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _All_|Optional|**Boolean**|**True** if delays are removed from all tasks. **False** if delays are removed from selected tasks only.|

### Return Value

 **Boolean**


## Remarks

Using the  **LevelingClear** method without specifying any arguments displays the **Clear Leveling** dialog box.

The  **LevelingClear** method has no effect if a task has a priority of 1000 (do not level).


