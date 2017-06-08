---
title: Assignment.AppendNotes Method (Project)
ms.prod: project-server
api_name:
- Project.Assignment.AppendNotes
ms.assetid: 78ccad76-ac3f-c11e-9d88-2ed133358671
ms.date: 06/08/2017
---


# Assignment.AppendNotes Method (Project)

Appends text to the Notes field.


## Syntax

 _expression_. **AppendNotes**( ** _Value_** )

 _expression_ A variable that represents an **Assignment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.


