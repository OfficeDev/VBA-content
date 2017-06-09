---
title: Task.AppendNotes Method (Project)
ms.prod: project-server
api_name:
- Project.Task.AppendNotes
ms.assetid: ab0177cb-c7cd-444f-0d19-9b798eba8b4a
ms.date: 06/08/2017
---


# Task.AppendNotes Method (Project)

Appends text to the Notes field.


## Syntax

 _expression_. **AppendNotes**( ** _Value_** )

 _expression_ A variable that represents a **Task** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.


