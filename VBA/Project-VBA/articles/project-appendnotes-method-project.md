---
title: Project.AppendNotes Method (Project)
ms.prod: project-server
api_name:
- Project.Project.AppendNotes
ms.assetid: 65214275-905f-abcf-f75e-7589c4737e62
ms.date: 06/08/2017
---


# Project.AppendNotes Method (Project)

Appends text to the Notes field.


## Syntax

 _expression_. **AppendNotes**( ** _Value_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.


