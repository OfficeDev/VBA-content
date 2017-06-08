---
title: Resource.AppendNotes Method (Project)
ms.prod: project-server
api_name:
- Project.Resource.AppendNotes
ms.assetid: b11bc28f-147f-0591-056b-87e9f6c2db71
ms.date: 06/08/2017
---


# Resource.AppendNotes Method (Project)

Appends text to the Notes field.


## Syntax

 _expression_. **AppendNotes**( ** _Value_** )

 _expression_ A variable that represents a **Resource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.


