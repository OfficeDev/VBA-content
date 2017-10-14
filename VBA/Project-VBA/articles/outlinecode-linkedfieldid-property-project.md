---
title: OutlineCode.LinkedFieldID Property (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode.LinkedFieldID
ms.assetid: 310202bc-6db7-11b8-d380-af26ef12ad11
ms.date: 06/08/2017
---


# OutlineCode.LinkedFieldID Property (Project)

Gets or sets the outline code field ID for a linked lookup table. Obsolete in Project. Read/write  **Long**.


## Syntax

 _expression_. **LinkedFieldID**

 _expression_ A variable that represents an **OutlineCode** object.


## Remarks

A local outline code can import a lookup table from another outline code, but cannot link to it or share it with another outline code or an enterprise text custom field. the  **LinkedFieldID** property always returns -1.


