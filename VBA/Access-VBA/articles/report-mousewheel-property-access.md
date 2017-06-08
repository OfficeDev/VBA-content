---
title: Report.MouseWheel Property (Access)
keywords: vbaac10.chm13872
f1_keywords:
- vbaac10.chm13872
ms.prod: access
api_name:
- Access.Report.MouseWheel
ms.assetid: ea9d6443-abfd-6140-e167-548f4aafd342
ms.date: 06/08/2017
---


# Report.MouseWheel Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **MouseWheel** event occurs. Read/write.


## Syntax

 _expression_. **MouseWheel**

 _expression_ A variable that represents a **Report** object.


## Remarks

Valid values for this property are "macroname", where  _macroname_ is the name of the macro; "[Event Procedure]", which indicates the event procedure associated with the **BeforeInsert** event for the specified object; or "=functionname()", where _functionname_ is the name of a user-defined function.


## See also


#### Concepts


[Report Object](report-object-access.md)

