---
title: Report.FontUnderline Property (Access)
keywords: vbaac10.chm13752
f1_keywords:
- vbaac10.chm13752
ms.prod: access
api_name:
- Access.Report.FontUnderline
ms.assetid: 37f62220-069d-939d-7ad0-e9f25ae6bf36
ms.date: 06/08/2017
---


# Report.FontUnderline Property (Access)

You can use the  **FontUnderline** property to specify whether text is underlined in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

 Read/write **Integer**.


## Syntax

 _expression_. **FontUnderline**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **FontUnderline** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|The text is underlined.|
|**False**|(Default) The text isn't underlined.|
For reports, you can use this property only in an event procedure or in a macro specified by the  **OnPrint** event property setting.

You can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.


## See also


#### Concepts


[Report Object](report-object-access.md)

