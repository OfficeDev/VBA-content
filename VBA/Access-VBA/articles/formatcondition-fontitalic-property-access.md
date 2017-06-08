---
title: FormatCondition.FontItalic Property (Access)
keywords: vbaac10.chm10055
f1_keywords:
- vbaac10.chm10055
ms.prod: access
api_name:
- Access.FormatCondition.FontItalic
ms.assetid: eb7dceb1-ecfb-59b4-f810-28c2becd8a95
ms.date: 06/08/2017
---


# FormatCondition.FontItalic Property (Access)

You can use the  **FontItalic** property to specify whether text is italic in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

 Read/write **Boolean**.


## Syntax

 _expression_. **FontItalic**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

The  **FontItalic** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|The text is italic.|
|**False**|(Default) The text isn't italic.|
For reports, you can use this property only in an event procedure or in a macro specified by the  **OnPrint** event property setting.

You can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-access.md)

