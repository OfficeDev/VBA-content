---
title: TextBox.FontItalic Property (Access)
keywords: vbaac10.chm11087
f1_keywords:
- vbaac10.chm11087
ms.prod: access
api_name:
- Access.TextBox.FontItalic
ms.assetid: f982c1ce-ad47-a05e-6b12-1eb51dbc0eb7
ms.date: 06/08/2017
---


# TextBox.FontItalic Property (Access)

You can use the  **FontItalic** property to specify whether text is italic in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

 Read/write **Boolean**.


## Syntax

 _expression_. **FontItalic**

 _expression_ A variable that represents a **TextBox** object.


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


[TextBox Object](textbox-object-access.md)

