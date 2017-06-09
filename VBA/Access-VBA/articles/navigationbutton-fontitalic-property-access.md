---
title: NavigationButton.FontItalic Property (Access)
keywords: vbaac10.chm10475
f1_keywords:
- vbaac10.chm10475
ms.prod: access
api_name:
- Access.NavigationButton.FontItalic
ms.assetid: e4975f8e-be04-8a18-df90-9974159820fb
ms.date: 06/08/2017
---


# NavigationButton.FontItalic Property (Access)

You can use the  **FontItalic** property to specify whether text is italic in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

 Read/write **Boolean**.


## Syntax

 _expression_. **FontItalic**

 _expression_ A variable that represents a **NavigationButton** object.


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


[NavigationButton Object](navigationbutton-object-access.md)

