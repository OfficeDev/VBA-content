---
title: ComboBox.FontUnderline Property (Access)
keywords: vbaac10.chm11418
f1_keywords:
- vbaac10.chm11418
ms.prod: access
api_name:
- Access.ComboBox.FontUnderline
ms.assetid: 54ee770c-4e75-fbc7-0453-99fc2c2456c1
ms.date: 06/08/2017
---


# ComboBox.FontUnderline Property (Access)

You can use the  **FontUnderline** property to specify whether text is underlined in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

 Read/write **Boolean**.


## Syntax

 _expression_. **FontUnderline**

 _expression_ A variable that represents a **ComboBox** object.


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


[ComboBox Object](combobox-object-access.md)

