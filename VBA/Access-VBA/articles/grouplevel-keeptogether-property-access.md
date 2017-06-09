---
title: GroupLevel.KeepTogether Property (Access)
keywords: vbaac10.chm12245
f1_keywords:
- vbaac10.chm12245
ms.prod: access
api_name:
- Access.GroupLevel.KeepTogether
ms.assetid: 65bc99df-7b0f-ec66-5add-0943ef0cd1f3
ms.date: 06/08/2017
---


# GroupLevel.KeepTogether Property (Access)

You can use the  **KeepTogether** property for a group in a report to keep parts of a group ? including the group header, detail section, and group footer ? together on the same page. For example, you might want a group header to always be printed on the same page with the first detail section. Read/write **Byte**.


## Syntax

 _expression_. **KeepTogether**

 _expression_ A variable that represents a **GroupLevel** object.


## Remarks

The  **KeepTogether** property for a group uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|No|0|(Default) Prints the group without keeping the group header, detail section, and group footer on the same page.|
|Whole Group|1|Prints the group header, detail section, and group footer on the same page.|
|With First Detail|2|Prints the group header on a page only if it can also print the first detail record.|
In Visual Basic, you set the  **KeepTogether** property for a group in report Design view or the **Open** event procedure of a report by using the **GroupLevel** property.

To set the  **KeepTogether** property for a group to a value other than No, you must set the **GroupHeader** or **GroupFooter** property or both to Yes for the selected field or expression.

A group includes the group header, detail section, and group footer. If you set the  **KeepTogether** property for a group to Whole Group and the group is too large to fit on one page, Microsoft Access will ignore the setting for that group. Similarly, if you set this property to With First Detail and either the group header or detail record is too large to fit on one page, the setting will be ignored.

If the  **KeepTogether** property for a section is set to No and the **KeepTogether** property for a group is set to Whole Group or With First Detail, the **KeepTogether** property setting for the section is ignored.


## See also


#### Concepts


[GroupLevel Object](grouplevel-object-access.md)

