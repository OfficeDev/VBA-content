---
title: DoCmd.ShowToolbar Method (Access)
keywords: vbaac10.chm4185
f1_keywords:
- vbaac10.chm4185
ms.prod: access
api_name:
- Access.DoCmd.ShowToolbar
ms.assetid: 63663cc5-a591-c847-25c8-25777cf7806a
ms.date: 06/08/2017
---


# DoCmd.ShowToolbar Method (Access)

The  **ShowToolbar** method carries out the ShowToolbar action in Visual Basic.


## Syntax

 _expression_. **ShowToolbar**( ** _ToolbarName_**, ** _Show_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ToolbarName_|Required|**Variant**|A string expression that's the valid name of a custom toolbar you've created. If you run Visual Basic code containing the  **ShowToolbar** method in a library database, Microsoft Access looks for the toolbar with this name first in the library database, then in the current database.|
| _Show_|Optional|**AcShowToolbar**| An[AcShowToolbar](acshowtoolbar-enumeration-access.md) constant that specifies whether to display or hide the toolbar and in which views to display or hide it. The default value is **acToolbarYes**.|

## Remarks

You can use the  **ShowToolbar** method to display or hide a custom toolbar.

If you want to show a particular toolbar on just one form or report, you can set the  **OnActivate** property of the form or report to the name of a macro that contains a ShowToolbar action to show the toolbar. Then set the **OnDeactivate** property of the form or report to the name of a macro that contains a ShowToolbar action to hide the toolbar.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

