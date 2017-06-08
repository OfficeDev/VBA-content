---
title: NavigationControl.AutoTab Property (Access)
keywords: vbaac10.chm11063
f1_keywords:
- vbaac10.chm11063
ms.prod: access
api_name:
- Access.NavigationControl.AutoTab
ms.assetid: 3d484269-c00b-3f5e-8492-6e0ca60460b8
ms.date: 06/08/2017
---


# NavigationControl.AutoTab Property (Access)

You can use the  **AutoTab** property to specify whether an automatic tab occurs when the last character permitted by a text box control's input mask is entered. An automatic tab moves the focus to the next control in the form's tab order. Read/write **Boolean**.


## Syntax

 _expression_. **AutoTab**

 _expression_ A variable that represents a **NavigationControl** object.


## Remarks

The  **AutoTab** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|Generates a tab when the last allowable character in a text box is entered.|
|No|**False**|(Default) Doesn't generate a tab when the last allowable character in a text box is entered.|
You can also set the default for this property by setting a control's  **DefaultControl** property in Visual Basic.

You can also create an input mask for a text box control bound to a field by setting the  **InputMask** property for the field in the form's underlying table or query. If the field is dragged to a form from the field list, the field's input mask is inherited by the text box control.

You could use the  **AutoTab** property if you have a text box on a form for which you usually enter the maximum number of characters for each record. After you have entered the maximum number of characters, focus automatically moves to the next control in the tab order. For example, you could use this property for a CategoryType field that must always be five characters long.


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

