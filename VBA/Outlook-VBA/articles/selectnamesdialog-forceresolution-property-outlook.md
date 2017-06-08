---
title: SelectNamesDialog.ForceResolution Property (Outlook)
keywords: vbaol11.chm832
f1_keywords:
- vbaol11.chm832
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.ForceResolution
ms.assetid: f859e464-8d06-f44c-e388-f6b6427bec1a
ms.date: 06/08/2017
---


# SelectNamesDialog.ForceResolution Property (Outlook)

Returns or sets a  **Boolean** that determines if Outlook must resolve all recipients in the object specified by **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)** before the user can click **OK** to accept the typed or selected recipients in the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **ForceResolution**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

The default value is  **True** . If a recipient cannot be resolved, Outlook will prompt the user to resolve the ambiguous names. The user must have all recipients in the recipient edit box resolved before being able to click **OK**.

 **ForceResolution** is ignored if the user clicks **Cancel** or the Close icon.


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

