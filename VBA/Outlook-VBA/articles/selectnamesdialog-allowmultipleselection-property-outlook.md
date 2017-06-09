---
title: SelectNamesDialog.AllowMultipleSelection Property (Outlook)
keywords: vbaol11.chm831
f1_keywords:
- vbaol11.chm831
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.AllowMultipleSelection
ms.assetid: e8b67f2a-b6c1-16af-6762-801536d4f93f
ms.date: 06/08/2017
---


# SelectNamesDialog.AllowMultipleSelection Property (Outlook)

Returns or sets a  **Boolean** that determines whether more than one address entry can be selected at a time in the **Select Names** dialog. Read/write.


## Syntax

 _expression_ . **AllowMultipleSelection**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

The default value of  **AllowMultipleSelection** is **True**. If  **AllowMultipleSelection** is set to **True**, the user can select multiple recipients by using the  **CTRL** or **SHIFT** key. If **AllowMultipleSelection** is set to **False**, multiple selection is disabled. 

Setting  **AllowMultipleSelection** to **False** does not ensure that only one recipient can be selected. The user can type additional recipients in the edit box or select from the recipient list multiple times. To ensure that only one recipient can be selected in the dialog, set **AllowMultipleSelect** to **False** and **[SelectNamesDialog.NumberOfRecipientSelectors](selectnamesdialog-numberofrecipientselectors-property-outlook.md)** to **olShowNone** .


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

