---
title: SelectNamesDialog.NumberOfRecipientSelectors Property (Outlook)
keywords: vbaol11.chm834
f1_keywords:
- vbaol11.chm834
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.NumberOfRecipientSelectors
ms.assetid: 2cb40e5f-b122-d032-9343-54fe98bc5455
ms.date: 06/08/2017
---


# SelectNamesDialog.NumberOfRecipientSelectors Property (Outlook)

Returns or sets a  **[OlRecipientSelectors](olrecipientselectors-enumeration-outlook.md)** constant that determines the number of recipient edit boxes (each associated with a command button) displayed in the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **NumberOfRecipientSelectors**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

A recipient edit box allows you to enter recipient names. Each recipient edit box is associated with a command button in the  **Select Names** dialog box; examples of a command button for a recipient edit box are the **To** and **Cc** command buttons. The default value of **NumberOfRecipientSelectors** is **OlRecipientSelectors.olToCcBcc** .

If you set  **NumberOfRecipientSelectors** to **OlRecipientSelectors.olShowTo** and then subsequently set the text for **[SelectNamesDialog.CcLabel](selectnamesdialog-cclabel-property-outlook.md)** or **[SelectNamesDialog.BccLabel](selectnamesdialog-bcclabel-property-outlook.md)** , the **NumberOfRecipientSelectors** will remain unchanged.

If you set  **NumberOfRecipientSelectors** to **OlRecipientSelectors.olShowNone** , then the **[SelectNamesDialog.AllowMultipleSelection](selectnamesdialog-allowmultipleselection-property-outlook.md)** property will be ignored.


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

