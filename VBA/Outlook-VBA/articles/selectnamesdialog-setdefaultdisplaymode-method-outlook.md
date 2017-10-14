---
title: SelectNamesDialog.SetDefaultDisplayMode Method (Outlook)
keywords: vbaol11.chm836
f1_keywords:
- vbaol11.chm836
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.SetDefaultDisplayMode
ms.assetid: d6df1ad3-22b1-bda1-532a-a3bd34aa4ad1
ms.date: 06/08/2017
---


# SelectNamesDialog.SetDefaultDisplayMode Method (Outlook)

Sets the default display mode for the  **Select Names** dialog box, specifying its caption and button labels.


## Syntax

 _expression_ . **SetDefaultDisplayMode**( **_defaultMode_** )

 _expression_ A variable that represents a **SelectNamesDialog** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _defaultMode_|Required| **[OlDefaultSelectNamesDisplayMode](oldefaultselectnamesdisplaymode-enumeration-outlook.md)**|A constant in the  **OlDefaultSelectNamesDisplayMode** enumeration that determines the default caption and button labels for the **Select Names** dialog box.|

## Remarks

 **SetDefaultDisplayMode** is optional. If you do not call **SetDefaultDisplayMode** before calling **[Display](selectnamesdialog-display-method-outlook.md)** , the default display mode will be **OlDefaultSelectNamesDisplayMode.olDefaultMail** . To set the display mode to a different value, you should call **SetDefaultDisplayMode** before calling the **Display** method.

This method allows you to display the  **Select Names** dialog box without using a resource file to localize the values for the caption, the **To** label, **Cc** label, and **Bcc** label. You can override the built-in behavior by setting your own values for **[Caption](selectnamesdialog-caption-property-outlook.md)** , **[ToLabel](selectnamesdialog-tolabel-property-outlook.md)** , **[CcLabel](selectnamesdialog-cclabel-property-outlook.md)** , and **[BccLabel](selectnamesdialog-bcclabel-property-outlook.md)** .

You can set additional properties (for example, setting  **[NumberOfRecipientSelectors](selectnamesdialog-numberofrecipientselectors-property-outlook.md)** to **olRecipientSelectors.olToCc** ) after calling **SetDefaultDisplayMode** . The **Select Names** dialog box will observe the subsequent setting.


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

