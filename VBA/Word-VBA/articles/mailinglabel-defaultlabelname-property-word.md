---
title: MailingLabel.DefaultLabelName Property (Word)
keywords: vbawd10.chm152502281
f1_keywords:
- vbawd10.chm152502281
ms.prod: word
api_name:
- Word.MailingLabel.DefaultLabelName
ms.assetid: f874d60e-e75d-a8b8-6118-e73e467920f9
ms.date: 06/08/2017
---


# MailingLabel.DefaultLabelName Property (Word)

Returns or sets the name for the default mailing label. Read/write  **String** .


## Syntax

 _expression_ . **DefaultLabelName**

 _expression_ A variable that represents a **[MailingLabel](mailinglabel-object-word.md)** object.


## Remarks

To find the string for the specified built-in label, select the label in the  **Label Options** dialog box ( **Tools** menu, **Envelopes and Labels** dialog box, **Labels** tab, **Options** button). Then click **Details** and view the **Label** name box, which contains the correct string to use for this property. To set a custom label as the default mailing label, use the label name that appears in the **Details** dialog box, or use the **[Name](customlabel-name-property-word.md)** property with a **[CustomLabel](customlabel-object-word.md)** object.

Creating a new label document from a  **CustomLabel** object automatically sets the **DefaultLabelName** property to the name of the **CustomLabel** object.


## Example

This example returns the name of the current default mailing label.


```
Msgbox Application.MailingLabel.DefaultLabelName
```

This example sets the Avery Standard, 5160 Address label as the default mailing label.




```vb
Application.MailingLabel.DefaultLabelName = "5160"
```


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

