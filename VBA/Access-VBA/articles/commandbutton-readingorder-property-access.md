---
title: CommandButton.ReadingOrder Property (Access)
keywords: vbaac10.chm10505
f1_keywords:
- vbaac10.chm10505
ms.prod: access
api_name:
- Access.CommandButton.ReadingOrder
ms.assetid: 5a47e95d-7421-147f-084a-74130cf524c7
ms.date: 06/08/2017
---


# CommandButton.ReadingOrder Property (Access)

You can use the  **ReadingOrder** property to specify or determine the reading order of words in text. Read/write **Byte**.


## Syntax

 _expression_. **ReadingOrder**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **ReadingOrder** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Context|0|Reading order is determined by the language of the first character entered. If a right-to-left language character is entered first, reading order is right to left. If a left-to-right language character is entered first, reading order is left to right.|
|Left-to-Right|1|Sets the reading order to left to right.|
|Right-to-Left|2|Sets the reading order to right to left.|

## Example

The following example sets the reading order to right to left for the "Address" text box on the "International Shipping" form.


```vb
Forms("International Shipping").Controls("Address").ReadingOrder = 2
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

