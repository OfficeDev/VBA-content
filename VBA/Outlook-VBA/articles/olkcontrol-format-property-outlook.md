---
title: OlkControl.Format Property (Outlook)
keywords: vbaol11.chm1000542
f1_keywords:
- vbaol11.chm1000542
ms.prod: outlook
api_name:
- Outlook.OlkControl.Format
ms.assetid: f2fbaf25-ae06-b954-0de2-a368ce023fb0
ms.date: 06/08/2017
---


# OlkControl.Format Property (Outlook)

Returns or sets a  **Long** that specifies how a value is to be displayed in the control. Read/write.


## Syntax

 _expression_ . **Format**

 _expression_ A variable that represents an **OlkControl** object.


## Remarks

The  **Format** property can be a constant in an enumeration that describes how to display a value. For example, you can specify **Format** as the constant **olFormatCurrencyDecimal** that is defined in the **[OlFormatCurrency](olformatcurrency-enumeration-outlook.md)** enumeration to display a currency value in an **[OlkTextBox](olktextbox-object-outlook.md)** control.

The  **Format** property is specific to the property in the Outlook Object Model that the control is bound to. The latter is indicated by **[OlkControl.ItemProperty](olkcontrol-itemproperty-property-outlook.md)** . If the control is not bound to any property, then accessing **Format** will return an error.


## See also


#### Concepts


[OlkControl Class](olkcontrol-object-outlook.md)

