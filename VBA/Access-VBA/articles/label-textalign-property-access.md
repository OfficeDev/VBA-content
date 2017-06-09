---
title: Label.TextAlign Property (Access)
keywords: vbaac10.chm10216
f1_keywords:
- vbaac10.chm10216
ms.prod: access
api_name:
- Access.Label.TextAlign
ms.assetid: 088c8577-2057-8936-6a47-3c304c8e0eb2
ms.date: 06/08/2017
---


# Label.TextAlign Property (Access)

The  **TextAlign** property specifies the text alignment in new controls. Read/write **Byte**.


## Syntax

 _expression_. **TextAlign**

 _expression_ A variable that represents a **Label** object.


## Remarks

The  **TextAlign** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|General|0|(Default) The text aligns to the left; numbers and dates align to the right.|
|Left|1|The text, numbers, and dates align to the left.|
|Center|2|The text, numbers, and dates are centered.|
|Right|3|The text, numbers, and dates align to the right.|
|Distribute|4|The text, numbers, and dates are evenly distributed.|
You can set the default for the  **TextAlign** property by using a control's default control style or the **DefaultControl** property in Visual Basic.


## Example

The following example aligns the text in the "Address" text box on the "Suppliers" form to the right.


```vb
Forms("Suppliers").Controls("Address").TextAlign = 3
```


## See also


#### Concepts


[Label Object](label-object-access.md)

