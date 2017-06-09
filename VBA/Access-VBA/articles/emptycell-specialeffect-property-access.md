---
title: EmptyCell.SpecialEffect Property (Access)
keywords: vbaac10.chm14313
f1_keywords:
- vbaac10.chm14313
ms.prod: access
api_name:
- Access.EmptyCell.SpecialEffect
ms.assetid: f5858a41-9ba2-83a8-6457-3a5a04352d5a
ms.date: 06/08/2017
---


# EmptyCell.SpecialEffect Property (Access)

You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.


## Syntax

 _expression_. **SpecialEffect**

 _expression_ A variable that represents an **EmptyCell** object.


## Remarks

The  **SpecialEffect** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Flat|0|The object appears flat and has the system's default colors or custom colors that were set in Design view.|
|Raised|1|The object has a highlight on the top and left and a shadow on the bottom and right.|
|Sunken|2|The object has a shadow on the top and left and a highlight on the bottom and right.|
|Etched|3|The object has a sunken line surrounding the control.|
|Shadowed|4|The object has a shadow below and to the right of the control.|
|Chiseled|5|The object has a sunken line below the control.|
The  **SpecialEffect** property setting affects related property settings for the **BorderStyle**, **BorderColor**, and **BorderWidth** properties. For example, if the **SpecialEffect** property is set to Raised, the settings for the **BorderStyle**, **BorderColor**, and **BorderWidth** properties are ignored. In addition, changing or setting the **BorderStyle**, **BorderColor**, and **BorderWidth** properties may cause Microsoft Access to change the **SpecialEffect** property setting to Flat.


 **Note**  When you set the  **SpecialEffect** property of a text box to Shadowed, the vertical height of the text display area is reduced. You can adjust the **Height** property of the text box to increase the size of the text display area.


## Example

The following example sets the appearance of the text box "OrganizationName1" on the "Mailing List" form to raised.


```vb
Forms("Mailing List").Controls("OrganizationName1").SpecialEffect = 1
```


## See also


#### Concepts


[EmptyCell Object](emptycell-object-access.md)

