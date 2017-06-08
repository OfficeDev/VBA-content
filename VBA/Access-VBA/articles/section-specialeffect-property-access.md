---
title: Section.SpecialEffect Property (Access)
keywords: vbaac10.chm12202
f1_keywords:
- vbaac10.chm12202
ms.prod: access
api_name:
- Access.Section.SpecialEffect
ms.assetid: 5143b888-5263-6ccd-1d98-3c7168ca7734
ms.date: 06/08/2017
---


# Section.SpecialEffect Property (Access)

You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.


## Syntax

 _expression_. **SpecialEffect**

 _expression_ A variable that represents a **Section** object.


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


## Example

The following example sets the appearance of the text box "OrganizationName1" on the "Mailing List" form to raised.


```vb
Forms("Mailing List").Controls("OrganizationName1").SpecialEffect = 1
```


## See also


#### Concepts


[Section Object](section-object-access.md)

