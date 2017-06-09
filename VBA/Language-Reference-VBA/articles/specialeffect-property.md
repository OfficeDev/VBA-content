---
title: SpecialEffect Property
keywords: fm20.chm5225095
f1_keywords:
- fm20.chm5225095
ms.prod: office
api_name:
- Office.SpecialEffect
ms.assetid: db3fa148-42f3-fded-9ec4-6d46c963fdce
ms.date: 06/08/2017
---


# SpecialEffect Property



Specifies the visual appearance of an object.
 **Syntax**
For CheckBox, OptionButton, ToggleButton _object_. **SpecialEffect** [= _fmButtonEffect_ ]
For other controls _object_. **SpecialEffect** [= _fmSpecialEffect_ ]
The  **SpecialEffect** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmButtonEffect_|Optional. The desired visual appearance for a  **CheckBox**, **OptionButton**, or **ToggleButton**.|
| _fmSpecialEffect_|Optional. The desired visual appearance of an object other than a  **CheckBox**, **OptionButton**, or **ToggleButton**.|
 **Settings**
The settings for  _fmSpecialEffect_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmSpecialEffectFlat_|0|Object appears flat, distinguished from the surrounding form by a border, a change of color, or both. Default for  **Image** and **Label**, valid for all controls.|
| _fmSpecialEffectRaised_|1|Object has a highlight on the top and left and a shadow on the bottom and right. Not valid for check boxes or option buttons.|
| _fmSpecialEffectSunken_|2|Object has a shadow on the top and left and a highlight on the bottom and right. The control and its border appear to be carved into the form that contains them. Default for  **CheckBox** and **OptionButton**, valid for all controls (default).|
| _fmSpecialEffectEtched_|3|Border appears to be carved around the edge of the control. Not valid for check boxes or option buttons.|
| _fmSpecialEffectBump_|6|Object has a ridge on the bottom and right and appears flat on the top and left. Not valid for check boxes or option buttons.|
For a  **Frame**, the default value is _Sunken_.
Note that only  _Flat_ and _Sunken_ (0 and 2) are acceptable values for **CheckBox**, **OptionButton**, and **ToggleButton**. All values listed are acceptable for other controls.
 **Remarks**
You can use either the  **SpecialEffect** or the **BorderStyle** property to specify the edging for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **SpecialEffect** to **fmSpecialEffectRaised**, the system sets **BorderStyle** to zero ( **fmBorderStyleNone** ).
For a  **Frame**, **BorderStyle** is ignored if **SpecialEffect** is **fmSpecialEffectFlat**.
 **SpecialEffect** uses the[system colors](glossary-vba.md) to define its borders.

 **Note**  Although the  **SpecialEffect** property exists on the **ToggleButton**, the property is disabled. You cannot set or return a value for this property on the **ToggleButton**.


