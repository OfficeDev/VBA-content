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


| <strong>Part</strong>    | <strong>Description</strong>                                                                                                                                   |
|:-------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>          | Required. A valid object.                                                                                                                                      |
| <em>fmButtonEffect</em>  | Optional. The desired visual appearance for a  <strong>CheckBox</strong>, <strong>OptionButton</strong>, or <strong>ToggleButton</strong>.                     |
| <em>fmSpecialEffect</em> | Optional. The desired visual appearance of an object other than a  <strong>CheckBox</strong>, <strong>OptionButton</strong>, or <strong>ToggleButton</strong>. |

 **Settings**
The settings for  _fmSpecialEffect_ are:


| <strong>Constant</strong>      | <strong>Value</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                  |
|:-------------------------------|:-----------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>fmSpecialEffectFlat</em>   | 0                      | Object appears flat, distinguished from the surrounding form by a border, a change of color, or both. Default for  <strong>Image</strong> and <strong>Label</strong>, valid for all controls.                                                                                 |
| <em>fmSpecialEffectRaised</em> | 1                      | Object has a highlight on the top and left and a shadow on the bottom and right. Not valid for check boxes or option buttons.                                                                                                                                                 |
| <em>fmSpecialEffectSunken</em> | 2                      | Object has a shadow on the top and left and a highlight on the bottom and right. The control and its border appear to be carved into the form that contains them. Default for  <strong>CheckBox</strong> and <strong>OptionButton</strong>, valid for all controls (default). |
| <em>fmSpecialEffectEtched</em> | 3                      | Border appears to be carved around the edge of the control. Not valid for check boxes or option buttons.                                                                                                                                                                      |
| <em>fmSpecialEffectBump</em>   | 6                      | Object has a ridge on the bottom and right and appears flat on the top and left. Not valid for check boxes or option buttons.                                                                                                                                                 |

For a  **Frame**, the default value is _Sunken_.
Note that only  _Flat_ and _Sunken_ (0 and 2) are acceptable values for **CheckBox**, **OptionButton**, and **ToggleButton**. All values listed are acceptable for other controls.
 **Remarks**
You can use either the  **SpecialEffect** or the **BorderStyle** property to specify the edging for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **SpecialEffect** to **fmSpecialEffectRaised**, the system sets **BorderStyle** to zero ( **fmBorderStyleNone** ).
For a  **Frame**, **BorderStyle** is ignored if **SpecialEffect** is **fmSpecialEffectFlat**.
 **SpecialEffect** uses the[system colors](glossary-vba.md) to define its borders.

 **Note**  Although the  **SpecialEffect** property exists on the **ToggleButton**, the property is disabled. You cannot set or return a value for this property on the **ToggleButton**.


