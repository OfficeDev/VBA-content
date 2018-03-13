---
title: ShowDropButtonWhen Property
keywords: fm20.chm2001900
f1_keywords:
- fm20.chm2001900
ms.prod: office
api_name:
- Office.ShowDropButtonWhen
ms.assetid: 82c7a038-a4fa-7253-ec24-c97e6841293e
ms.date: 06/08/2017
---


# ShowDropButtonWhen Property



Specifies when to show the drop-down button for a  **ComboBox** or **TextBox**.
 **Syntax**
 _object_. **ShowDropButtonWhen** [= _fmShowDropButtonWhen_ ]
The  **ShowDropButtonWhen** property syntax has these parts:


| <strong>Part</strong>         | <strong>Description</strong>                                                  |
|:------------------------------|:------------------------------------------------------------------------------|
| <em>object</em>               | Required. A valid object.                                                     |
| <em>fmShowDropButtonWhen</em> | Optional. The circumstances under which the drop-down button will be visible. |

 **Settings**
The settings for  _fmShowDropButtonWhen_ are:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                              |
|:------------------------------------|:-----------------------|:----------------------------------------------------------|
| <em>fmShowDropButtonWhenNever</em>  | 0                      | Do not show the drop-down button under any circumstances. |
| <em>fmShowDropButtonWhenFocus</em>  | 1                      | Show the drop-down button when the control has the focus. |
| <em>fmShowDropButtonWhenAlways</em> | 2                      | Always show the drop-down button.                         |

For a  **ComboBox**, the default value is _fmShowDropButtonWhenAlways_; for a **TextBox**, the default value is _fmShowDropButtonWhenNever_.

