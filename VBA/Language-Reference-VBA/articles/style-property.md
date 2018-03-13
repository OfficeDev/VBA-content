---
title: Style Property
ms.prod: office
api_name:
- Office.Style
ms.assetid: b951714c-360e-47c7-85a4-c3260d898b1d
ms.date: 06/08/2017
---


# Style Property



For  **ComboBox**, specifies how the user can choose or set the control's value. For **MultiPage** and **TabStrip**, identifies the style of the tabs on the control.
 **Syntax**
For ComboBox _object_. **Style** [= _fmStyle_ ]
For MultiPage and TabStrip _object_. **Style** [= _fmTabStyle_ ]
The  **Style** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                     |
|:----------------------|:-------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                        |
| <em>fmStyle</em>      | Optional. Specifies how a user sets the value of a  <strong>ComboBox</strong>.                   |
| <em>fmTabStyle</em>   | Optional. Specifies the tab style in a  <strong>MultiPage</strong> or <strong>TabStrip</strong>. |

 **Settings**
The settings for  _fmStyle_ are:


| <strong>Constant</strong>     | <strong>Value</strong> | <strong>Description</strong>                                                                                                                                       |
|:------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>fmStyleDropDownCombo</em> | 0                      | The  <strong>ComboBox</strong> behaves as a drop-down combo box. The user can type a value in the edit region or select a value from the drop-down list (default). |
| <em>fmStyleDropDownList</em>  | 2                      | The  <strong>ComboBox</strong> behaves as a list box. The user must choose a value from the list.                                                                  |

The settings for  _fmTabStyle_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmTabStyleTabs_|0|Displays tabs on the tab bar (default).|
| _fmTabStyleButtons_|1|Displays buttons on the tab bar.|
| _fmTabStyleNone_|2|Does not display the tab bar.|

