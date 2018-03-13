---
title: ListStyle Property
keywords: fm20.chm2001450
f1_keywords:
- fm20.chm2001450
ms.prod: office
api_name:
- Office.ListStyle
ms.assetid: b07cb0d3-7782-7fe4-dea2-9cfddebf3096
ms.date: 06/08/2017
---


# ListStyle Property



Specifies the visual appearance of the list in a  **ListBox** or **ComboBox**.
 **Syntax**
 _object_. **ListStyle** [= _fmListStyle_ ]
The  **ListStyle** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>            |
|:----------------------|:----------------------------------------|
| <em>object</em>       | Required. A valid object.               |
| <em>fmListStyle</em>  | Optional. The visual style of the list. |

 **Settings**
The settings for  _fmListStyle_ are:


| <strong>Constant</strong>  | <strong>Value</strong> | <strong>Description</strong>                                                                                                                                                                                                                          |
|:---------------------------|:-----------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>fmListStylePlain</em>  | 0                      | Looks like a regular list box, with the background of items highlighted.                                                                                                                                                                              |
| <em>fmListStyleOption</em> | 1                      | Shows option buttons, or check boxes for a multi-select list (default). When the user selects an item from the group, the option button associated with that item is selected and the option buttons for the other items in the group are deselected. |

 **Remarks**
The  **ListStyle** property lets you change the visual presentation of a **ListBox** or **ComboBox**. By specifying a setting other than **fmListStylePlain**, you can present the contents of either control as a group of individual items, with each item including a visual cue to indicate whether it is selected.
If the control supports a single selection (the  **MultiSelect** property is set to **fmMultiSelectSingle** ), the user can press one button in the group. If the control supports multi-select, the user can press two or more buttons in the group.

