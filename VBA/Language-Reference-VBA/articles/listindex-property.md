---
title: ListIndex Property
keywords: fm20.chm2001430
f1_keywords:
- fm20.chm2001430
ms.prod: office
api_name:
- Office.ListIndex
ms.assetid: ec6f8ad0-472b-18e3-6c56-bf2e94152504
ms.date: 06/08/2017
---


# ListIndex Property



Identifies the currently selected item in a  **ListBox** or **ComboBox**.
 **Syntax**
 _object_. **ListIndex** [= _Variant_ ]
The  **ListIndex** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Variant_|Optional. The currently selected item in the control.|
 **Remarks**
The  **ListIndex** property contains an index of the selected row in a list. Values of **ListIndex** range from -1 to one less than the total number of rows in a list (that is, **ListCount** - 1). When no rows are selected, **ListIndex** returns -1. When the user selects a row in a **ListBox** or **ComboBox**, the system sets the **ListIndex** value. The **ListIndex** value of the first row in a list is 0, the value of the second row is 1, and so on.

 **Note**  If you use the  **MultiSelect** property to create a **ListBox** that allows multiple selections, the **Selected** property of the **ListBox** (rather than the **ListIndex** property) identifies the selected rows. The **Selected** property is an[array](vbe-glossary.md) with the same number of values as the number of rows in the **ListBox**. For each row in the list box, **Selected** is **True** if the row is selected and **False** if it is not. In a **ListBox** that allows multiple selections, **ListIndex** returns the index of the row that has[focus](vbe-glossary.md), regardless of whether that row is currently selected.

The  **ListIndex** value is also available by setting the **BoundColumn** property to 0 for a combo box or list box. If **BoundColumn** is 0, the underlying[data source](glossary-vba.md) to which the combo box or list box is[bound](glossary-vba.md) contains the same list index value as **ListIndex**.

