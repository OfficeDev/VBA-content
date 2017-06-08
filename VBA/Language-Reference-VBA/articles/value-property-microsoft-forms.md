---
title: Value Property (Microsoft Forms)
keywords: fm20.chm2002180
f1_keywords:
- fm20.chm2002180
ms.prod: office
ms.assetid: bd61f3ae-54b3-6382-6ecf-0c5598279330
ms.date: 06/08/2017
---


# Value Property (Microsoft Forms)



Specifies the state or content of a given control.
 **Syntax**
 _object_. **Value** [= _Variant_ ]
The  **Value** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Variant_|Optional. The state or content of the control.|
 **Settings**


|**Control**|**Description**|
|:-----|:-----|
|**CheckBox**|An integer value indicating whether the item is selected:|
||Null Indicates the item is in a null state, neither selected nor [cleared](glossary-vba.md).|
||-1 True. Indicates the item is selected.|
||0 False. Indicates the item is cleared.|
|**OptionButton**|Same as  **CheckBox**.|
|**ToggleButton**|Same as  **CheckBox**.|
|**ScrollBar**|An integer between the values specified for the  **Max** and **Min** properties.|
|**SpinButton**|Same as  **ScrollBar**.|
|**ComboBox, ListBox**|The value in the  **BoundColumn** of the currently selected rows.|
|**CommandButton**|Always  **False**.|
|**MultiPage**|An integer indicating the currently active page.|
||Zero (0) indicates the first page. The maximum value is one less than the number of pages.|
|**TextBox**|The text in the edit region.|
 **Remarks**
For a  **CommandButton**, setting the **Value** property to **True** in a macro or procedure initiates the button's Click event.
For a  **ComboBox**, changing the contents of **Value** does not change the value of **BoundColumn**. To add or delete entries in a **ComboBox**, you can use the **AddItem** or **RemoveItem** method.
 **Value** cannot be used with a multi-select list box.

