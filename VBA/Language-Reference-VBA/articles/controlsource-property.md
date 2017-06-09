---
title: ControlSource Property
keywords: fm20.chm2000980
f1_keywords:
- fm20.chm2000980
ms.prod: office
api_name:
- Office.ControlSource
ms.assetid: 69e5e7bb-5be9-2cca-7693-ac9020578762
ms.date: 06/08/2017
---


# ControlSource Property



Identifies the data location used to set or store the  **Value** property of a control. The **ControlSource** property accepts worksheet ranges from Microsoft Excel.
 **Syntax**
 _object_. **ControlSource** [= _String_ ]
The  **ControlSource** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. Specifies the worksheet cell linked to the  **Value** property of a control.|
 **Remarks**
The  **ControlSource** property identifies a cell or field; it does not contain the data stored in the cell or field. If you change the **Value** of the control, the change is automatically reflected in the linked cell or field. Similarly, if you change the value of the linked cell or field, the change is automatically reflected in the **Value** of the control.
You cannot specify another control for the  **ControlSource**. Doing so causes an error.
The default value for  **ControlSource** is an empty string. If **ControlSource** contains a value other than an empty string, it identifies a linked cell or field. The contents of that cell or field are automatically copied to the **Value** property when the control is loaded.

 **Note**  If the  **Value** property is **Null**, no value appears in the location identified by **ControlSource**.


