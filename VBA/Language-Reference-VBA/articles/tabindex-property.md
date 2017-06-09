---
title: TabIndex Property
keywords: fm20.chm2002010
f1_keywords:
- fm20.chm2002010
ms.prod: office
api_name:
- Office.TabIndex
ms.assetid: 5924d02f-d96c-2b81-6c41-c69ea68ad048
ms.date: 06/08/2017
---


# TabIndex Property



Specifies the position of a single object in the form's [tab order](vbe-glossary.md).
 **Syntax**
 _object_. **TabIndex** [= _Integer_ ]
The  **TabIndex** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Integer_|Optional. An integer from 0 to one less than the number of controls on the form that have a  **TabIndex** property. Assigning a **TabIndex** value of less than 0 generates an error. If you assign a **TabIndex** value greater than the largest index value, the system resets the value to the maximum allowable value.|
 **Remarks**
The index value of the first object in the tab order is zero.

