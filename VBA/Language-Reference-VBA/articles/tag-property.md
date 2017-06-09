---
title: Tag Property
keywords: fm20.chm2002060
f1_keywords:
- fm20.chm2002060
ms.prod: office
api_name:
- Office.Tag
ms.assetid: 9cc2496d-f3c9-fca0-1e48-eb4ed0905b51
ms.date: 06/08/2017
---


# Tag Property



Stores additional information about an object.
 **Syntax**
 _object_. **Tag** [= _String_ ]
The  **Tag** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression identifying the object. The default is a zero-length string ("").|
 **Remarks**
Use the  **Tag** property to assign an identification string to an object without affecting other property settings or attributes.
For example, you can use  **Tag** to check the identity of a form or control that is passed as a variable to a procedure.

