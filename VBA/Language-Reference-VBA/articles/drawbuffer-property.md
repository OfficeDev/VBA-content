---
title: DrawBuffer Property
keywords: fm20.chm5225033
f1_keywords:
- fm20.chm5225033
ms.prod: office
api_name:
- Office.DrawBuffer
ms.assetid: 6f859070-13c0-5da3-40e6-51f6676cec3b
ms.date: 06/08/2017
---


# DrawBuffer Property



Specifies the number of pixels set aside for off-screen memory in rendering a frame.
 **Syntax**
 _object_. **DrawBuffer** [= _value_ ]


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _value_|An integer from 16,000 through 1,048,576 equal to the maximum number of pixels the object can render off-screen.|
 **Remarks**
The  **DrawBuffer** property specifies the maximum number of pixels that can be drawn at one time as the display repaints. The actual memory used by the object depends upon the screen resolution of the display. If you set a large value for **DrawBuffer**, performance will be slower. A large buffer only helps when several large images overlap.
Use the  **Properties** window to specify the value of **DrawBuffer**.

