---
title: HelpContextID Property
keywords: fm20.chm2001260
f1_keywords:
- fm20.chm2001260
ms.prod: office
api_name:
- Office.HelpContextID
ms.assetid: 734940ce-ee04-09d6-7911-7b303beadf23
ms.date: 06/08/2017
---


# HelpContextID Property



The  **HelpContextID** property associates a specific topic in a custom Microsoft Windows Help file with a specific control.
 **Syntax**
 _object_. **HelpContextID** [= _Long_ ]
The  **HelpContextID** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A positive integer specifies the [context ID](glossary-vba.md) of a topic in the Help file associated with the object. Zero indicates no Help topic is associated with the object (default). Must be a valid context ID in the specified Help file.|
 **Remarks**
The topic identified by the  **HelpContextID** property is available to users when a form is running. To display the topic, the user must either select the control or set[focus](vbe-glossary.md) to the control, and then press F1.
The  **HelpContextID** property refers to a topic in a custom Help file you have created to describe your form or application. In Visual Basic, the custom Help file is a property of the[project](vbe-glossary.md).

