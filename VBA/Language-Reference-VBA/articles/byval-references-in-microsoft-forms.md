---
title: ByVal References in Microsoft Forms
keywords: fm20.chm5225276
f1_keywords:
- fm20.chm5225276
ms.prod: office
ms.assetid: 0523f039-caa8-823c-ed4d-27e4dc3561f6
ms.date: 06/08/2017
---


# ByVal References in Microsoft Forms

The ByVal keyword in Microsoft Forms indicates that an argument is passed as a value; this is the standard meaning of ByVal in Visual Basic. However, in Microsoft Forms, you can use ByVal with a ReturnBoolean, ReturnEffect, ReturnInteger, or ReturnString object. When you do, the value passed is not a simple data type; it is a pointer to the object.

When used with these objects, ByVal refers to the object, not the method of passing parameters. Each of the objects listed above has a  **Value** property that you can set. You can also pass that value into and out of a function. Because you can change the values of the object's members, events produce results consistent with ByRef behavior, even though the event syntax says the parameter is ByVal.

Assigning a value to an argument associated with a ReturnBoolean, ReturnEffect, ReturnInteger, or ReturnString is no different from setting the value of any other argument. For example, if the event syntax indicates a  _Cancel_ argument used with the ReturnBoolean object, the statement is still valid, just as it is with other data types.


