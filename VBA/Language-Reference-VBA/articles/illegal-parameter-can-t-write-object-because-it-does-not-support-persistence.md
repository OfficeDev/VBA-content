---
title: Illegal parameter. Can't write object because it does not support persistence.
keywords: vblr6.chm1000330
f1_keywords:
- vblr6.chm1000330
ms.prod: office
ms.assetid: a8ef56eb-88b5-eb7f-be83-e6b18b9756bc
ms.date: 06/08/2017
---


# Illegal parameter. Can't write object because it does not support persistence.

This error has the following causes and solutions:



- You attempted to write an object to a PropertyBag object, but the object doesn't support one of the ActiveX persistence interfaces. To fix this, you must have access to the code for the object. It must a Visual Basic-created object and have its Persistable property set to True. Also, the Class must be either Public or Public Createable and be in an ActiveX Dll, ActiveX Exe, or UserControl project.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

