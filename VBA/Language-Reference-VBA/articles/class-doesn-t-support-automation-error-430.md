---
title: Class doesn't support Automation (Error 430)
keywords: vblr6.chm1000430
f1_keywords:
- vblr6.chm1000430
ms.prod: office
ms.assetid: f3d5d8a8-4d53-f8bc-b5dc-62f0820fe8fc
ms.date: 06/08/2017
---


# Class doesn't support Automation (Error 430)

Not all objects expose an Automation interface. This error has the following cause and solution:



- The [class](vbe-glossary.md) you specified in the **GetObject** or **CreateObject** function call was found, but has not exposed a programmability interface.
    
    You can't write code to control an object's behavior unless it has been exposed for Automation. Check the documentation of the application that created the object for limitations on the use of Automation with this class of object.
    
- You changed a project from .dll to .exe, or vice versa. If, for example, you have a .dll server already compiled and registered, and then you change the project type to .exe and recompile it, the fact that the .dll and .exe are already registered on your system prevents you from creating either object. You must manually unregister the old .dll or .exe to avoid the problem. This is caused by the combination of project compatibility and changing a project from an .exe to a .dll. In project compatibility, the CLSID is preserved, but not the IID. Since the CLSID is preserved, the class ends up being registered with two servers — one an in-process server, the other a local server. When an instance is created, the in-process one is chosen. When the querying of the interface occurs, the .dll does not support the IID because it's new.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

