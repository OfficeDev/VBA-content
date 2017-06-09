---
title: This interaction between compiled and design environment components is not supported
keywords: vblr6.chm1000373
f1_keywords:
- vblr6.chm1000373
ms.prod: office
ms.assetid: e24e520f-9561-deac-f2be-ba14af1db6ed
ms.date: 06/08/2017
---


# This interaction between compiled and design environment components is not supported

This error has the following causes and solutions:



- This occurs when two components are running together, where one component (such as a form or a UserControl) was previously compiled and is now running using the runtime (msvbvm60.dll), and the other component is being run in the IDE. For example, a compiled UserControl running on a form in the IDE. The problem occurs because the internal memory structure between an item running in the IDE and a compiled object is slightly different and not always compatible. In general, though, you shouldn't encounter a problem with this unless you are passing an instance of a UserControl (Me) to a host form or other component through a  **Property** or **Sub** procedure.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

