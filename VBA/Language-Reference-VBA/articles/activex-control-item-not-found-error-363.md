---
title: ActiveX control 'item' not found (Error 363)
keywords: vblr6.chm1117794
f1_keywords:
- vblr6.chm1117794
ms.prod: office
ms.assetid: 5c97e208-a788-f8af-6fd7-f80ab7728c12
ms.date: 06/08/2017
---


# ActiveX control 'item' not found (Error 363)

The form being loaded contains an [ActiveX control](vbe-glossary.md) that isn't part of the current[project](vbe-glossary.md). This error has the following causes and solutions:



- You may have manually edited the project's .vbp file to add a form containing an ActiveX control that isn't already part of the project. After the project loads, use the  **Add File** command on the **File** menu to add the ActiveX control to the project.
    
- You may have manually edited the project's .vbp file to add a form containing an earlier version of an ActiveX control than the ActiveX control that is already part of the project. After the project loads, delete the earlier version from the form and put the later version of the control on the form.
    


