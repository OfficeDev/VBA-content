---
title: Global Object (Visio)
keywords: vis_sdr.chm10000
f1_keywords:
- vis_sdr.chm10000
ms.prod: visio
api_name:
- Visio.Global
ms.assetid: 3c7dca10-f7b0-f3f7-59b1-7845338aa4a4
ms.date: 06/08/2017
---


# Global Object (Visio)

The Microsoft Visio global object is automatically available to Microsoft Visual Basic for Applications (VBA) code that is part of the VBA project of a Visio document. The Visio global object is not available to code in other contexts.


## Remarks

Members of the global object can be accessed without qualification. For example, to access the  **ActivePage** member of the global object:


```vb
    Set vsoPage = ActivePage 
```

The preceding syntax is different from the syntax you would use for accessing members of non-global objects. For example:




```vb
    Set vsoPage = vsoApplication.ActivePage
```


 **Note**  The VBA project of every Visio document also has a class module called  **ThisDocument** . When referenced from code in the VBA project, the **ThisDocument** module returns a reference to the project's **Document** object.


