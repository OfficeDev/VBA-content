---
title: Application.Documents Property (Visio)
keywords: vis_sdr.chm10013435
f1_keywords:
- vis_sdr.chm10013435
ms.prod: visio
api_name:
- Visio.Application.Documents
ms.assetid: dee2a72f-526c-7b10-57b4-c4fbca43b083
ms.date: 06/08/2017
---


# Application.Documents Property (Visio)

Returns the  **Documents** collection for a Microsoft Visio instance. Read-only.


## Syntax

 _expression_ . **Documents**

 _expression_ A variable that represents an **Application** object.


### Return Value

Documents


## Remarks

You can iterate through a  **Documents** collection by using the **Count** property to retrieve the number of documents in the collection. You can use the **Item** property to retrieve individual elements from a collection.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVApplication.Documents**
    

