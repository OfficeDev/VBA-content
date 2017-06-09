---
title: Documents Object (Visio)
keywords: vis_sdr.chm10085
f1_keywords:
- vis_sdr.chm10085
ms.prod: visio
api_name:
- Visio.Documents
ms.assetid: e9291149-964e-c6fb-4c62-bf2f35a6a0a7
ms.date: 06/08/2017
---


# Documents Object (Visio)

 Includes a **Document** object for each open document in a Microsoft Visio instance.


## Remarks

To retrieve a  **Documents** collection, use the **Documents** property of an **Application** object.

The default property of a  **Documents** collection is **Item** .

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocuments.GetEnumerator()** (to enumerate the **Document** objects.)
    
-  **Microsoft.Office.Interop.Visio.IVDocuments** (to access the collection.)
    

