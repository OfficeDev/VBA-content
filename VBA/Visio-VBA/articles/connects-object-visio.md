---
title: Connects Object (Visio)
keywords: vis_sdr.chm10070
f1_keywords:
- vis_sdr.chm10070
ms.prod: visio
api_name:
- Visio.Connects
ms.assetid: 8ac06fd8-0bbb-e9df-a08c-d697c4ac238e
ms.date: 06/08/2017
---


# Connects Object (Visio)

 Includes a **Connect** object for each connection between two shapes in a drawing, such as a line and a box in an organization chart.


## Remarks

The default property of a  **Connects** collection is **Item**.

Use the  **Connects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object to which the indicated **Shape** object is connected (glued).

Use the  **FromConnects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object that is connected (glued) to the indicated **Shape** object.

Use the  **Connects** property of a **Page** object to retrieve a **Connects** collection with an entry for every connection on the **Page** object.

Use the  **Connects** property of a **Master** object to retrieve a **Connects** collection with an entry for every connection in the **Master** object.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVConnects.GetEnumerator()** (to enumerate the **Connect** objects.)
    

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/add9261d-b2e7-f96f-55c2-8326f8b39813%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/52be9eb0-5130-2490-98a0-58215dead3d5%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/7a161248-2bf5-42e7-772d-e0f4de979776%28Office.15%29.aspx)|
|[FromSheet](http://msdn.microsoft.com/library/c9fa472c-9f5f-ea4f-adbc-e8741dda1482%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/3b43a3ae-cf92-cc05-2750-c37554d9202c%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/1d27b1c5-89f0-493c-b90c-9be46fc93ca0%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/e51e58fb-a7b5-b18e-3f53-8ab1ff4d2994%28Office.15%29.aspx)|
|[ToSheet](http://msdn.microsoft.com/library/a5884fda-45cb-9b2b-da19-788db429e6f1%28Office.15%29.aspx)|

