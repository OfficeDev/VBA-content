---
title: IVisEventProc Object (Visio)
keywords: vis_sdr.chm60150
f1_keywords:
- vis_sdr.chm60150
ms.prod: visio
ms.assetid: 332ec60d-c70a-9d7f-15ad-bb797f60b3a5
ms.date: 06/08/2017
---


# IVisEventProc Object (Visio)

The interface for handling event notifications in Microsoft Visio. 


## Remarks

In addition to the methods inherited from  **IDispatch** , the **IVisEventProc** interface contains a single function, **VisEventProc** , which returns a **Variant** . Because **IVisEventProc** inherits from **IDispatch** and hence from **IUnknown** , you must implement the methods in those interfaces as well as the **VisEventProc** method in **IVisEventProc** .

To handle event notifications in Visio, create a class module that implements the  **IVisEventProc** interface in Microsoft Visual Basic for Applications (VBA) or Microsoft Visual Basic, and then create an instance of this class to pass as an argument to the **AddAdvise** method of the **EventList** collection.

For more information about using the  **IVisEventProc** interface to handle event notifications, search for "Visio event objects" on MSDN, the Microsoft Developer Network. For more information about implementing IDispatch methods, search for "Implementing the IDispatch Interface" on MSDN.


