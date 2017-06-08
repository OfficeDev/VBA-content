---
title: Shapes.ItemFromID Property (Visio)
keywords: vis_sdr.chm11313775
f1_keywords:
- vis_sdr.chm11313775
ms.prod: visio
api_name:
- Visio.Shapes.ItemFromID
ms.assetid: 0e8e80a2-94f0-f451-b914-f8d8a56a3ef2
ms.date: 06/08/2017
---


# Shapes.ItemFromID Property (Visio)

Returns an item of a collection using the ID of the item. Read-only.


## Syntax

 _expression_ . **ItemFromID**( **_nID_** )

 _expression_ A variable that represents a **Shapes** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nID_|Required| **Long**|The ID of the object to retrieve.|

### Return Value

Shape


## Remarks

The ID of a  **Shape** object uniquely identifies the shape within its page or master. You can determine the ID of a shape by displaying the **Special** dialog box (select the shape, and then click **Shape Name** on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab.)

The ID of a  **Style** object uniquely identifies the style within its document.

The ID of a  **Font** object corresponds to the number stored in the Font cell of a row in a shape's Character Properties section. The ID associated with a particular font varies between systems or as fonts are installed on and removed from a given system.

The ID of an  **Event** object uniquely identifies an event in its **EventList** collection for the life of the collection.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShapes.get_ItemFromID**
    

