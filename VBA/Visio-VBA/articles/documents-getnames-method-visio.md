---
title: Documents.GetNames Method (Visio)
keywords: vis_sdr.chm10616315
f1_keywords:
- vis_sdr.chm10616315
ms.prod: visio
api_name:
- Visio.Documents.GetNames
ms.assetid: 3924a293-c07b-89fa-06d9-35e0e5395c9a
ms.date: 06/08/2017
---


# Documents.GetNames Method (Visio)

Returns the names of all items in a collection.


## Syntax

 _expression_ . **GetNames**( **_NameArray()_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameArray()_|Required| **String**|Out paramter. An array that receives the names of members of the indicated object.|

### Return Value

Nothing


## Remarks

If the  **GetNames** method succeeds, _NameArray()_ returns a one-dimensional array of _n_ strings indexed from 0 to _n_ - 1, where _n_ equals the **Count** property of the object. _NameArray()_ is an out paramter that is allocated by the **GetNames** method, which passes ownership back to the caller. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. Note that the **SafeArrayDestroy** procedure has the side effect of freeing the strings referenced by the array's entries. (Microsoft Visual Basic and Visual Basic for Applications take care of this for you.)


 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **GetNames** method to get the local name of more than one object. Use the **GetNamesU** method to get the universal name of more than one object.


