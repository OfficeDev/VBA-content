---
title: Style.GetFormulasU Method (Visio)
keywords: vis_sdr.chm11451935
f1_keywords:
- vis_sdr.chm11451935
ms.prod: visio
api_name:
- Visio.Style.GetFormulasU
ms.assetid: eadb8801-3fba-6c3d-214a-98a172555403
ms.date: 06/08/2017
---


# Style.GetFormulasU Method (Visio)

Returns the formulas of many cells.


## Syntax

 _expression_ . **GetFormulasU**( **_SRCStream()_** , **_formulaArray()_** )

 _expression_ A variable that represents a **Style** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SRCStream()_|Required| **Integer**|A stream identifying cells to be queried.|
| _formulaArray()_|Required| **Variant**|Out parameter. An array that receives formulas of queried cells.|

### Return Value

Nothing


## Remarks

The  **GetFormulasU** method is like the **FormulaU** property of a **Cell** object, except you can use it to obtain the formulas of many cells at once rather than one cell at a time. The **GetFormulasU** method is a specialization of the **GetResults** method, which can be used to obtain cell formulas or results. Setting up a call to the **GetFormulasU** method involves slightly less work than setting up the **GetResults** method.

You can use the  **GetFormulasU** method to get formulas of any set of cells.

 _SRCStream()_ is an array of 2-byte integers. For **Style** objects, _SRCStream()_ should be a one-dimensional array of 3 _n_ 2-byte integers for some _n_ >= 1. **GetFormulasU** interprets the stream as:




```
{sectionIdx, rowIdx, cellIdx}n
```

where  _sectionIdx_ is the section index of the desired cell, _rowIdx_ is its row index and _cellIdx_ is its cell index.

If the  **GetFormulasU** method succeeds, _formulaArray()_ returns a one-dimensional array of _n_ variants indexed from 0 to _n_ - 1. Each variant returns a formula as a string. _formulaArray()_ is an out parameter that is allocated by the **GetFormulasU** method, which passes ownership back to the caller. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. Note that the **SafeArrayDestroy** procedure has the side effect of clearing the variants referenced by the array's entries, hence deallocating any strings the **GetFormulas** method returns. (Microsoft Visual Basic and Microsoft Visual Basic for Applications take care of this for you.) The **GetFormulasU** method fails if _formulaArray()_ is **Null** .


 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **GetFormulas** method to get more than one formula when you are using local syntax. Use the **GetFormulasU** method to get more than one formula when you are using universal syntax.


