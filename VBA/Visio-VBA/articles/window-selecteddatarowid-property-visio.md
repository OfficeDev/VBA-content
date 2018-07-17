---
title: Window.SelectedDataRowID Property (Visio)
keywords: vis_sdr.chm11660250
f1_keywords:
- vis_sdr.chm11660250
ms.prod: visio
api_name:
- Visio.Window.SelectedDataRowID
ms.assetid: 8ed4a690-c96f-c134-5b84-459938bd39e8
ms.date: 06/08/2017
---


# Window.SelectedDataRowID Property (Visio)

Gets or sets the ID of the data row that is selected (or that is the primary row selected, when multiple rows are selected) on the active tab of the  **External Data Window** in the Microsoft Visio user interface (UI). Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **SelectedDataRowID**

 _expression_ An expression that returns a **Window** object.


### Return Value

Long


## Remarks

The  **SelectedDataRowID** property works only when the **Window** object represents the **External Data Window**. Calling the property on any other window type results in an error. The  **External Data Window** must already be displayed in the Visio UI before you call **SelectedDataRowID** .

Setting  **SelectedDataRowID** clears the current selection and sets the selection to the row you specify.

If multiple rows are selected when you get  **SelectedDataRowID** , the property returns the ID of the row that has the focus.


