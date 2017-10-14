---
title: ToggleButton.SoftEdges Property (Access)
keywords: vbaac10.chm14639
f1_keywords:
- vbaac10.chm14639
ms.prod: access
api_name:
- Access.ToggleButton.SoftEdges
ms.assetid: 23c63821-966c-4d9f-7304-5b6e31b85675
ms.date: 06/08/2017
---


# ToggleButton.SoftEdges Property (Access)

Gets or sets the  **Soft Edges** effect applied to the specified object. Read/write **Long**.


## Syntax

 _expression_. **SoftEdges**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **SoftEdges** property uses one of the values listed in the following table.



|**Value**|**Effect**|
|:-----|:-----|
|0 (Default)|No Soft Edges|
|1|1 Point|
|2|2.5 Points|
|3|5 Points|
|4|10 Points|
|5|25 Points|
|6|50 Points|
To see the available shadow effects and apply a shadow through the user interface, first open the form or report in Layout view or Design view by right-clicking the form or report in the Navigation Pane, and then clicking the view you want. Then, click the object to which you want to apply a soft edge effect. Next, on the  **Format** tab, in the **Control Formatting** group, click **Shape Effects**, then click  **Soft Edges** and choose a soft edge effect. Notice that the shadow effects are indexed from top to bottom.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

