---
title: ToggleButton.Shadow Property (Access)
keywords: vbaac10.chm14638
f1_keywords:
- vbaac10.chm14638
ms.prod: access
api_name:
- Access.ToggleButton.Shadow
ms.assetid: 0095ff4e-56f0-9b56-73e2-2e3066ee8b03
ms.date: 06/08/2017
---


# ToggleButton.Shadow Property (Access)

Gets or sets the  **Shadow** effect applied to the specified object. Read/write **Long**.


## Syntax

 _expression_. **Shadow**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **Shadow** property uses one of the values listed in the following table.



|**Value**|**Effect**|
|:-----|:-----|
|0 Default)|None|
|1|Offset Diagonal Bottom Right|
|2|Offset Bottom|
|3|Offset Diagonal Bottom Left|
|4|Offset Right|
|5|Offset Center|
|6|Offset Left|
|7|Offset Diagonal Top Right|
|8|Offset Top|
|9|Offset Diagonal Top Left|
|10|Inside Diagonal Top Left|
|11|Inside Top|
|12|Inside Diagonal Top Right|
|13|Inside Left|
|14|Inside Center|
|15|Inside Right|
|16|Inside Diagonal Bottom Left|
|17|Inside Bottom|
|18|Inside Diagonal Bottom Right|
|19|Perspective Diagonal Upper Left|
|20|Perspective Diagonal Upper Right|
|21|Below|
|22|Perspective Diagonal Lower Left|
|23|Perspective Diagonal Lower Right|
To see the available shadow effects and apply a shadow through the user interface, first open the form or report in Layout view or Design view by right-clicking the form or report in the Navigation Pane, and then clicking the view you want. Then, click the object to which you want to apply a shadow effect. Next, on the  **Format** tab, in the **Control Formatting** group, click **Shape Effects**, then click  **Shadow** and choose a shadow effect. Notice that the shadow effects are indexed from left to right, and then top to bottom. So the first item, under No Shadow, has the value 0. Then, under Outer, the first row contains shadow effects with values from 1 to 3. The second row from 4 to 6, and so on.

This property is not surfaced in the property sheet. 


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

