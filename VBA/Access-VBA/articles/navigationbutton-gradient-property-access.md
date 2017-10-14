---
title: NavigationButton.Gradient Property (Access)
keywords: vbaac10.chm14662
f1_keywords:
- vbaac10.chm14662
ms.prod: access
api_name:
- Access.NavigationButton.Gradient
ms.assetid: b23fb655-67bf-645f-f510-c4baafe02e58
ms.date: 06/08/2017
---


# NavigationButton.Gradient Property (Access)

Gets or sets the gradient fill applied to the specified object. Read/write  **Long**.


## Syntax

 _expression_. **Gradient**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

The  **Gradient** property contains a numeric expression that represents the gradient fill applied to the specified object. The default value of the **Gradient** property is 0, which does not apply a gradient. The values are listed in the following table.



|**Value**|**Gradient Fill**|
|:-----|:-----|
|0|None|
|1|Linear Diagonal - Top Left to Bottom Right, Light|
|2|Linear Down, Light|
|3|Linear Diagonal - Top Right to Bottom Left, Light|
|4|From Bottom Right Corner, Light|
|5|From Bottom Left Corner, Light|
|6|Linear Left, Light|
|7|Linear Center, Light|
|8|Linear Right, Light|
|9|From Top Right Corner, Light|
|10|From Top Left Corner, Light|
|11|Linear Diagonal - Bottom Left to Top Right, Light|
|12|Linear Up, Light|
|13|Linear Diagonal - Bottom Right to Top Left, Light|
|14|Linear Diagonal - Top Left to Bottom Right, Dark|
|15|Linear Down, Dark|
|16|Linear Diagonal - Top Right to Bottom Left, Dark|
|17|From Bottom Right Corner, Dark|
|18|From Bottom Left Corner, Dark|
|19|Linear Left, Dark|
|20|Linear Center, Dark|
|21|Linear Right, Dark|
|22|From Top Right Corner, Dark|
|23|From Top Left Corner, Dark|
|24|Linear Diagonal - Bottom Left to Top Right, Dark|
|25|Linear Up, Dark|
|26|Linear Diagonal - Bottom Right to Top Left, Light|
To see the available gradient fills and apply a gradient through the user interface, first open the form or report in Layout view or Design view by right-clicking the form or report in the Navigation Pane, and then clicking the view you want. Then, click the object to which you want to apply a gradient fill. Next, on the  **Format** tab, in the **Control Formatting** group, click **Shape Fill**, then click  **Gradient** and choose a gradient fill. You can hover over a gradient fill to see a description.

This property is not surfaced in the property sheet.


## Example

The following code example sets the gradient fill to Linear Down, Light.


```vb
Me.ctl.Gradient = 2
```


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

