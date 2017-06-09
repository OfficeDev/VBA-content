---
title: CommandButton.Glow Property (Access)
keywords: vbaac10.chm14630
f1_keywords:
- vbaac10.chm14630
ms.prod: access
api_name:
- Access.CommandButton.Glow
ms.assetid: e6c147b4-c378-90bd-7132-f44021994ecd
ms.date: 06/08/2017
---


# CommandButton.Glow Property (Access)

Gets or sets the Glow effect applied to the specified object. Read/write  **Long**.


## Syntax

 _expression_. **Glow**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **Glow** property contains a numeric expression that represents the glow effect applied to the specified object. The default value of the **Glow** property is 0, which does not apply a glow effect. The values are listed in the following table.



|**Value**|**Effect**|
|:-----|:-----|
|0|None|
|1|Blue, 5 pt glow, Accent color 1|
|2|Red, 5 pt glow, Accent color 2|
|3|Olive Green, 5 pt glow, Accent color 3|
|4|Purple, 5 pt glow, Accent color 4|
|5|Aqua, 5 pt glow, Accent color 5|
|6|Orange, 5 pt glow, Accent color 6|
|7|Blue, 8 pt glow, Accent color 1|
|8|Red, 8 pt glow, Accent color 2|
|9|Olive Green, 8 pt glow, Accent color 3|
|10|Purple, 8 pt glow, Accent color 4|
|11|Aqua, 8 pt glow, Accent color 5|
|12|Orange, 8 pt glow, Accent color 6|
|13|Blue, 11 pt glow, Accent color 1|
|14|Red, 11 pt glow, Accent color 2|
|15|Olive Green, 11 pt glow, Accent color 3|
|16|Purple, 11 pt glow, Accent color 4|
|17|Aqua, 11 pt glow, Accent color 5|
|18|Orange, 11 pt glow, Accent color 6|
|19|Blue, 18 pt glow, Accent color 1|
|20|Red, 18 pt glow, Accent color 2|
|21|Olive Green, 18 pt glow, Accent color 3|
|22|Purple, 18 pt glow, Accent color 4|
|23|Aqua, 18 pt glow, Accent color 5|
|24|Orange, 18 pt glow, Accent color 6|
To see the available glow effects and apply a glow through the user interface, first open the form or report in Layout view or Design view by right-clicking the form or report in the Navigation Pane, and then clicking the view you want. Then, click the object to which you want to apply a glow effect. Next, on the  **Format** tab, in the **Control Formatting** group, click **Shape Effects**, then click  **Glow** and choose a glow effect. Notice that the glow effects are indexed from left to right, and then top to bottom. So the first item, under No Glow, has the value 0. Then, under Glow, the first row contains glow effects with values from 1 to 6. The second row from 7 to 12, the third row from 13 to 18, and the fourth row from 19 to 24 .

This property is not surfaced in the property sheet.


## Example

The following code example sets the glow effect to Aqua, 11 pt glow, Accent color 5.


```vb
Me.ctl.Glow = 17
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

