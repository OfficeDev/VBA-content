---
title: Application.CommandBars Property (Word)
keywords: vbawd10.chm158335033
f1_keywords:
- vbawd10.chm158335033
ms.prod: word
api_name:
- Word.Application.CommandBars
ms.assetid: 1082697d-edc8-c619-40d1-466d2ebf3817
ms.date: 06/08/2017
---


# Application.CommandBars Property (Word)

Returns a  **CommandBars** collection that represents the menu bar and all the toolbars in Microsoft Word.


## Syntax

 _expression_ . **CommandBars**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Use the  **[CustomizationContext](application-customizationcontext-property-word.md)** property to set the template or document context prior to accessing the **CommandBars** collection.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example enlarges all command bar buttons and enables ToolTips.


```vb
With CommandBars 
 .LargeButtons = True 
 .DisplayTooltips = True 
End With
```

This example displays the  **Drawing** toolbar at the bottom of the application window.




```vb
With CommandBars("Drawing") 
 .Visible = True 
 .Position = msoBarBottom 
End With
```

This example adds the Versions command button to the  **Standard** toolbar.




```
CustomizationContext = NormalTemplate 
CommandBars("Standard").Controls.Add Type:=msoControlButton, _ 
 ID:=2522, Before:=4
```


## See also


#### Concepts


[Application Object](application-object-word.md)

