---
title: Application.CommandBars Property (Publisher)
keywords: vbapb10.chm131088
f1_keywords:
- vbapb10.chm131088
ms.prod: publisher
api_name:
- Publisher.Application.CommandBars
ms.assetid: 21537c04-d406-6016-4f35-2f6ce6851db2
ms.date: 06/08/2017
---


# Application.CommandBars Property (Publisher)

Sets or returns a  **CommandBars** collection that represents the menu bar and all the toolbars in Microsoft Publisher.


## Syntax

 _expression_. **CommandBars**

 _expression_A variable that represents a  **Application** object.


### Return Value

CommandBars


## Example

This example enlarges all command bar buttons, enables ToolTips, and shows all menu items when displaying menus.


```vb
Sub CmdBars() 
 
 With CommandBars 
 .LargeButtons = False 
 .DisplayTooltips = True 
 .AdaptiveMenus = False 
 End With 
 
End Sub
```

This example displays the  **Objects** toolbar at the bottom of the application window.




```vb
Sub ShowObjectsToolbar 
 
 With CommandBars("Objects") 
 .Visible = True 
 .Position = msoBarBottom 
 End With 
 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

