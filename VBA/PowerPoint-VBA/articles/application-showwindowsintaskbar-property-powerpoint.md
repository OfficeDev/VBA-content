---
title: Application.ShowWindowsInTaskbar Property (PowerPoint)
keywords: vbapp10.chm502041
f1_keywords:
- vbapp10.chm502041
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ShowWindowsInTaskbar
ms.assetid: ad386fe5-9985-a1cc-cc52-1552bc12cad4
ms.date: 06/08/2017
---


# Application.ShowWindowsInTaskbar Property (PowerPoint)

Determines whether there is a separate Windows taskbar button for each open presentation. Read/write.


## Syntax

 _expression_. **ShowWindowsInTaskbar**

 _expression_ A variable that represents a **Application** object.


### Return Value

MsoTriState


## Remarks

When set to  **True**, this property simulates the look of a single-document interface (SDI), which makes it easier to navigate between open presentations. However, if you work with multiple presentations while other applications are open, you may want to set this property to **False** to avoid filling your taskbar with unnecessary buttons.

The value of the  **ShowWindowsInTaskbar** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|There is not a separate Windows taskbar button for each open presentation.|
|**msoTrue**| The default. There is a separate Windows taskbar button for each open presentation.|

## Example

This example specifies that each open presentation doesn't have a separate Windows taskbar button.


```vb
Application.ShowWindowsInTaskbar = msoFalse
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

