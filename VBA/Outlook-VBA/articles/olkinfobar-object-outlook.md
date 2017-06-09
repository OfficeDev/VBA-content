---
title: OlkInfoBar Object (Outlook)
keywords: vbaol11.chm1000304
f1_keywords:
- vbaol11.chm1000304
ms.prod: outlook
api_name:
- Outlook.OlkInfoBar
ms.assetid: 1aec19db-d28b-ef9b-3227-45aa4a296de6
ms.date: 06/08/2017
---


# OlkInfoBar Object (Outlook)

A control that provides an area to display specific information on a custom form.


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook InfoBar Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The following is an example of this control at runtime. This control supports Microsoft Windows themes.


![Information bar](images/olInfoBar_ZA10119648.gif)



If there is no information to display, this control will automatically resize to a height of zero.

You can specify only the placement of the control, as there are no configurable options or settings other than its position.

For more information about Outlook controls, see [Controls in a Custom Form](http://msdn.microsoft.com/library/fcba1b34-c526-5d01-8644-cb8852bd2348%28Office.15%29.aspx). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|**Name**|
|:-----|
|[Click](olkinfobar-click-event-outlook.md)|
|[DoubleClick](olkinfobar-doubleclick-event-outlook.md)|
|[MouseDown](olkinfobar-mousedown-event-outlook.md)|
|[MouseMove](olkinfobar-mousemove-event-outlook.md)|
|[MouseUp](olkinfobar-mouseup-event-outlook.md)|

## Properties



|**Name**|
|:-----|
|[MouseIcon](olkinfobar-mouseicon-property-outlook.md)|
|[MousePointer](olkinfobar-mousepointer-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
