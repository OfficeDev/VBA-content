---
title: CalloutFormat.Accent Property (Publisher)
keywords: vbapb10.chm2490624
f1_keywords:
- vbapb10.chm2490624
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Accent
ms.assetid: 8e31544c-79ed-3882-98d1-42fc88f58115
ms.date: 06/08/2017
---


# CalloutFormat.Accent Property (Publisher)

Returns or sets an  **MsoTriState**constant indicating whether a vertical accent bar separates the callout text from the callout line. Read/write.


## Syntax

 _expression_. **Accent**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

MsoTriState


## Remarks

The  **Accent** property value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoCTrue**|Not used with this property.|
| **msoFalse**|A vertical accent bar does not separate the callout text from the callout line.|
| **msoTriStateMixed**|Return value only; indicates a combination of  **msoTrue** and **msoFalse** in the specified shape range.|
| **msoTriStateToggle**|Set value only; switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|A vertical accent bar separates the callout text from the callout line.|

## Example

This example adds an oval to the active publication and a callout that points to the oval. The callout text will not have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
With ActiveDocument.Pages(1).Shapes 
 ' Add an oval. 
 .AddShape Type:=msoShapeOval, _ 
 Left:=180, Top:=200, Width:=280, Height:=130 
 
 ' Add a callout. 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=170, Width:=170, Height:=40) 
 
 ' Add text to the callout. 
 .TextFrame.TextRange.Text = "This is an oval" 
 
 ' Add an accent bar to the callout. 
 With .Callout 
 .Accent = msoTrue 
 .Border = msoFalse 
 End With 
 End With 
End With 

```


