---
title: CalloutFormat.Border Property (Word)
keywords: vbawd10.chm163905640
f1_keywords:
- vbawd10.chm163905640
ms.prod: word
api_name:
- Word.CalloutFormat.Border
ms.assetid: 4928f59e-1a09-32b9-0e73-ac7f9fbbb047
ms.date: 06/08/2017
---


# CalloutFormat.Border Property (Word)

Returns or sets whether the text in the specified callout is surrounded by a border. Read/write  **MsoTriState** .


## Syntax

 _expression_ . **Border**

 _expression_ Required. A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


## Example

This example adds an oval to the active document and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes 
 .AddShape msoShapeOval, 180, 200, 280, 130 
 With .AddCallout(msoCalloutTwo, 420, 170, 170, 40) 
 .TextFrame.TextRange.Text = "My oval" 
 With .Callout 
 .Accent = True 
 .Border = False 
 End With 
 End With 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

