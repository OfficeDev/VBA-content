---
title: WebOptionButton Object (Publisher)
keywords: vbapb10.chm4325375
f1_keywords:
- vbapb10.chm4325375
ms.prod: publisher
api_name:
- Publisher.WebOptionButton
ms.assetid: acdbaebd-b333-02b1-bf4d-d7e92148a275
ms.date: 06/08/2017
---


# WebOptionButton Object (Publisher)

Represents a Web option button control. The  **WebOptionButton** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** method to create new Web option button. Use the **[WebOptionButton](shape-weboptionbutton-property-publisher.md)** property to access a Web option button control shape. This example creates a new Web option button and specifies that its default state is selected; then it adds a text box next to it to describe it.
 

 

```
Sub CreateNewWebOptionButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlOptionButton, Left:=100, _ 
 Top:=123, Width:=16, Height:=10).WebOptionButton 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=120, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Advanced User" 
 End With 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](weboptionbutton-application-property-publisher.md)|
|[Parent](weboptionbutton-parent-property-publisher.md)|
|[ReturnDataLabel](weboptionbutton-returndatalabel-property-publisher.md)|
|[Selected](weboptionbutton-selected-property-publisher.md)|
|[Value](weboptionbutton-value-property-publisher.md)|

