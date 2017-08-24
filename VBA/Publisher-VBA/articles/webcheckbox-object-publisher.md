---
title: WebCheckBox Object (Publisher)
keywords: vbapb10.chm4390911
f1_keywords:
- vbapb10.chm4390911
ms.prod: publisher
api_name:
- Publisher.WebCheckBox
ms.assetid: adcdf233-50b8-acbe-e52f-1e86e175b31d
ms.date: 06/08/2017
---


# WebCheckBox Object (Publisher)

Represents a Web check box control. The  **WebCheckBox** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** method to create a Web check box. Use the **[WebCheckBox](shape-webcheckbox-property-publisher.md)** property to access a Web check box control shape. This example creates a new Web check box and specifies that its default state is checked; then it adds a text box next to it to describe it.
 

 

```
Sub CreateNewWebCheckBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=123, Width:=17, Height:=12).WebCheckBox 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=118, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Description text for Web check box" 
 End With 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](webcheckbox-application-property-publisher.md)|
|[Parent](webcheckbox-parent-property-publisher.md)|
|[ReturnDataLabel](webcheckbox-returndatalabel-property-publisher.md)|
|[Selected](webcheckbox-selected-property-publisher.md)|
|[Value](webcheckbox-value-property-publisher.md)|

