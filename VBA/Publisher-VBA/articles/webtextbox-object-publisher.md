---
title: WebTextBox Object (Publisher)
keywords: vbapb10.chm4259839
f1_keywords:
- vbapb10.chm4259839
ms.prod: publisher
api_name:
- Publisher.WebTextBox
ms.assetid: 74fde391-734c-6672-dadb-59bc58232c0f
ms.date: 06/08/2017
---


# WebTextBox Object (Publisher)

Represents a Web text box control. The  **WebTextBox** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** method to create new Web option button. Use the **[WebTextBox](shape-webtextbox-property-publisher.md)** property to access a Web text box control shape. This example creates a new Web text box, specifies default text, indicates that entry is required, and limits entry to 50 characters.
 

 

```
Sub CreateWebTextBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=150, Height:=15).WebTextBox 
 .DefaultText = "Please Enter Your Full Name" 
 .RequiredControl = msoTrue 
 .Limit = 50 
 End With 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](webtextbox-application-property-publisher.md)|
|[DefaultText](webtextbox-defaulttext-property-publisher.md)|
|[EchoAsterisks](webtextbox-echoasterisks-property-publisher.md)|
|[Limit](webtextbox-limit-property-publisher.md)|
|[Parent](webtextbox-parent-property-publisher.md)|
|[RequiredControl](webtextbox-requiredcontrol-property-publisher.md)|
|[ReturnDataLabel](webtextbox-returndatalabel-property-publisher.md)|

