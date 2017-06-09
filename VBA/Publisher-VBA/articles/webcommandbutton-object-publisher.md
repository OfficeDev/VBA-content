---
title: WebCommandButton Object (Publisher)
keywords: vbapb10.chm3997695
f1_keywords:
- vbapb10.chm3997695
ms.prod: publisher
api_name:
- Publisher.WebCommandButton
ms.assetid: 86605945-eca1-ab80-1a1a-f8a5977d9282
ms.date: 06/08/2017
---


# WebCommandButton Object (Publisher)

Represents a Web command button control. The  **WebCommandButton** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** method to create new Web command button. Use the **[WebCommandButton](shape-webcommandbutton-property-publisher.md)** property to access a Web command button control shape. This example creates a Web form Submit command button and sets the script path and file name to run when a user clicks the button.
 

 

```
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" _ 
 &amp; "scripts/ispscript.cgi" 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[ActionURL](webcommandbutton-actionurl-property-publisher.md)|
|[Application](webcommandbutton-application-property-publisher.md)|
|[ButtonText](webcommandbutton-buttontext-property-publisher.md)|
|[ButtonType](webcommandbutton-buttontype-property-publisher.md)|
|[DataFileFormat](webcommandbutton-datafileformat-property-publisher.md)|
|[DataFileName](webcommandbutton-datafilename-property-publisher.md)|
|[DataRetrievalMethod](webcommandbutton-dataretrievalmethod-property-publisher.md)|
|[EmailAddress](webcommandbutton-emailaddress-property-publisher.md)|
|[EmailSubject](webcommandbutton-emailsubject-property-publisher.md)|
|[HiddenFields](webcommandbutton-hiddenfields-property-publisher.md)|
|[Parent](webcommandbutton-parent-property-publisher.md)|
|[PostFormData](webcommandbutton-postformdata-property-publisher.md)|

