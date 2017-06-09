---
title: PageBackground Object (Publisher)
keywords: vbapb10.chm8191999
f1_keywords:
- vbapb10.chm8191999
ms.prod: publisher
api_name:
- Publisher.PageBackground
ms.assetid: 647f5a84-0971-2f69-d281-c9ab402968a4
ms.date: 06/08/2017
---


# PageBackground Object (Publisher)

Represents the background of a page.
 


## Example

Use the  **Background** property of a **Page** object to return a **PageBackground** object. The following example creates a **PageBackground** object and sets it to the background of the first page of the active document.
 

 

```
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
 
```

Use  **PageBackground.Exists** to determine if a background already exists for the specified **Page** object. The following example builds upon the previous example. First a **PageBackground** object is created and set to the background of the first page of the active document. Then a test is made to check if a background exists for the page already. If not then one is created by calling the **Create** method of the **PageBackground** object.
 

 



```
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
If objPageBackground.Exists = False Then 
 objPageBackground.Create 
End If 
 
```

Use  **PageBackground.Fill** to return a **FillFormat** object. The following example builds upon the previous example. First a **PageBackground** object is created and set to the background of the first page of the active document. Then a test is made to check if a background exists for the page already. If not then one is created by calling the **Create** method of the **PageBackground** object. A **FillFormat** object is returned by using the **Fill** property of the **PageBackground** object. A few of the available properties of the **FillFormat** object are then set.
 

 



```
Dim objPageBackground As PageBackground 
Dim objFillFormat As FillFormat 
 
Set objPageBackground = ActiveDocument.Pages(1).Background 
If objPageBackground.Exists = False Then 
 objPageBackground.Create 
End If 
 
Set objFillFormat = objPageBackground.Fill 
With objFillFormat 
 .BackColor.RGB = RGB(Red:=0, GReen:=155, Blue:=99) 
 .ForeColor.RGB = RGB(Red:=155, GReen:=234, Blue:=0) 
 .TwoColorGradient msoGradientDiagonalDown, 4 
End With 
 
```

Use  **PageBackground.Delete** to delete a background for the specified page. The following example deletes the background of the first page in the active document. (The following example assumes the specified page has an existing background. A run-time error occurs if the page does not contain a background.)
 

 



```
ActiveDocument.Pages(1).Background.Delete
```


## Methods



|**Name**|
|:-----|
|[Create](pagebackground-create-method-publisher.md)|
|[Delete](pagebackground-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](pagebackground-application-property-publisher.md)|
|[Exists](pagebackground-exists-property-publisher.md)|
|[Fill](pagebackground-fill-property-publisher.md)|
|[Parent](pagebackground-parent-property-publisher.md)|

