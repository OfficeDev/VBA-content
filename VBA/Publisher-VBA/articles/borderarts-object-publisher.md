---
title: BorderArts Object (Publisher)
keywords: vbapb10.chm7798783
f1_keywords:
- vbapb10.chm7798783
ms.prod: publisher
api_name:
- Publisher.BorderArts
ms.assetid: 0fc016f6-154e-3591-34b3-e094bbad9d16
ms.date: 06/08/2017
---


# BorderArts Object (Publisher)

A collection of all BorderArt available for use in the specified publication. BorderArt is predefined picture borders that can be applied to text boxes, picture frames, or rectangles.
 


## Remarks

The  **BorderArts** collection includes any custom BorderArt types created by the user for the specified publication.
 

 

## Example

Use the  **[Item](borderarts-item-method-publisher.md)** property of a **BorderArts** collection to return a specific **[BorderArt](borderart-object-publisher.md)** object. The Index argument of the **Item** property can be the number or name of the BorderArt object.
 

 
This example returns the BorderArt "Apples" from the active publication. 
 

 



```
Dim bdaTemp As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts.Item (Index:="Apples") 
```

Use the  **[Count](borderarts-count-property-publisher.md)** property to return the number of BorderArt types available in the specified document. The following example displays the number of BorderArt types in the active document.
 

 



```
Sub CountBorderArts() 
 MsgBox ActiveDocument.BorderArts.Count 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](borderarts-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](borderarts-application-property-publisher.md)|
|[Count](borderarts-count-property-publisher.md)|
|[Parent](borderarts-parent-property-publisher.md)|

