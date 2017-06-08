---
title: CalloutFormat.CustomLength Method (Word)
keywords: vbawd10.chm163905548
f1_keywords:
- vbawd10.chm163905548
ms.prod: word
api_name:
- Word.CalloutFormat.CustomLength
ms.assetid: b9c2a9d5-873e-9292-04e1-c2e05388589b
ms.date: 06/08/2017
---


# CalloutFormat.CustomLength Method (Word)

Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved.


## Syntax

 _expression_ . **CustomLength**( **_Length_** )

 _expression_ Required. A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Length_|Required| **Single**|The length of the first segment of the callout, in points.|

## Remarks

Use the  **AutomaticLength** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).

Applying this method sets the  **AutoLength** property to **False** and sets the **Length** property to the value specified for the Length argument.


## Example

This example switches between an automatically scaling first segment and one with a fixed length for the callout line for the first shape on the active document. For the example to work, the first shape must be a callout.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 

```


```vb
With docActive.Shapes(1).Callout 
 If .AutoLength Then 
 .CustomLength 50 
 Else 
 .AutomaticLength 
 End If 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

