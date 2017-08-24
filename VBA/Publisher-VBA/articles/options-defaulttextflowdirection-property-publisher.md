---
title: Options.DefaultTextFlowDirection Property (Publisher)
keywords: vbapb10.chm1048628
f1_keywords:
- vbapb10.chm1048628
ms.prod: publisher
api_name:
- Publisher.Options.DefaultTextFlowDirection
ms.assetid: 7c17768a-cd9c-704d-fa27-f0dfd7648054
ms.date: 06/08/2017
---


# Options.DefaultTextFlowDirection Property (Publisher)

Returns or sets a  **PbDirectionType** constant that represents a global Microsoft Publisher option, indicating whether text flows from left to right or from right to left in a publication. Read/write.


## Syntax

 _expression_. **DefaultTextFlowDirection**

 _expression_A variable that represents a  **Options** object.


### Return Value

PbDirectionType


## Remarks

The  **DefaultTextFlowDirection** property value can be one of the **[PbDirectionType](pbdirectiontype-enumeration-publisher.md)** constants declared in the Publisher type library.

This property generates an error if you are not running a bi-directional-enabled version of Publisher (for example, Arabic).


## Example

This example sets the default direction for new publications and text flow in a bi-directional-enabled version of Publisher.


```vb
Sub SetDefaultDirection() 
 With Options 
 .DefaultPubDirection = pbDirectionRightToLeft 
 .DefaultTextFlowDirection = pbDirectionRightToLeft 
 End With 
End Sub
```


