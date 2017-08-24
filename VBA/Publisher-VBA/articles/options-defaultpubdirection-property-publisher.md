---
title: Options.DefaultPubDirection Property (Publisher)
keywords: vbapb10.chm1048624
f1_keywords:
- vbapb10.chm1048624
ms.prod: publisher
api_name:
- Publisher.Options.DefaultPubDirection
ms.assetid: 628352c1-040f-9ab1-d0f1-308b2c26679c
ms.date: 06/08/2017
---


# Options.DefaultPubDirection Property (Publisher)

Returns or sets a  **PbDirectionType** constant that represents the default direction in which text flows when a new publication is created. Read/write.


## Syntax

 _expression_. **DefaultPubDirection**

 _expression_A variable that represents a  **Options** object.


### Return Value

PbDirectionType


## Remarks

The  **DefaultPubDirection** property value can be one of the **[PbDirectionType](pbdirectiontype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

This property generates an error if you are not running a bi-directional-enabled version of Microsoft Publisher (for example, Arabic).


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


