---
title: AnimationBehavior.CommandEffect Property (PowerPoint)
keywords: vbapp10.chm657013
f1_keywords:
- vbapp10.chm657013
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.CommandEffect
ms.assetid: e457389c-402f-43e2-fbda-fdc286378501
ms.date: 06/08/2017
---


# AnimationBehavior.CommandEffect Property (PowerPoint)

Returns a  **CommandEffect** object for the specified animation behavior. Read-only.


## Syntax

 _expression_. **CommandEffect**

 _expression_ A variable that represents a **AnimationBehavior** object.


### Return Value

CommandEffect


## Remarks

You can send events, call functions, and send OLE verbs to embedded objects using this property.


## Example

The following example shows how to set a command effect animation behavior.


```vb
    Set bhvEffect = effectNew.Behaviors.Add(msoAnimTypeCommand)

 

    With bhvEffect.CommandEffect

         .Type = msoAnimCommandTypeVerb

         .Command = Play

    End With
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

