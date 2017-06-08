---
title: CommandEffect Object (PowerPoint)
keywords: vbapp10.chm668000
f1_keywords:
- vbapp10.chm668000
ms.prod: powerpoint
api_name:
- PowerPoint.CommandEffect
ms.assetid: 2ae803e4-1c94-46d0-45ac-38a62dc15b00
ms.date: 06/08/2017
---


# CommandEffect Object (PowerPoint)

Represents a command effect for an animation behavior. You can send events, call functions, and send OLE verbs to embedded objects using this object.


## Remarks

Use the  **CommandEffect** property of the **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object to return a **CommandEffect** object. Command effects can be changed using the **CommandEffect** object's **Command** and **Type** properties.


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


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

