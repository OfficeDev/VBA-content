---
title: CommandEffect.Command Property (PowerPoint)
keywords: vbapp10.chm668004
f1_keywords:
- vbapp10.chm668004
ms.prod: powerpoint
api_name:
- PowerPoint.CommandEffect.Command
ms.assetid: 64440745-d84a-f0e8-6857-ca0f7ada42b6
ms.date: 06/08/2017
---


# CommandEffect.Command Property (PowerPoint)

Represents the command to be executed for the command effect. Read/write.


## Syntax

 _expression_. **Command**

 _expression_ A variable that represents a **CommandEffect** object.


### Return Value

String


## Remarks

You can send OLE verbs to embedded objects using this property.

If the shape is an OLE object, then the ole object will execute the command if it understands the verb.

If the shape is a media object (sound/video), Microsoft PowerPoint understands the following verbs: play, stop, pause, togglepause, resume and playfrom. Any other command sent to the shape will be ignored.


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


[CommandEffect Object](commandeffect-object-powerpoint.md)

