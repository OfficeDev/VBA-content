---
title: Presentation.Final Property (PowerPoint)
keywords: vbapp10.chm583104
f1_keywords:
- vbapp10.chm583104
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Final
ms.assetid: 03b16954-2f23-905b-8392-d88070e86e9f
ms.date: 06/08/2017
---


# Presentation.Final Property (PowerPoint)

Determines whether the presentation is marked as final (read-only). Read/write.


## Syntax

 _expression_. **Final**

 _expression_ An expression that returns a **Presentation** object.


### Return Value

Boolean


## Remarks

The setting of the  **Final** property corresponds to the status of the **Mark As Final** command in the PowerPoint user interface (click the **Office** button, and then point to **Prepare**). 

Marking a presentation as final makes the presentation read-only and prevents changes to the presentation. When a presentation is marked as final, typing, editing commands, and proofing marks are disabled or turned off and the presentation becomes read-only.

 Setting the **Final** property to **True** helps you communicate that you are sharing a completed version of a presentation. It also helps prevent reviewers or readers from making inadvertent changes to the presentation.


- The  **Final** property is not a security feature. Anyone who receives an electronic copy of a presentation that has been marked as final can edit that presentation by removing **Mark as Final** status from the presentation.
    
- Presentations that have been marked as final in a Office program will not be read-only if they are opened in earlier versions of Microsoft Office programs.
    

## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

