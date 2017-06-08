---
title: GroupCriterion2.FontColorEx Property (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion2.FontColorEx
ms.assetid: 5556ce54-315e-f676-6771-baec87853d8a
ms.date: 06/08/2017
---


# GroupCriterion2.FontColorEx Property (Project)

Gets or sets the color of the font for a field used as a criterion in a group definition. Read/write  **Long**.


## Syntax

 _expression_. **FontColorEx**

 _expression_ An expression that returns a **GroupCriterion2** object.


## Remarks

RGB colors can be expressed in decimal or hexadecimal values. In Project, red is the last byte of a hexadecimal value. For example, if the value of CellColorEx is 65535, the color is blue (&;HFF0000). 

The valid range for a normal RGB color is 0 to 16,777,215 (&;HFFFFFF&;). Each color setting (property or argument) is a 4-byte integer. The high byte of a number in this range equals 0. The lower 3 bytes, from least to most significant byte, determine the amount of red, green, and blue, respectively. The red, green, and blue components are each represented by a number between 0 and 255 (&;HFF). 


## See also


#### Concepts


[GroupCriterion2 Object](groupcriterion2-object-project.md)

