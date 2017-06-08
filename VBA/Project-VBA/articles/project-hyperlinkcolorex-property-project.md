---
title: Project.HyperlinkColorEx Property (Project)
ms.prod: project-server
api_name:
- Project.Project.HyperlinkColorEx
ms.assetid: ee305b13-9375-47d4-4cae-c81af86f3606
ms.date: 06/08/2017
---


# Project.HyperlinkColorEx Property (Project)

Gets or sets a hexadecimal representation of the color used to denote unfollowed hyperlinks. Read/write  **Long**.


## Syntax

 _expression_. **HyperlinkColorEx**

 _expression_ An expression that returns a **Project** object.


## Remarks

RGB colors can be expressed in decimal or hexadecimal values. In Project, red is the last byte of a hexadecimal value. For example, if the value of CellColorEx is 65535, the color is blue (&;HFF0000). 

The valid range for a normal RGB color is 0 to 16,777,215 (&;HFFFFFF&;). Each color setting (property or argument) is a 4-byte integer. The high byte of a number in this range equals 0. The lower 3 bytes, from least to most significant byte, determine the amount of red, green, and blue, respectively. The red, green, and blue components are each represented by a number between 0 and 255 (&;HFF). 


