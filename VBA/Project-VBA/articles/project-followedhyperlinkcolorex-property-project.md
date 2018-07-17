---
title: Project.FollowedHyperlinkColorEx Property (Project)
keywords: vbapj.chm132288
f1_keywords:
- vbapj.chm132288
ms.prod: project-server
api_name:
- Project.Project.FollowedHyperlinkColorEx
ms.assetid: 72683515-81d3-915b-6da0-2593fbca0d00
ms.date: 06/08/2017
---


# Project.FollowedHyperlinkColorEx Property (Project)

Gets or sets the color used to denote followed hyperlinks. Read/write  **Long**.


## Syntax

 _expression_. **FollowedHyperlinkColorEx**

 _expression_ An expression that returns a **Project** object.


## Remarks

RGB colors can be expressed in decimal or hexadecimal values. In Project, red is the last byte of a hexadecimal value. For example, if the value of CellColorEx is 65535, the color is blue (&;HFF0000). 

The valid range for a normal RGB color is 0 to 16,777,215 (&;HFFFFFF&;). Each color setting (property or argument) is a 4-byte integer. The high byte of a number in this range equals 0. The lower 3 bytes, from least to most significant byte, determine the amount of red, green, and blue, respectively. The red, green, and blue components are each represented by a number between 0 and 255 (&;HFF). 


