---
title: Can't create AutoRedraw image (Error 480)
keywords: vblr6.chm480
f1_keywords:
- vblr6.chm480
ms.prod: office
ms.assetid: 35a64aeb-89ad-26fa-2a06-dbbf3d5457e4
ms.date: 06/08/2017
---


# Can't create AutoRedraw image (Error 480)

Visual Basic can't create a persistent bitmap for automatic redraw of the form or picture. This error has the following cause and solution:



- There isn't enough available memory for the  **AutoRedraw** property to be set to **True**. Set the **AutoRedraw** property to **False** and perform your own redraw in the Paint event procedure, or make the **PictureBox** control or **Form** object smaller and try the operation again with **AutoRedraw** set to **True**.
    


