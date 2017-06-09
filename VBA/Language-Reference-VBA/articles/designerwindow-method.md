---
title: DesignerWindow Method
keywords: vbob6.chm102199
f1_keywords:
- vbob6.chm102199
ms.prod: office
api_name:
- Office.DesignerWindow
ms.assetid: 1a116dab-56ce-087e-1789-614a3709c9cc
ms.date: 06/08/2017
---


# DesignerWindow Method



Returns the  **Window** object that represents the component's[designer](vbe-glossary.md).
 **Syntax**
 _object_**.DesignerWindow**
The  _object_ placeholder is an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.
 **Remarks**
If the component supports a designer but doesn't have an open designer, using the  **DesignerWindow** method creates the designer, but it isn't visible. To make the window visible, set the **Window** object's **Visible** property to **True**.

