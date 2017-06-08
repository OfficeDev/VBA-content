---
title: FileName Property
keywords: vbob6.chm102003
f1_keywords:
- vbob6.chm102003
ms.prod: office
api_name:
- Office.FileName
ms.assetid: 626a11bf-894c-1831-657c-44d34311afd1
ms.date: 06/08/2017
---


# FileName Property



Returns the full path name of the project file or host document.
 **Syntax**
 _object_**.Filename**
The  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.
 **Remarks**
Projects have no name other than the file name.
The path name returned is always provided as an absolute path (for example, "c:\projects\myproject.vba"), even if it is shown as a relative path (such as "..\projects\myproject.vba").

