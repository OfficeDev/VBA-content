---
title: Enum types defined in standard modules or private classes cannot be used in public object modules as parameters or return types for public procedures, as public data members, or as fields of public user defined types
keywords: vblr6.chm1040368
f1_keywords:
- vblr6.chm1040368
ms.prod: office
ms.assetid: 73942c44-c1a2-e75a-d5ee-c1d6e4fd98d0
ms.date: 06/08/2017
---


# Enum types defined in standard modules or private classes cannot be used in public object modules as parameters or return types for public procedures, as public data members, or as fields of public user defined types

This error has the following cause and solution:



- A non-exposed enum was used as a parameter or return type of a public procedure or a public data member of an exposed class.
    

Exposed here means that the enum is exposed from the ActiveX server that is being defined, which is equivalent to saying that it is declared in a public class of an ActiveX Exe or Dll project.

