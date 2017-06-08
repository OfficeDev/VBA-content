---
title: Key Property
keywords: vblr6.chm2181947
f1_keywords:
- vblr6.chm2181947
ms.prod: office
api_name:
- Office.Key
ms.assetid: 6b2d19f0-9729-7c36-fc22-bde7d6366fc8
ms.date: 06/08/2017
---


# Key Property



 **Description**
Sets a  _key_ in a **Dictionary** object.
 **Syntax**
 _object_. **Key(**_key_**)** = _newkey_
The  **Key** property has the following parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **Dictionary** object.|
| _key_|Required.  _Key_ value being changed.|
| _newkey_|Required. New value that replaces the specified  _key_.|
 **Remarks**
If  _key_ is not found when changing a _key_, a[run-time error](vbe-glossary.md) will occur.

