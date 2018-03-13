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


| <strong>Part</strong> | <strong>Description</strong>                                        |
|:----------------------|:--------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>Dictionary</strong> object. |
| <em>key</em>          | Required.  <em>Key</em> value being changed.                        |
| <em>newkey</em>       | Required. New value that replaces the specified  <em>key</em>.      |

 **Remarks**
If  _key_ is not found when changing a _key_, a[run-time error](vbe-glossary.md) will occur.

