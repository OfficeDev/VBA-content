---
title: TextBox.KeyboardLanguage Property (Access)
keywords: vbaac10.chm11131
f1_keywords:
- vbaac10.chm11131
ms.prod: access
api_name:
- Access.TextBox.KeyboardLanguage
ms.assetid: a3b55e3e-16a9-87c7-6c03-bc8392e72c17
ms.date: 06/08/2017
---


# TextBox.KeyboardLanguage Property (Access)





## Syntax

 _expression_. **KeyboardLanguage**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

Valid values for this property are 0 (zero), which corresponds to the default system language, or  _plid_ + 2 where _plid_ is the primary language ID of a language installed on the current system. For example, the primary language ID of English is 9, so the corresponding **KeyboardLanguage** setting is 11. For a list of languages and their primary language IDs, search for "Primary Language IDs" in the MSDN Web site. (An exception to this list is Traditional Chinese which is represented by the value 200.)

Setting this property to a language that is not installed may either have no effect or cause an error.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

