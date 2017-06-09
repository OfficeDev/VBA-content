---
title: ComboBox.KeyboardLanguage Property (Access)
keywords: vbaac10.chm11464
f1_keywords:
- vbaac10.chm11464
ms.prod: access
api_name:
- Access.ComboBox.KeyboardLanguage
ms.assetid: 5eb0e03c-c931-45b5-7801-d790c4678768
ms.date: 06/08/2017
---


# ComboBox.KeyboardLanguage Property (Access)





## Syntax

 _expression_. **KeyboardLanguage**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

Valid values for this property are 0 (zero), which corresponds to the default system language, or  _plid_ + 2 where _plid_ is the primary language ID of a language installed on the current system. For example, the primary language ID of English is 9, so the corresponding **KeyboardLanguage** setting is 11. For a list of languages and their primary language IDs, search for "Primary Language IDs" in the MSDN Web site. (An exception to this list is Traditional Chinese which is represented by the value 200.)

Setting this property to a language that is not installed may either have no effect or cause an error.


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

