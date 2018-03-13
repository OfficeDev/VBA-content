---
title: AutoWordSelect Property
keywords: fm20.chm2000760
f1_keywords:
- fm20.chm2000760
ms.prod: office
api_name:
- Office.AutoWordSelect
ms.assetid: 24e9e8ff-5988-9ed3-4a2c-f3faa99248f9
ms.date: 06/08/2017
---


# AutoWordSelect Property



Specifies whether a word or a character is the basic unit used to extend a selection.
 **Syntax**
 _object_. **AutoWordSelect** [= _Boolean_ ]
The  **AutoWordSelect** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                   |
|:----------------------|:---------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                      |
| <em>Boolean</em>      | Optional. Specifies the basic unit used to extend a selection. |

 **Settings**
The settings for  _Boolean_ are:


| <strong>Value</strong> | <strong>Description</strong>             |
|:-----------------------|:-----------------------------------------|
| <strong>True</strong>  | Uses a word as the basic unit (default). |
| <strong>False</strong> | Uses a character as the basic unit.      |

 **Remarks**
The  **AutoWordSelect** property specifies how the selection extends or contracts in the edit region of a **TextBox** or **ComboBox**.
If the user places the insertion point in the middle of a word and then extends the selection while  **AutoWordSelect** is **True**, the selection includes the entire word.

