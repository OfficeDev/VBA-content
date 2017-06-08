---
title: Print Method
keywords: vblr6.chm1010081
f1_keywords:
- vblr6.chm1010081
ms.prod: office
api_name:
- Office.Print
ms.assetid: 489447fa-e0ea-404a-10f2-23dcd9a8e41a
ms.date: 06/08/2017
---


# Print Method



Prints text in the  **Immediate** window.
 **Syntax**
 _object_**.Print** [ _outputlist_ ]
The  **Print** method syntax has the following object qualifier and part:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Optional. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _outputlist_|Optional. [Expression](vbe-glossary.md) or list of expressions to print. If omitted, a blank line is printed.|
The  _outputlist_[argument](vbe-glossary.md) has the following syntax and parts:
{ **Spc(**_n_**)** |**Tab(**_n_**)** } _expression charpos_


|**Part**|**Description**|
|:-----|:-----|
|**Spc(**_n_**)**|Optional. Used to insert space characters in the output, where  _n_ is the number of space characters to insert.|
|**Tab(**_n_**)**|Optional. Used to position the insertion point at an absolute column number where  _n_ is the column number. Use **Tab** with no argument to position the insertion point at the beginning of the next[print zone](vbe-glossary.md).|
| _expression_|Optional. [Numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md) to print.|
| _charpos_|Optional. Specifies the insertion point for the next character. Use a semicolon ( **;** ) to position the insertion point immediately following the last character displayed. Use **Tab(**_n_**)** to position the insertion point at an absolute column number. Use **Tab** with no argument to position the insertion point at the beginning of the next print zone. If _charpos_ is omitted, the next character is printed on the next line.|
 **Remarks**
Multiple expressions can be separated with either a space or a semicolon.
All data printed to the  **Immediate** window is properly formatted using the decimal separator for the[locale](vbe-glossary.md) settings specified for your system. The[keywords](vbe-glossary.md) are output in the appropriate language for the[host application](vbe-glossary.md).
For [Boolean](vbe-glossary.md) data, either `True` or `False` is printed. The **True** and **False** keywords are translated according to the locale setting for the host application.
[Date](vbe-glossary.md) data is written using the standard short date format recognized by your system. When either the date or the time component is missing or zero, only the data provided is written.
Nothing is written if  _outputlist_ data is[Empty](vbe-glossary.md). However, if  _outputlist_ data is[Null](vbe-glossary.md),  `Null` is output. The **Null** keyword is appropriately translated when it is output.
For error data, the output is written as  `Error errorcode`. The  **Error** keyword is appropriately translated when it is output.
The  _object_ is required if the method is used outside a[module](vbe-glossary.md) having a default display space. For example an error occurs if the method is called in a[standard module](vbe-glossary.md) without specifying an _object_, but if called in a form module, _outputlist_ is displayed on the form.

 **Note**  Because the  **Print** method typically prints with proportionally-spaced characters, there is no correlation between the number of characters printed and the number of fixed-width columns those characters occupy. For example, a wide letter, such as a "W", occupies more than one fixed-width column, and a narrow letter, such as an "i", occupies less. To allow for cases where wider than average characters are used, your tabular columns must be positioned far enough apart. Alternatively, you can print using a fixed-pitch font (such as Courier) to ensure that each character uses only one column.


## Example

Using the  **Print** method, this example displays the value of the variable `MyVar` in the **Immediate** window. Note that the **Print** method only applies to objects that can display text.


```vb
Dim MyVar
MyVar = "Come see me in the Immediate pane."
Debug.Print MyVar

```


