---
title: Like Operator
keywords: vblr6.chm1008961
f1_keywords:
- vblr6.chm1008961
ms.prod: office
ms.assetid: 6df80925-8331-6c8c-4fd3-f397de0e44c1
ms.date: 06/08/2017
---


# Like Operator



Used to compare two strings.

**Syntax**

_result_ **=** _string_ **Like** _pattern_

The  **Like** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _string_|Required; any [string expression](vbe-glossary.md).|
| _pattern_|Required; any string expression conforming to the pattern-matching conventions described in Remarks.|

**Remarks**

If  _string_ matches _pattern_, _result_ is **True**; if there is no match, _result_ is **False**. If either _string_ or _pattern_ is[Null](vbe-glossary.md),  _result_ is **Null**.  
The behavior of the  **Like** operator depends on the **Option Compare** statement. The default[string-comparison](vbe-glossary.md) method for each[module](vbe-glossary.md) is **Option Compare** **Binary**.  

**Option Compare Binary** results in string comparisons based on a[sort order](vbe-glossary.md) derived from the internal binary representations of the characters. Sort order is determined by the code page. In the following example, a typical binary sort order is shown:  
A < B < E < Z < a < b < e < z < À < Ê < Ø < à < ê < ø
 
**Option Compare Text** results in string comparisons based on a case-insensitive, textual sort order determined by your system's[locale](vbe-glossary.md). When you sort the same characters using  **Option Compare Text**, the following text sort order is produced:  
(A=a) < (À=à) < (B=b) < (E=e) < (Ê=ê) < (Z=z) < (Ø=ø)

Built-in pattern matching provides a versatile tool for string comparisons. The pattern-matching features allow you to use wildcard characters, character lists, or character ranges, in any combination, to match strings. The following table shows the characters allowed in  _pattern_ and what they match:


|**Characters in  _pattern_**|**Matches in  _string_**|
|:-----|:-----|
|**?**|Any single character.|
|**\***|Zero or more characters.|
|**#**|Any single digit (0-9).|
|[ _charlist_ ]|Any single character in  _charlist_.|
|[ **!**_charlist_ ]|Any single character not in  _charlist_.|

A group of one or more characters ( _charlist_ ) enclosed in brackets ( **[ ]** ) can be used to match any single character in _string_ and can include almost any[character code](vbe-glossary.md), including digits.

 **Note**  To match the special characters left bracket ( **[** ), question mark ( **?** ), number sign ( **#** ), and asterisk ( **\*** ), enclose them in brackets. The right bracket ( **]** ) can't be used within a group to match itself, but it can be used outside a group as an individual character.

By using a hyphen ( **-** ) to separate the upper and lower bounds of the range, _charlist_ can specify a range of characters. For example, `[A-Z]` results in a match if the corresponding character position in _string_ contains any uppercase letters in the range A-Z. Multiple ranges are included within the brackets without delimiters.
The meaning of a specified range depends on the character ordering valid at [run time](vbe-glossary.md) (as determined by **Option Compare** and the[locale](vbe-glossary.md) setting of the system the code is running on). Using the **Option Compare Binary** example, the range `[A-E]` matches A, B and E. With **Option Compare Text**, `[A-E]` matches A, a, À, à, B, b, E, e. The range does not match Ê or ê because accented characters fall after unaccented characters in the sort order.
Other important rules for pattern matching include the following:

- An exclamation point ( **!** ) at the beginning of _charlist_ means that a match is made if any character except the characters in _charlist_ is found in _string_. When used outside brackets, the exclamation point matches itself.
    
- A hyphen ( **-** ) can appear either at the beginning (after an exclamation point if one is used) or at the end of _charlist_ to match itself. In any other location, the hyphen is used to identify a range of characters.
    
- When a range of characters is specified, they must appear in ascending sort order (from lowest to highest).  `[A-Z]` is a valid pattern, but `[Z-A]` is not.
    
- The character sequence  `[]` is considered a zero-length string ("").
    

In some languages, there are special characters in the alphabet that represent two separate characters. For example, several languages use the character "æ" to represent the characters "a" and "e" when they appear together. The  **Like** operator recognizes that the single special character and the two individual characters are equivalent.
When a language that uses a special character is specified in the system locale settings, an occurrence of the single special character in either  _pattern_ or _string_ matches the equivalent 2-character sequence in the other string. Similarly, a single special character in _pattern_ enclosed in brackets (by itself, in a list, or in a range) matches the equivalent 2-character sequence in _string_.

## Example

This example uses the  **Like** operator to compare a string to a pattern.


```vb
Dim MyCheck
MyCheck = "aBBBa" Like "a*a"    ' Returns True.
MyCheck = "F" Like "[A-Z]"    ' Returns True.
MyCheck = "F" Like "[!A-Z]"    ' Returns False.
MyCheck = "a2a" Like "a#a"    ' Returns True.
MyCheck = "aM5b" Like "a[L-P]#[!c-e]"    ' Returns True.
MyCheck = "BAT123khg" Like "B?T*"    ' Returns True.
MyCheck = "CAT123khg" Like "B?T*"    ' Returns False.
MyCheck = "ab" Like "a*b"    ' Returns True.
MyCheck = "a*b" Like "a[*]b"    ' Returns True.
MyCheck = "axxxxxb" Like "a[*]b"    ' Returns False.
MyCheck = "a[xyz" Like "a[[]*"    ' Returns True.
MyCheck = "a[xyz" Like "a[*"    ' Throws Error 93 (invalid pattern string).
```


