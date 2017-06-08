---
title: Wildcard Characters used in String Comparisons
keywords: vbui6.chm1105375
f1_keywords:
- vbui6.chm1105375
ms.prod: office
ms.assetid: 351a346d-95ee-1801-1e59-fb17befdb65f
ms.date: 06/08/2017
---


# Wildcard Characters used in String Comparisons

Built-in pattern matching provides a versatile tool for making string comparisons. The following table shows the wildcard characters you can use with the Like operator and the number of digits or strings they match.



|**Character(s) in  _pattern_**|**Matches in  _expression_**|
|:-----|:-----|
|?|Any single character|
|*|Zero or more characters|
|#|Any single digit (09)|
|[ _charlist_ ]|Any single character in  _charlist_|
|[! _charlist_ ]|Any single character not in  _charlist_|

A group of one or more characters ( _charlist_ ) enclosed in brackets ([ ]) can be used to match any single character in _expression_ and can include almost any characters in the[ANSI](vbe-glossary.md) character set, including digits. In fact, the special characters opening bracket ([ ), question mark (?), number sign (#), and asterisk (*) can be used to match themselves directly only if enclosed in brackets. The closing bracket ( ]) can't be used within a group to match itself, but it can be used outside a group as an individual character.

In addition to a simple list of characters enclosed in brackets,  _charlist_ can specify a range of characters by using a hyphen (-) to separate the upper and lower bounds of the range. For example, using [A-Z] in _pattern_ results in a match if the corresponding character position in _expression_ contains any of the uppercase letters in the range A through Z. Multiple ranges can be included within the brackets without any delimiting. For example, [a-zA-Z0-9] matches any alphanumeric character.
Other important rules for pattern matching include the following:


- An exclamation mark (!) at the beginning of  _charlist_ means that a match is made if any character except those in _charlist_ are found in _expression_. When used outside brackets, the exclamation mark matches itself.
    
- The hyphen (-) can be used either at the beginning (after an exclamation mark if one is used) or at the end of  _charlist_ to match itself. In any other location, the hyphen is used to identify a range of ANSI characters.
    
- When a range of characters is specified, they must appear in ascending sort order (A-Z or 0-100). [A-Z] is a valid pattern, but [Z-A] isn't.
    
- The character sequence [ ] is ignored; it's considered to be a zero-length string ("").
    


