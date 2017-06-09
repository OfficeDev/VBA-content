---
title: Mid Function
keywords: vblr6.chm1011070
f1_keywords:
- vblr6.chm1011070
ms.prod: office
ms.assetid: 5d5e7712-459a-d504-dae6-4b52a9a90c6f
ms.date: 06/08/2017
---


# Mid Function



Returns a  **Variant** ( **String** ) containing a specified number of characters from a string.
 **Syntax**
 **Mid** ( **_string_**, **_start_** [, **_length_** ])
The  **Mid** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_string_**|Required. [String expression](vbe-glossary.md) from which characters are returned. If **_string_** contains[Null](vbe-glossary.md),  **Null** is returned.|
|**_start_**|Required; [Long](vbe-glossary.md). Character position in  **_string_** at which the part to be taken begins. If **_start_** is greater than the number of characters in **_string_**, **Mid** returns a zero-length string ("").|
|**_length_**|Optional;  **Variant** ( **Long** ). Number of characters to return. If omitted or if there are fewer than **_length_** characters in the text (including the character at **_start_** ), all characters from the **_start_** position to the end of the string are returned.|
 **Remarks**
To determine the number of characters in  **_string_**, use the **Len** function.

 **Note**  Use the  **MidB** function with byte data contained in a string, as in double-byte character set languages. Instead of specifying the number of characters, the[arguments](vbe-glossary.md) specify numbers of bytes. For sample code that uses **MidB**, see the second example in the example topic.


## Example

The first example uses the  **Mid** function to return a specified number of characters from a string.


```vb
Dim MyString, FirstWord, LastWord, MidWords
MyString = "Mid Function Demo"    ' Create text string.
FirstWord = Mid(MyString, 1, 3)    ' Returns "Mid".
LastWord = Mid(MyString, 14, 4)    ' Returns "Demo".
MidWords = Mid(MyString, 5)    ' Returns "Function Demo".

```

The second example use  **MidB** and a user-defined function ( **MidMbcs** ) to also return characters from string. The difference here is that the input string is ANSI and the length is in bytes.




```vb
Function MidMbcs(ByVal str as String, start, length)
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, length), vbUnicode)
End Function

Dim MyString
MyString = "AbCdEfG"
' Where "A", "C", "E", and "G" are DBCS and "b", "d", 
' and "f" are SBCS.
MyNewString = Mid(MyString, 3, 4)
' Returns ""CdEf"
MyNewString = MidB(MyString, 3, 4)
' Returns ""bC"
MyNewString = MidMbcs(MyString, 3, 4)
' Returns "bCd"


```


