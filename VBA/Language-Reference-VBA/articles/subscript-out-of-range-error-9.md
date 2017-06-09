---
title: Subscript out of range (Error 9)
keywords: vblr6.chm1011240
f1_keywords:
- vblr6.chm1011240
ms.prod: office
ms.assetid: 37b59913-9318-35eb-0646-19cd72d4f459
ms.date: 06/08/2017
---


# Subscript out of range (Error 9)

Elements of [arrays](vbe-glossary.md) and members of[collections](vbe-glossary.md) can only be accessed within their defined ranges. This error has the following causes and solutions:



- You referenced a nonexistent array element. The subscript may be larger or smaller than the range of possible subscripts, or the array may not have dimensions assigned at this point in the application. Check the [declaration](vbe-glossary.md) of the array to verify its upper and lower bounds. Use the **UBound** and **LBound** functions to condition array accesses if you're working with arrays that are redimensioned. If the index is specified as a[variable](vbe-glossary.md), check the spelling of the variable name.
    
- You declared an array but didn't specify the number of elements. For example, the following code causes this error:
    
```vb
Dim MyArray() As Integer 
MyArray(8) = 234 ' Causes Error 9. 

  ```


    Visual Basic doesn't implicitly dimension unspecified array ranges as 0 - 10. Instead, you must use  **Dim** or **ReDim** to specify explicitly the number of elements in an array.
    
- You referenced a nonexistent collection member. Try using the  **For Each...Next** construct instead of specifying index elements.
    
- You used a shorthand form of subscript that implicitly specified an invalid element. For example, when you use the  **!** operator with a collection, the **!** implicitly specifies a key. For example, _object_**!**_keyname_**.** value is equivalent to _object_**.** item **(**_keyname_**).** value. In this case, an error is generated if _keyname_ represents an invalid key in the collection. To fix the error, use a valid key name or index for the collection.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

