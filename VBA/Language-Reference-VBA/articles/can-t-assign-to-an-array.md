---
title: Can't assign to an array
keywords: vblr6.chm1011081
f1_keywords:
- vblr6.chm1011081
ms.prod: office
ms.assetid: cc606f0f-7e50-c144-8003-90c7f976723d
ms.date: 06/08/2017
---


# Can't assign to an array

Each element of an [array](vbe-glossary.md) must have its value assigned individually. This error has the following causes and solutions:


- You inadvertently tried to assign a single value to an array [variable](vbe-glossary.md) without specifying the element to which the value should be assigned.
    
    
    
    To assign a single value to an array element, you must specify the element in a subscript. For example, if  `MyArray` is an integer array, the[expression ](vbe-glossary.md) ` MyArray = 5` is invalid, but the following expression is valid: `MyArray(UBound(MyArray)) = 5`
    
- You tried to assign a whole array to another array. 
    
    
    
    For example, if  `Arr1` is an array and `Arr2` is another array, the following two assignments are both invalid:
    


  ```
  Arr1 = Arr2    ' Invalid assignment. 
Arr1() = Arr2()    ' Invalid assignment. 

  ```


    To assign one array to another, make sure the array on the left-hand side of the assignment is resizable and the types of the array match.
    
     **Note**  You can place a whole array in a  **Variant**, resulting in a single variant variable containing the whole array:



```vb
Dim MyArr As Variant 
MyVar = Arr2() 
  ```


    You then reference the elements of the array in the variant with the same subscript notation as for a normal array, for example:
    


  ```
  MyVar(3) = MyVar(1) + MyVar(5) 
  ```




For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

