---
title: Invalid ReDim
keywords: vblr6.chm1011326
f1_keywords:
- vblr6.chm1011326
ms.prod: office
ms.assetid: 32682d86-d3c1-bf15-677c-9b0efd5e9e75
ms.date: 06/08/2017
---


# Invalid ReDim

Not every [array](vbe-glossary.md) can be redimensioned. This error has the following causes and solutions:



- A [variable](vbe-glossary.md) was implicitly declared a **Variant**, and you attempted to use **ReDim** to change it to an array.
    
    A  **Variant** can contain an array, but if it isn't explicitly declared, you can't use **ReDim** to make it into an array. Declare the **Variant** before using **ReDim** to specify the number of elements it can contain. For example, in the following code, `ReDim AVar(10)` causes an invalid **ReDim** error, but `ReDim BVar(10)` does not:
    


```vb
AVar = 1    ' Implicit declaration of AVar. 
ReDim AVar(10)    ' Causes invalid ReDim error. 
'. 
'. 
'. 
Dim BVar    ' Explicit declaration of BVar. 
ReDim BVar(10)    ' No error. 
```


    
    
- You tried to use  **ReDim** to change more than one dimension of an array contained within a **Variant**. You can only use **ReDim** to change the size of the last dimension of an array in a **Variant**. To create an array with multiple dimensions that can be redimensioned, the array can't be contained within a **Variant**, and you have to declare it the normal way.
    
- You can use  **ReDim** only to change the number of elements in a normal array, not the type of those elements. If you want an array in which you can change the types of the elements, use an array contained within a **Variant**. If you declare the array first, changing the types and the number of its elements can be accomplished as follows:
    
```vb
Dim MyVar As Variant    ' Declare the variable. 
ReDim MyVar(10) As String    ' ReDim it as array of String subtypes. 
ReDim MyVar(20) As Integer    ' ReDim it as array of Integer subtypes. 
ReDim MyVar(5) As Variant    ' ReDim it as array of Variant subtypes. 

  ```


    
    
- You attempted to use  **ReDim** with an array that is a member of an Automation object.
    
    Remove the  **ReDim**.
    
     **Note**  If you don't specify a type for a variable, the variable receives the default type,  **Variant**. This isn't always obvious. For example, the following code declares two variables, the first, `MyVar`, is a  **Variant**; the second, `AnotherVar`, is an  **Integer**.




```vb
Dim MyVar, AnotherVar As Integer 

```

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

