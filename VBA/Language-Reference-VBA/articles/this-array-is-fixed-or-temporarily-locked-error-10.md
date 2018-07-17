---
title: This array is fixed or temporarily locked (Error 10)
keywords: vblr6.chm1019365
f1_keywords:
- vblr6.chm1019365
ms.prod: office
ms.assetid: 075c8897-c6e6-839f-a372-1e2249fc99e8
ms.date: 06/08/2017
---


# This array is fixed or temporarily locked (Error 10)

Not all [arrays](vbe-glossary.md) can be redimensioned. Even arrays specifically declared to be dynamic and arrays within **Variant**[variables](vbe-glossary.md) are sometimes locked temporarily. This error has the following causes and solutions:



- You tried to use  **ReDim** to change the number of elements of a fixed-size array . For example, in the following code, the fixed array `FixedArr` is received by `SomeArr` in the `NextOne` procedure, and then an attempt is made to resize `SomeArr`:
    
```vb
Sub FirstOne 
  Dim FixedArr(25) As Integer    ' Create a fixed-size array and 
  NextOne FixedArr()    ' pass it to another procedure. 
End Sub 
 
Sub NextOne(SomeArr() As Integer) 
  ReDim SomeArr(35)        ' Error 10 occurs here. 
  '. . . 
End Sub 
```


     Make the original array dynamic rather than fixed by declaring it with **ReDim** (if the array is declared within a procedure), or by declaring it without specifying the number of elements (if the array is declared at[module level](vbe-glossary.md)).
    
- You tried to redimension a module-level dynamic array, in which one element has been passed as an [argument](vbe-glossary.md) to a procedure. For example, in the following code, `ModArray` is a dynamic, module-level array whose forty-fifth element is being passed[by reference](vbe-glossary.md) to the `Test` procedure:
    
```vb
Dim ModArray () As Integer    ' Create a module-level dynamic array. 
'. . . 
 
Sub AliasError() 
  ReDim ModArray (1 To 73) As Integer 
Test ModArray(45)    ' Pass an element of the module-level  
' array to the Test procedure. 
End Sub 
 
Sub Test(SomeInt As Integer) 
  ReDim ModArray (1 To 40) As Integer  ' Error occurs here. 
End Sub 
```


    There is no need to pass an element of the module-level array in this case, since it's visible within all procedures in the module. However, if an element is passed, the array is locked to prevent a deallocation of memory for the reference [parameter](vbe-glossary.md) within the procedure, causing unpredictable behavior when the procedure returns.
    
- You attempted to assign a value to a  **Variant** variable containing an array, but the **Variant** is currently locked. For example, if your code uses a **For Each...Next** loop to iterate over a variant containing an array, the array is locked on entry into the loop, and then released at the termination of the loop:
    
```vb
SomeArray = Array(9,8,7,6,5,4,3,2,1) 
 
For Each X In SomeArray 
  SomeArray = 301    ' Causes error since array is locked. 
Next X 
```


     Use a **For...Next** rather than a **For Each...Next** loop to iterate. When an array is the object of a **For Each...Next** loop, you can read the array, but not write to it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

