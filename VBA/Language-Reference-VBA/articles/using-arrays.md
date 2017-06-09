---
title: Using Arrays
keywords: vbcn6.chm1076800
f1_keywords:
- vbcn6.chm1076800
ms.prod: office
ms.assetid: 23244a15-4c53-1f39-fdd5-d1fc2b394382
ms.date: 06/08/2017
---


# Using Arrays

You can declare an [array](vbe-glossary.md) to work with a set of values of the same [data type](vbe-glossary.md). An array is a single [variable](vbe-glossary.md) with many compartments to store values, while a typical variable has only one storage compartment in which it can store only one value. Refer to the array as a whole when you want to refer to all the values it holds, or you can refer to its individual elements.

For example, to store daily expenses for each day of the year, you can declare one array variable with 365 elements, rather than declaring 365 variables. Each element in an array contains one value. The following statement declares the array variable with 365 elements. By default, an array is indexed beginning with zero, so the upper bound of the array is 364 rather than 365.



```vb
Dim curExpense(364) As Currency 

```

To set the value of an individual element, you specify the element's index. The following example assigns an initial value of 20 to each element in the array.



```vb
Sub FillArray() 
 Dim curExpense(364) As Currency 
 Dim intI As Integer 
 For intI = 0 to 364 
 curExpense(intI) = 20 
 Next 
End Sub
```


## Changing the Lower Bound

You can use the  **Option Base** statement at the top of a [module](vbe-glossary.md) to change the default index of the first element from 0 to 1. In the following example, the **Option Base** statement changes the index for the first element, and the **Dim** statement declares the array variable with 365 elements.


```vb
Option Base 1 
Dim curExpense(365) As Currency 

```

You can also explicitly set the lower bound of an array by using a  **To** clause, as shown in the following example.




```vb
Dim curExpense(1 To 365) As Currency 
Dim strWeekday(7 To 13) As String 

```


## Storing Variant Values in Arrays

There are two ways to create arrays of  **Variant** values. One way is to declare an array of [Variant data type](vbe-glossary.md), as shown in the following example:


```vb
Dim varData(3) As Variant 
varData(0) = "Claudia Bendel" 
varData(1) = "4242 Maple Blvd" 
varData(2) = 38 
varData(3) = Format("06-09-1952", "General Date") 

```

The other way is to assign the array returned by the  **Array** function to a **Variant** variable, as shown in the following example.




```vb
Dim varData As Variant 
varData = Array("Ron Bendel", "4242 Maple Blvd", 38, _ 
Format("06-09-1952", "General Date")) 

```

You identify the elements in an array of  **Variant** values by index, no matter which technique you use to create the array. For example, the following statement can be added to either of the preceding examples.




```vb
MsgBox "Data for " &; varData(0) &; " has been recorded." 

```


## Using Multidimensional Arrays

In Visual Basic, you can declare arrays with up to 60 dimensions. For example, the following statement declares a 2-dimensional, 5-by-10 array.


```vb
Dim sngMulti(1 To 5, 1 To 10) As Single 

```

If you think of the array as a matrix, the first argument represents the rows and the second argument represents the columns.

Use nested  **For...Next** statements to process multidimensional arrays. The following procedure fills a two-dimensional array with **Single** values.




```vb
Sub FillArrayMulti() 
 Dim intI As Integer, intJ As Integer 
 Dim sngMulti(1 To 5, 1 To 10) As Single 
 
 ' Fill array with values. 
 For intI = 1 To 5 
 For intJ = 1 To 10 
 sngMulti(intI, intJ) = intI * intJ 
 Debug.Print sngMulti(intI, intJ) 
 Next intJ 
 Next intI 
End Sub
```


