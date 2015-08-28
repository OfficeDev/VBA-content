
# Declaring Arrays

 **Last modified:** July 28, 2015

 [Arrays](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) are declared the same way as other [variables](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), using the  **Dim**,  **Static**,  **Private**, or  **Public** statements. The difference between scalar variables (those that aren't arrays) and array variables is that you generally must specify the size of the array. An array whose size is specified is a fixed-size array. An array whose size can be changed while a program is running is a dynamic array.

Whether an array is indexed from 0 or 1 depends on the setting of the  **Option** **Base** statement. If **Option** **Base** **1** is not specified, all array indexes begin at zero.


## Declaring a Fixed Array

In the following line of code, a fixed-size array is declared as an  **Integer** array having 11 rows and 11 columns:


```
Dim MyArray(10, 10) As Integer 

```

The first argument represents the rows; the second argument represents the columns.

As with any other variable declaration, unless you specify a  [data type](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) for the array, the data type of the elements in a declared array is **Variant**. Each numeric  **Variant** element of the array uses 16 bytes. Each string **Variant** element uses 22 bytes. To write code that is as compact as possible, explicitly declare your arrays to be of a data type other than **Variant**. The following lines of code compare the size of several arrays:




```
' Integer array uses 22 bytes (11 elements * 2 bytes). 
ReDim MyIntegerArray(10) As Integer 
 
' Double-precision array uses 88 bytes (11 elements * 8 bytes). 
ReDim MyDoubleArray(10) As Double 
 
' Variant array uses at least 176 bytes (11 elements * 16 bytes). 
ReDim MyVariantArray(10) 
 
' Integer array uses 100 * 100 * 2 bytes (20,000 bytes). 
ReDim MyIntegerArray (99, 99) As Integer 
 
' Double-precision array uses 100 * 100 * 8 bytes (80,000 bytes). 
ReDim MyDoubleArray (99, 99) As Double 
 
' Variant array uses at least 160,000 bytes (100 * 100 * 16 bytes). 
ReDim MyVariantArray(99, 99) 

```

The maximum size of an array varies, based on your operating system and how much memory is available. Using an array that exceeds the amount of RAM available on your system is slower because the data must be read from and written to disk.


## Declaring a Dynamic Array

By declaring a dynamic array, you can size the array while the code is running. Use a  **Static**,  **Dim**,  **Private**, or  **Public** statement to declare an array, leaving the parentheses empty, as shown in the following example.


```
Dim sngArray() As Single 

```


 **Note**  You can use the  **ReDim** statement to declare an array implicitly within a procedure. Be careful not to misspell the name of the array when you use the **ReDim** statement. Even if the **Option Explicit** statement is included in the module, a second array will be created.

In a procedure within the array's  [scope](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), use the  **ReDim** statement to change the number of dimensions, to define the number of elements, and to define the upper and lower bounds for each dimension. You can use the **ReDim** statement to change the dynamic array as often as necessary. However, each time you do this, the existing values in the array are lost. Use **ReDim Preserve** to expand an array while preserving existing values in the array. For example, the following statement enlarges the array by 10 elements without losing the current values of the original elements.




```
ReDim Preserve varArray(UBound(varArray) + 10) 

```


 **Note**  When you use the  **Preserve** [keyword](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) with a dynamic array, you can change only the upper bound of the last dimension, but you can't change the number of dimensions.

