---
title: Creating Object Variables
keywords: vbcn6.chm1011337
f1_keywords:
- vbcn6.chm1011337
ms.prod: office
ms.assetid: 6cff962e-4a3e-dfc3-8491-d31a308b1c55
ms.date: 06/08/2017
---


# Creating Object Variables

You can treat an [object variable](vbe-glossary.md) exactly the same as the [object](vbe-glossary.md) to which it refers. You can set or return the [properties](vbe-glossary.md) of the object or use any of its [methods](vbe-glossary.md).

 **To create an object variable:**




1. Declare the object variable.
    
2. Assign the object variable to an object.
    


## Declaring an Object Variable

Use the  **Dim** statement or one of the other declaration statements ( **Public**, **Private**, or **Static** ) to declare an object variable. A [variable](vbe-glossary.md) that refers to an object must be a **Variant**, an **Object**, or a specific type of object. For example, the following declarations are valid:


```vb
' Declare MyObject as Variant data type. 
Dim MyObject 
' Declare MyObject as Object data type. 
Dim MyObject As Object 
' Declare MyObject as Font type. 
Dim MyObject As Font 

```


 **Note**  If you use an object variable without declaring it first, the [data type](vbe-glossary.md) of the object variable is **Variant** by default.

You can declare an object variable with the  **Object** data type when the specific [object type](vbe-glossary.md) is not known until the procedure runs. Use the **Object** data type to create a generic reference to any object.

If you know the specific object type, you should declare the object variable as that object type. For example, if the application contains a Sample object type, you can declare an object variable for that object using either of these statements:




```vb
Dim MyObject As Object ' Declared as generic object. 
Dim MyObject As Sample ' Declared only as Sample object. 

```

Declaring specific object types provides automatic type checking, faster code, and improved readability.


## Assigning an Object Variable to an Object

Use the  **Set** statement to assign an object to an object variable. You can assign an [object expression](vbe-glossary.md) or **Nothing**. For example, the following object variable assignments are valid:


```vb
Set MyObject = YourObject ' Assign object reference. 
Set MyObject = Nothing ' Discontinue association. 

```

You can combine declaring an object variable with assigning an object to it by using the  **New** [keyword](vbe-glossary.md) with the **Set** statement. For example:




```vb
Set MyObject = New Object ' Create and Assign 

```

Setting an object variable equal to  **Nothing** discontinues the association of the object variable with any specific object. This prevents you from accidentally changing the object by changing the variable. An object variable is always set to **Nothing** after closing the associated object so you can test whether or not the object variable points to a valid object. For example:




```
If Not MyObject Is Nothing Then 
 ' Variable refers to valid object. 
 . . . 
End If 

```

Of course, this test can never determine with absolute certainty whether or not a user has closed the application containing the object to which the object variable refers.


## Referring to the Current Instance of an Object

Use the  **Me** keyword to refer to the current instance of the object where the code is running. All procedures associated with the current object have access to the object referred to as **Me**. Using **Me** is particularly useful for passing information about the current instance of an object to a procedure in another module. For example, suppose you have the following procedure in a module:


```vb
Sub ChangeObjectColor(MyObjectName As Object) 
 MyObjectName.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256) 
End Sub
```

You can call the procedure and pass the current instance of the object as an argument using the following statement:




```
ChangeObjectColor Me 

```


