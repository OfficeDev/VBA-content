---
title: Add Method (Visual Basic for Applications)
keywords: vblr6.chm1014017
f1_keywords:
- vblr6.chm1014017
ms.prod: office
ms.assetid: c9e9eb2e-42b1-9821-67ab-2f68fb87d1d0
ms.date: 06/08/2017
---


# Add Method (Visual Basic for Applications)



Adds a [member](vbe-glossary.md) to a **Collection** object.
 **Syntax**
 _object_**.Add  _item_,** **_key_,** **_before_,** **_after_**
The  **Add** method syntax has the following object qualifier and[named arguments](vbe-glossary.md):


| <strong>Part</strong>            | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
|:---------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>                  | Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| <strong><em>item</em></strong>   | Required. An [expression](vbe-glossary.md) of any type that specifies the member to add to the[collection](vbe-glossary.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| <strong><em>key</em></strong>    | Optional. A unique [string expression](vbe-glossary.md) that specifies a key string that can be used, instead of a positional index, to access a member of the collection.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| <strong><em>before</em></strong> | Optional. An expression that specifies a relative position in the collection. The member to be added is placed in the collection before the member identified by the  <strong><em>before</em></strong>[argument](vbe-glossary.md). If a [numeric expression](vbe-glossary.md),  <strong><em>before</em></strong> must be a number from 1 to the value of the collection's <strong>Count</strong> property. If a string expression, <strong><em>before</em></strong> must correspond to the <strong><em>key</em></strong> specified when the member being referred to was added to the collection. You can specify a <strong><em>before</em></strong> position or an <strong><em>after</em></strong> position, but not both. |
| <strong><em>after</em></strong>  | Optional. An expression that specifies a relative position in the collection. The member to be added is placed in the collection after the member identified by the  <strong><em>after</em></strong> argument. If numeric, <strong><em>after</em></strong> must be a number from 1 to the value of the collection's <strong>Count</strong> property. If a string, <strong><em>after</em></strong> must correspond to the <strong><em>key</em></strong> specified when the member referred to was added to the collection. You can specify a <strong><em>before</em></strong> position or an <strong><em>after</em></strong> position, but not both.                                                                         |

 **Remarks**
Whether the  **_before_** or **_after_** argument is a string expression or numeric expression, it must refer to an existing member of the collection, or an error occurs.
An error also occurs if a specified  **_key_** duplicates the **_key_** for an existing member of the collection.

## Example

This example uses the  **Add** method to add `Inst` objects (instances of a class called `Class1` containing a **Public** variable `InstanceName`) to a collection called  `MyClasses`. To see how this works, insert a class module and declare a public variable called  `InstanceName` at module level of `Class1` (type `Public InstanceName`) to hold the names of each instance. Leave the default name as  `Class1`. Copy and paste the following code into the  `Form_Load` event procedure of a form module.


```vb
Dim MyClasses As New Collection    ' Create a Collection object.
Dim Num As Integer    ' Counter for individualizing keys.
Dim Msg
Dim TheName    ' Holder for names user enters.
Do
    Dim Inst As New Class1    ' Create a new instance of Class1.
    Num = Num + 1    ' Increment Num, then get a name.
    Msg = "Please enter a name for this object." &; Chr(13) _
     &; "Press Cancel to see names in collection."
    TheName = InputBox(Msg, "Name the Collection Items")
    Inst.InstanceName = TheName    ' Put name in object instance.
    ' If user entered name, add it to the collection.
    If Inst.InstanceName <> "" Then
        ' Add the named object to the collection.
        MyClasses. Add item := Inst, key := CStr(Num)
    End If
    ' Clear the current reference in preparation for next one.
    Set Inst = Nothing
Loop Until TheName = ""
For Each x In MyClasses
    MsgBox x.instancename, , "Instance Name"
Next
```


