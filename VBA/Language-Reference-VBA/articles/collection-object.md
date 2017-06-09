---
title: Collection Object
keywords: vblr6.chm1013894
f1_keywords:
- vblr6.chm1013894
ms.prod: office
api_name:
- Office.Collection
ms.assetid: 1bc5c060-34c7-84e7-c99c-f20266a2d071
ms.date: 06/08/2017
---


# Collection Object



A  **Collection** object is an ordered set of items that can be referred to as a unit.
 **Remarks**
The  **Collection** object provides a convenient way to refer to a related group of items as a single object. The items, or[members](vbe-glossary.md), in a collection need only be related by the fact that they exist in the [collection](vbe-glossary.md). Members of a collection don't have to share the same [data type](vbe-glossary.md).
A collection can be created the same way other objects are created. For example:



```vb
Dim X As New Collection

```

Once a collection is created, members can be added using the  **Add** method and removed using the **Remove** method. Specific members can be returned from the collection using the **Item** method, while the entire collection can be iterated using the **For** **Each...Next** statement.

## Example

This example creates a  **Collection** object ( `MyClasses`), and then creates a dialog box in which users can add objects to the collection. To see how this works, choose the  **Class Module** command from the **Insert** menu and declare a public variable called `InstanceName` at module level of Class1 (type **Public** `InstanceName)` to hold the names of each instance. Leave the default name as Class1. Copy and paste the following code into the General section of another module, and then start it with the statement ), `ClassNamer`and then creates a dialog box in which users can add objects to the collection. To see how this works, choose the  **Class Module** command from the **Insert** menu and declare a public variable called `InstanceName` at module level of Class1 (type **Public** `InstanceName)` to hold the names of each instance. Leave the default name as Class1. Copy and paste the following code into the General section of another module, and then start it with the statement `ClassNamer` in another procedure. (This example only works with host applications that support classes.)


```vb
Sub ClassNamer()
    Dim MyClasses As New Collection    ' Create a Collection object.
    Dim Num    ' Counter for individualizing keys.
    Dim Msg As String    ' Variable to hold prompt string.
    Dim TheName, MyObject, NameList    ' Variants to hold information.
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
            MyClasses.Add item := Inst, key := CStr(Num)
        End If
        ' Clear the current reference in preparation for next one.
        Set Inst = Nothing
    Loop Until TheName = ""
    For Each MyObject In MyClasses    ' Create list of names.
        NameList = NameList &; MyObject.InstanceName &; Chr(13)
    Next MyObject
    ' Display the list of names in a message box.
    MsgBox NameList, , "Instance Names In MyClasses Collection"

    For Num = 1 To MyClasses.Count    ' Remove name from the collection.
        MyClasses.Remove 1    ' Since collections are reindexed
                ' automatically, remove the first
    Next        ' member on each iteration.
End Sub
```


