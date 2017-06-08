---
title: Assigning Ranges
keywords: vbawd10.chm5209930
f1_keywords:
- vbawd10.chm5209930
ms.prod: word
ms.assetid: abcdd18e-8d0e-13cd-0ec2-721dde99f9d6
ms.date: 06/08/2017
---


# Assigning Ranges

There are several ways to assign an existing  **[Range](range-object-word.md)** object to a variable. This topic explains the results of two different techniques. In the following examples, the `Range1` and `Range2` variables refer to **Range** objects. For example, the following instructions assign the first and second words in the active document to the `Range1` and `Range2` variables.


```vb
Set Range1 = ActiveDocument.Words(1) 
Set Range2 = ActiveDocument.Words(2)
```


## Setting a Range object variable equal to another Range object variable

The following instruction assigns a range variable named  `Range2` to represent the same location as `Range1`.


```vb
Set Range2 = Range1
```

You now have two variables that represent the same range. When you manipulate the start or endpoint or the text of  `Range2`, it affects  `Range1` and vice versa.

Note that the following instruction is the same as  `Range2.Text = Range1.Text`. This instruction assigns the default property of  `Range1`, which is the  **[Text](range-text-property-word.md)** property, to the default property of `Range2`. It doesn't change what the objects actually refer to.




```
Range2 = Range1
```

The ranges ( `Range2` and and `Range1`) have the same contents, but they may point to different locations in the document or even different documents.


## Using the Duplicate property

The following instruction creates a new duplicated  **Range** object, `Range2`, which has the same start and endpoints and text as  `Range1`.


```vb
Set Range2 = Range1.Duplicate
```

If you change the start or endpoint of  `Range1`, it doesn't affect  `Range2,` and vice versa. Because these two ranges point to the same location in the document, changing the text in one range affects the text in the other range.


