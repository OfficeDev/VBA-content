---
title: Me <keyword>
keywords: vblr6.chm1008868
f1_keywords:
- vblr6.chm1008868
ms.prod: office
ms.assetid: 6d062019-bb49-7acb-5f03-7bb5a2a09681
ms.date: 06/08/2017
---


# Me <keyword>

The  **Me**[keyword](vbe-glossary.md) behaves like an implicitly declared[variable](vbe-glossary.md). It is automatically available to every [procedure](vbe-glossary.md) in a[class module](vbe-glossary.md). When a [class](vbe-glossary.md) can have more than one instance, **Me** provides a way to refer to the specific instance of the class where the code is executing. Using **Me** is particularly useful for passing information about the currently executing instance of a class to a procedure in another[module](vbe-glossary.md). For example, suppose you have the following procedure in a module:


```vb
Sub ChangeFormColor(FormName As Form) 
 FormName.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256) 
End Sub
```


You can call this procedure and pass the current instance of the Form class as an [argument](vbe-glossary.md) using the following[statement](vbe-glossary.md):




```
ChangeFormColor Me 

```


