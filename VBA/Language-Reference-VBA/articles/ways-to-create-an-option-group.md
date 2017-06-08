---
title: Ways to create an option group
keywords: fm20.chm5225256
f1_keywords:
- fm20.chm5225256
ms.prod: office
ms.assetid: 03e01236-e877-11a1-5de7-52d6307185e7
ms.date: 06/08/2017
---


# Ways to create an option group

By default, all  **OptionButton** controls on a[container](vbe-glossary.md) (such as a form, a **MultiPage**, or a **Frame** ) are part of a single option group. This means that selecting one of the buttons automatically sets all other option buttons on the form to **False**.

If you want more than one option group on the form, there are two ways to create additional groups:




- Use the  **GroupName** property to identify related buttons.
    
- Put related buttons in a  **Frame** on the form.
    

The first method is recommended over the second because it reduces the number of controls required in the application. This reduces the disk space required for your application and can improve the performance of your application as well.

 **Note**  A  **TabStrip** is not a container. Option buttons in the **TabStrip** are included in the form's option group. You can use **GroupName** to create an option group from buttons in a **TabStrip**.


