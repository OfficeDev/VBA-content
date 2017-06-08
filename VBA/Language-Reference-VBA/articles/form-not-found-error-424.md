---
title: Form not found (Error 424)
keywords: vblr6.chm1117811
f1_keywords:
- vblr6.chm1117811
ms.prod: office
ms.assetid: e2f313ac-40ea-911e-b1cb-c4ccd8745b2e
ms.date: 06/08/2017
---


# Form not found (Error 424)

The form was not found. This error has the following cause and solution:

You tried to add a form to the  **UserForms** collection using the **Add** method, but there is no form[class](vbe-glossary.md) of that name. For example, `UserForms.Add "MyForm"`, where , where  `MyForm` doesn't exist.

 **Note**  Make sure that the class name is available to your [project](vbe-glossary.md).


