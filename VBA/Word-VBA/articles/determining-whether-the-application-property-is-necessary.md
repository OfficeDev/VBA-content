---
title: Determining Whether the Application Property Is Necessary
keywords: vbawd10.chm5210509
f1_keywords:
- vbawd10.chm5210509
ms.prod: word
ms.assetid: a8666cf9-f42c-8dc6-ac40-df487b4bfeeb
ms.date: 06/08/2017
---


# Determining Whether the Application Property Is Necessary

Many of the properties and methods of the  **[Application](application-object-word.md)** object can be used without the **Application** object qualifier. For example the **[ActiveDocument](application-activedocument-property-word.md)** property can be used without the **Application** object qualifier. Instead of writing `Application.ActiveDocument.PrintOut`, you can write  `ActiveDocument.PrintOut`.

Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the Object Browser, click **&lt;globals&gt;** at the top of the list in the **Classes** box.

