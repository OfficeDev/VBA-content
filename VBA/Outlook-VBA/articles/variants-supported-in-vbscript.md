---
title: Variants Supported in VBScript
keywords: olfm10.chm3077516
f1_keywords:
- olfm10.chm3077516
ms.prod: outlook
ms.assetid: 85e65768-5648-592b-08eb-9b646ec1b8a3
ms.date: 06/08/2017
---


# Variants Supported in VBScript

Microsoft VBScript in Outlook uses only the  **Variant** data type.

A  **Variant** is a special kind of data type that can contain different kinds of information, depending on how it's used. Because **Variant** is the only data type in VBScript, it's also the data type returned by all functions in VBScript.

At its simplest, a  **Variant** can contain either numeric or string information. A **Variant** behaves as a number when you use it in a numeric context and as a string when you use it in a string context. That is, if you're working with data that looks like numbers, VBScript assumes that it is numbers and does the thing that is most appropriate for numbers. Similarly, if you're working with data that can only be string data, VBScript treats it as string data. Of course, you can always make numbers behave as strings by enclosing them in quotation marks (" ").


