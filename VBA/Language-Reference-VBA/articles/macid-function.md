---
title: MacID Function
keywords: vblr6.chm1009753
f1_keywords:
- vblr6.chm1009753
ms.prod: office
ms.assetid: 4df07ec9-c165-ab5a-2864-ef1d9168d7e5
ms.date: 06/08/2017
---


# MacID Function



Used on the Macintosh to convert a 4-character [constant](vbe-glossary.md) to a value that may be used by <strong>Dir</strong>, <strong>Kill</strong>, <strong>Shell</strong>, and <strong>AppActivate</strong>.
 
<strong>Syntax</strong>
 
<strong>MacID(</strong> constant <strong>)</strong>
The required  
<em>constant</em> argument consists of 4 characters used to specify a resource type, file type, application signature, or Apple Event, for example, TEXT, OBIN, "XLS5" for Excel files ("XLS8" for Excel 97), Microsoft Word uses "W6BN" ("W8BN" for Word 97), and so on.
 
<strong>Remarks</strong>
 
<strong>MacID</strong> is used with <strong>Dir</strong> and <strong>Kill</strong> to specify a Macintosh file type. Since the Macintosh does not support <strong>\</strong>* and <strong>?</strong> as wildcards, you can use a four-character constant instead to identify groups of files. For example, the following statement returns TEXT type files from the current folder:



```
Dir("SomePath", MacID("TEXT"))
```

 **MacID** is used with **Shell** and **AppActivate** to specify an application using the application's unique signature.

