---
title: TextRange2.Words Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Words
ms.assetid: bab78b31-ebd6-649e-0b05-5b21552f8f22
ms.date: 06/08/2017
---


# TextRange2.Words Property (Office)

Gets a  **TextRange2** object that represents the specified subset of text words. Read-only.


## Syntax

 _expression_. **Words**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first word in the returned range.|
| _Length_|Optional|**Long**|The number of words to be returned.|

### Return Value

TextRange2


## Remarks

If both  _Start_ and _Length_ are omitted, the returned range starts with the first word and ends with the last paragraph in the specified range.

If  _Start_ is specified but _Length_ is omitted, the returned range contains one word.

If  _Length_ is specified but _Start_ is omitted, the returned range starts with the first word in the specified range.

If  _Start_ is greater than the number of words in the specified text, the returned range starts with the last word in the specified range.

If  _Length_ is greater than the number of words from the specified starting word to the end of the text, the returned range contains all those words.


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

