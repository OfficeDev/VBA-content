---
title: Application.References Property (Access)
keywords: vbaac10.chm12564
f1_keywords:
- vbaac10.chm12564
ms.prod: access
api_name:
- Access.Application.References
ms.assetid: da78f26f-1127-796d-bba1-f1c0d98a582e
ms.date: 06/08/2017
---


# Application.References Property (Access)

You can use the  **References** property to access the **[References](references-object-access.md)** collection and its related properties, methods, and events. Read-only **References** collection.


## Syntax

 _expression_. **References**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **References** collection corresponds to the list of references in the **References** dialog box, available by clicking **References** on the **Tools** menu. Each **Reference** object represents one selected reference in the list. References that appear in the **References** dialog box but haven't been selected aren't in the **References** collection.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Display List of References - Both Good and Broken](http://www.utteraccess.com/forum/Display-List-References-t126939.html)
    

## Example

The following example displays a message indicating the number of boxes checked in the  **References** dialog box.


```vb
MsgBox "There are " &; Application.References.Count &; " references."
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

