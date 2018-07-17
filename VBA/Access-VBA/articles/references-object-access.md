---
title: References Object (Access)
keywords: vbaac10.chm12648
f1_keywords:
- vbaac10.chm12648
ms.prod: access
api_name:
- Access.References
ms.assetid: ac020382-4ece-f138-d1b9-d05b0fe0f523
ms.date: 06/08/2017
---


# References Object (Access)

The  **References** collection contains **Reference** objects representing each reference that's currently set.


## Remarks

The  **Reference** objects in the **References** collection correspond to the list of references in the **References** dialog box, available by clicking **References** on the **Tools** menu. Each **Reference** object represents one selected reference in the list. References that appear in the **References** dialog box but haven't been selected aren't in the **References** collection.

You can enumerate through the  **References** collection by using the **For Each...Next** statement.

The  **References** collection belongs to the Microsoft Access **Application** object.

Individual  **Reference** objects in the **References** collection are indexed beginning with 1.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Display List of References - Both Good and Broken](http://www.utteraccess.com/forum/Display-List-References-t126939.html)
    

## Events



|**Name**|
|:-----|
|[ItemAdded](references-itemadded-event-access.md)|
|[ItemRemoved](references-itemremoved-event-access.md)|

## Methods



|**Name**|
|:-----|
|[AddFromFile](references-addfromfile-method-access.md)|
|[AddFromGuid](references-addfromguid-method-access.md)|
|[Item](references-item-method-access.md)|
|[Remove](references-remove-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Count](references-count-property-access.md)|
|[Parent](references-parent-property-access.md)|

## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
