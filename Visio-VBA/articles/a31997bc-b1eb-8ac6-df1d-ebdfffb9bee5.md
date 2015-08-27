
# ValidationRuleSets.Item Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Returns the  ** [ValidationRuleSet](cd2fc58a-5d7c-cf31-7aab-41bdeee9f105.md)** object that has the specified universal name or index position. Read-only.


## Syntax

 _expression_. **Item**( **_NameUOrIndex_**)

 _expression_A variable that represents a  ** [ValidationRuleSets](f08d7f04-13ec-8175-2aa6-94b0b67ee76b.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|NameUOrIndex|Required| **Variant**|The universal name of the object, or the index number of the object in its collection.|

### Return Value

 **ValidationRuleSet**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections.

