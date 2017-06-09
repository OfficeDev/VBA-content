---
title: Characters.TabPropsRow Property (Visio)
keywords: vis_sdr.chm10214480
f1_keywords:
- vis_sdr.chm10214480
ms.prod: visio
api_name:
- Visio.Characters.TabPropsRow
ms.assetid: 83002645-df6c-5565-b62a-983960a8a8a3
ms.date: 06/08/2017
---


# Characters.TabPropsRow Property (Visio)

Returns the index of the row in the Tabs section of the ShapeSheet that contains tab formatting information for a  **Characters** object. Read-only.


## Syntax

 _expression_ . **TabPropsRow**( **_BiasLorR_** )

 _expression_ A variable that represents a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BiasLorR_|Required| **Integer**|The direction of the search.|

### Return Value

Integer


## Remarks

You can retrieve rows that represent runs of tab formatting by specifying a row index as an argument to the  **CellsSRC** property of a shape. You can also view or change tab formats on the **Tabs** tab of the **Text** dialog box (on the **Home** tab, click the **Font** arrow, and then click **Tabs**).

If the tab format for the  **Characters** object is represented by more than one tab properties row, the **TabPropsRow** property returns -1. Under these circumstances, Microsoft Visio ignores the value of the _BiasLorR_ argument. (Two paragraphs that have the same tab formatting share the same row in the ShapeSheet. Visio creates a new ShapeSheet row only when tab formatting changes, for example from left-aligned to right-aligned.)

If the  **Characters** object spans several characters within the same tab properties row, **ParaPropsRow** returns the index of that row. In this case as well, Visio ignores the _BiasLorR_ argument.

If the  **Characters** object represents an insertion point rather than a sequence of characters (that is, if its **Begin** and **End** properties return the same value), use the _BiasLOrR_ argument to determine which row index to return.



|**Constant **|**Value **|
|:-----|:-----|
| **visBiasLetVisioChoose**|0 |
| **visBiasLeft**|1 |
| **visBiasRight**|2 |
Specify  **visBiasLeft** for the row that covers tab formatting for the character to the left of the insertion point. Use **visBiasRight** for the row that covers tab formatting for the character to the right of the insertion point.

If you specify  **visBiasLetVisioChoose** , Visio uses the same logic it would apply to new text typed in the user interface starting at the insertion point. Usually, that means that Visio will apply the tab formatting of the character to the left of the insertion point to the new text, so **TabPropsRow** will return the same value it would if passed **visBiasLeft** . (For an explanation of the meaning of "left" in this context, see the following note.) However, if the insertion point is at the beginning of a new paragraph, **TabPropsRow** returns the value it would return if passed **visBiasRight** .




 **Note**  In the context of a  **Characters** object, "left" means logically prior. In other words, one character is to the "left" of another if it would have been typed first in the course of normal writing. It is necessary to make this distinction because in some languages, characters are normally written from right to left, and not from left to right.


