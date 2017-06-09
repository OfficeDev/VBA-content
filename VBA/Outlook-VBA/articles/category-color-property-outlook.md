---
title: Category.Color Property (Outlook)
keywords: vbaol11.chm2427
f1_keywords:
- vbaol11.chm2427
ms.prod: outlook
api_name:
- Outlook.Category.Color
ms.assetid: 42814031-97ee-bb71-7c24-4ddd367d793c
ms.date: 06/08/2017
---


# Category.Color Property (Outlook)

Returns or sets an  **[OlCategoryColor](olcategorycolor-enumeration-outlook.md)** constant that indicates the color used by the **[Category](category-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **Color**

 _expression_ A variable that represents a **Category** object.


## Remarks

You can share the same color for multiple categories, by specifying the same constant that represents the category color in the  **OlCategoryColor** enumeration for those **Category** objects.


## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing color assignments for each  **Category** object contained in the **[Categories](namespace-categories-property-outlook.md)** collection associated with the default **[NameSpace](namespace-object-outlook.md)** object.


```vb
Private Sub ListCategoryColors() 
 Dim objNameSpace As NameSpace 
 Dim objCategory As Category 
 Dim strOutput As String 
 
 ' Obtain a NameSpace object reference. 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 ' Check if the Categories collection for the Namespace 
 ' contains one or more Category objects. 
 If objNameSpace.Categories.Count > 0 Then 
 
 ' Enumerate the Categories collection, checking 
 ' the value of the Color property for 
 ' each Category object. 
 For Each objCategory In objNameSpace.Categories 
 
 ' Add the name of the Category object to 
 ' the output string. 
 strOutput = strOutput &; objCategory.Name 
 
 ' Add information about the assigned color 
 ' to the output string. 
 Select Case objCategory.Color 
 Case OlCategoryColor.olCategoryColorNone 
 strOutput = strOutput &; ": No color" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorBlack 
 strOutput = strOutput &; ": Black " &; vbCrLf 
 Case OlCategoryColor.olCategoryColorBlue 
 strOutput = strOutput &; ": Blue" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorGray 
 strOutput = strOutput &; ": Gray" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorGreen 
 strOutput = strOutput &; ": Green" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightBlue 
 strOutput = strOutput &; ": Light blue" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightGray 
 strOutput = strOutput &; ": Light gray" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightGreen 
 strOutput = strOutput &; ": Light green" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightMaroon 
 strOutput = strOutput &; ": Light maroon" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightOlive 
 strOutput = strOutput &; ": Light olive" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightOrange 
 strOutput = strOutput &; ": Light orange" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightPeach 
 strOutput = strOutput &; ": Light peach" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightPurple 
 strOutput = strOutput &; ": Light purple" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightRed 
 strOutput = strOutput &; ": Light red" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightSteel 
 strOutput = strOutput &; ": Light steel" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightTeal 
 strOutput = strOutput &; ": Light teal" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorLightYellow 
 strOutput = strOutput &; ": Light yellow" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorMaroon 
 strOutput = strOutput &; ": Maroon" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorOlive 
 strOutput = strOutput &; ": Olive" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorOrange 
 strOutput = strOutput &; ": Orange" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorPeach 
 strOutput = strOutput &; ": Peach" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorPurple 
 strOutput = strOutput &; ": Purple" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorRed 
 strOutput = strOutput &; ": Red" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorSteel 
 strOutput = strOutput &; ": Steel" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorTeal 
 strOutput = strOutput &; ": Teal" &; vbCrLf 
 Case OlCategoryColor.olCategoryColorYellow 
 strOutput = strOutput &; ": Yellow" &; vbCrLf 
 Case Else 
 strOutput = strOutput &; ": Unknown" &; vbCrLf 
 End Select 
 Next 
 End If 
 
 ' Display the output string. 
 MsgBox strOutput 
 
 ' Clean up. 
 Set objCategory = Nothing 
 Set objNameSpace = Nothing 
 
End Sub
```


## See also


#### Concepts


[Category Object](category-object-outlook.md)

