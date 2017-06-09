---
title: HangulHanjaConversionDictionaries Object (Word)
ms.prod: word
ms.assetid: b6ed1c54-428b-c160-a2bd-642978660f44
ms.date: 06/08/2017
---


# HangulHanjaConversionDictionaries Object (Word)

A collection of  **Dictionary** objects that includes the active custom Hangul-Hanja conversion dictionaries.


## Remarks

Use the  **HangulHanjaDictionaries** property to return the collection of currently active custom conversion dictionaries. The following example displays the names of all the active custom conversion dictionaries.


```vb
For Each d In HangulHanjaDictionaries 
 Msgbox d.Name 
Next d
```

Use the  **Add** method to add a new custom conversion dictionary to the collection of active custom conversion dictionaries. If there isn't a file with the name specified by **FileName** , Microsoft Word creates it. The following example adds "Hanja1.hhd" to the collection of custom conversion dictionaries.




```
CustomDictionaries.Add FileName:="Hanja1.hhd"
```

Use the  **ClearAll** method to unload all custom conversion dictionaries. Note, however, that this method doesn't delete the dictionary files. After you use this method, the number of custom conversion dictionaries in the collection is 0 (zero). The following example clears the custom conversion dictionaries and creates a new custom conversion dictionary file. The new dictionary is set as the active custom dictionary to which Word will automatically add any new words it encounters.




```vb
With HangulHanjaDictionaries 
 .ClearAll 
 .Add FileName:= "Hanja1.hhd" 
 .ActiveCustomDictionary = HangulHanjaDictionaries(1) 
End With
```

You set the custom dictionary to which new words are added by using the  **ActiveCustomDictionary** property. If you try to set this property to a dictionary that isn't a custom conversion dictionary, an error occurs.

The  **Maximum** property returns the maximum number of simultaneous custom conversion dictionaries that the application can support. For Word, this maximum is 10.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


