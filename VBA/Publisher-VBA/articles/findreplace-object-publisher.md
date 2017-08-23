---
title: "Объект FindReplace (издатель)"
keywords: vbapb10.chm8388607
f1_keywords: vbapb10.chm8388607
ms.prod: publisher
api_name: Publisher.FindReplace
ms.assetid: 96dcf5fe-4f3e-07b7-c248-46edd370fc31
ms.date: 06/08/2017
ms.openlocfilehash: b160d2add764d67146495f6f9a7821f3bc531b1c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplace-object-publisher"></a>Объект FindReplace (издатель)

Представляет условия для операции поиска. Свойства и методы объекта **FindReplace** соответствуют параметрам в диалоговом окне **Найти и заменить** .
 


## <a name="remarks"></a>Заметки

Если свойство **ReplaceScope** **pbReplaceScopeOne** или **pbReplaceScopeAll**, свойство **ReplaceWithText** должен иметь значение избежать текст из, заменен проверкой значение по умолчанию пустую **строку** для этого свойства.
 

 

## <a name="example"></a>Пример

Используйте свойство **Поиск** возвращает объект **FindReplace** . Следующий пример выделяет следующее вхождение слово «фабрики».
 

 

```
With ActiveDocument.Find 
 .Clear 
 .FindText = "factory" 
 .Execute 
End With
```

Присвойте свойству **ReplaceScope** для определения степени поиска. В следующем примере заменяется первого появления имя «Visual Basic Scripting Edition» с «VBScript».
 

 



```
With ActiveDocument.Find 
 .Clear 
 .FindText = "Visual Basic Scripting Edition" 
 .ReplaceWithText = "VBScript" 
 .ReplaceScope = pbReplaceScopeOne 
 .Execute 
End With
```

В следующем примере показано, как может осуществляться атрибуты шрифта FoundTextRange при **ReplaceScope** задано значение **pbReplaceScopeNone**.
 

 



```
Dim objFindReplace As FindReplace 
 
Set objFindReplace = ActiveDocument.Find 
With objFindReplace 
 .Clear 
 .FindText = "important" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 If .FoundTextRange.Font.Italic = msoFalse Then 
 .FoundTextRange.Font.Italic = msoTrue 
 End If 
 Loop 
End With
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Очистить](findreplace-clear-method-publisher.md)|
|[Выполнение](findreplace-execute-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](findreplace-application-property-publisher.md)|
|[FindText](findreplace-findtext-property-publisher.md)|
|[Вперед](findreplace-forward-property-publisher.md)|
|[FoundTextRange](findreplace-foundtextrange-property-publisher.md)|
|[MatchAlefHamza](findreplace-matchalefhamza-property-publisher.md)|
|[MatchCase](findreplace-matchcase-property-publisher.md)|
|[MatchDiacritics](findreplace-matchdiacritics-property-publisher.md)|
|[MatchKashida](findreplace-matchkashida-property-publisher.md)|
|[MatchWholeWord](findreplace-matchwholeword-property-publisher.md)|
|[MatchWidth](findreplace-matchwidth-property-publisher.md)|
|[Родительский раздел](findreplace-parent-property-publisher.md)|
|[ReplaceScope](findreplace-replacescope-property-publisher.md)|
|[ReplaceWithText](findreplace-replacewithtext-property-publisher.md)|

