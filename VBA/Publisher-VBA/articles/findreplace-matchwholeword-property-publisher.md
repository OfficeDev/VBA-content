---
title: "Свойство FindReplace.MatchWholeWord (издатель)"
keywords: vbapb10.chm8323083
f1_keywords: vbapb10.chm8323083
ms.prod: publisher
api_name: Publisher.FindReplace.MatchWholeWord
ms.assetid: 512d37bc-c900-ee17-8a8e-5875db0a2f85
ms.date: 06/08/2017
ms.openlocfilehash: 8fe8a6516d8cf29b09ecfc05c4609284954aebbf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacematchwholeword-property-publisher"></a>Свойство FindReplace.MatchWholeWord (издатель)

Задает или возвращает значение **типа Boolean** , указывающий ли слово целиком, будут сопоставлены в операции поиска. Чтение и запись. **Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MatchWholeWord**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Значение по умолчанию для **MatchWholeWord** имеет **значение False**.


## <a name="example"></a>Пример

В этом примере будет выберите все вхождения слова «фактов» и быстрого форматирования.


```vb
With ActiveDocument.Find 
 .Clear 
 .MatchWholeWord = True 
 .FindText = "fact" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With 

```

В этом примере следует предыдущего примера, за исключением того, что целых слов не совпадать. Таким образом word «факт» в word «фабрики» или «factoid» будут иметь применяется полужирным шрифтом.




```vb
With ActiveDocument.Find 
 .Clear 
 .MatchWholeWord = False 
 .FindText = "fact" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With 

```


