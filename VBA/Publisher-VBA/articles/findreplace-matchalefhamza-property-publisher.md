---
title: "Свойство FindReplace.MatchAlefHamza (издатель)"
keywords: vbapb10.chm8323079
f1_keywords: vbapb10.chm8323079
ms.prod: publisher
api_name: Publisher.FindReplace.MatchAlefHamza
ms.assetid: a8bdfbc3-13b5-e6a1-d86c-95e8f58ec263
ms.date: 06/08/2017
ms.openlocfilehash: 043d840a5a6e7b7c36d84b462a45706468d9094c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacematchalefhamza-property-publisher"></a>Свойство FindReplace.MatchAlefHamza (издатель)

Задает или возвращает **логическое значение** , указывающее операцию поиска будет соответствовать этого и гамза. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MatchAlefHamza**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство не может быть доступно в зависимости от языка, включен в операционной системе. Значение по умолчанию — **False**.

Возвращает **Отказано в доступе** , если арабский не включено.


## <a name="example"></a>Пример

В этом примере выполняется поиск первого появления слово «» арабского языка в документе в формате сопоставления этого и гамза.


```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "" 
 .MatchAlefHamza = True 
 .Execute 
End With 

```

В этом примере исходя из предыдущего, за исключением того, что алиф гамза не совпадать. Таким образом слова «» или «», оба доступны из-за этого и гамза будет игнорироваться.




```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "" 
 .MatchAlefHamza = False 
 .Execute 
End With 

```


