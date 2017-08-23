---
title: "Свойство FindReplace.MatchWidth (издатель)"
keywords: vbapb10.chm8323084
f1_keywords: vbapb10.chm8323084
ms.prod: publisher
api_name: Publisher.FindReplace.MatchWidth
ms.assetid: b9f89092-6ac0-bbf9-4bfd-d3cce2359b80
ms.date: 06/08/2017
ms.openlocfilehash: f5063b1583df840d9769075e81d86fd448127487
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacematchwidth-property-publisher"></a>Свойство FindReplace.MatchWidth (издатель)

Задает или возвращает значение **типа Boolean** представляющее ли операции поиска будут учитывать ширину знаков поиск текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MatchWidth**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство не может быть доступно в зависимости от языка, включен в операционной системе. Значение по умолчанию — **False**.

Возвращает «Доступ запрещен», если не включена восточно-азиатских языков.


## <a name="example"></a>Пример

В следующем примере выполняется поиск каждого экземпляры слово «ширина» в активном документе и применяет жирное форматирование. Свойство **MatchWidth** имеет значение **False** , поэтому полное или половинной ширины символов, оба доступны. Например, поиска будет быстрого форматирования word «width» (половинной ширины знаков) и слово "w я h d t» (полной ширины знаков).


```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "width" 
 .MatchWidth = False 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With
```

В следующем примере выполняется поиск каждого экземпляры слово «ширина» в активном документе и применяет жирное форматирование. Свойство **MatchWidth** имеет значение **True** , поэтому будет найден полное или знаков половинной ширины. Например поиск будет быстрого форматирования для «ширина». Он не будет применять форматирование к слово «w я h d t».




```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "width" 
 .MatchWidth = True 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With
```


