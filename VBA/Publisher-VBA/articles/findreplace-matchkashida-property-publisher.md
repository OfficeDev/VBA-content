---
title: "Свойство FindReplace.MatchKashida (издатель)"
keywords: vbapb10.chm8323082
f1_keywords: vbapb10.chm8323082
ms.prod: publisher
api_name: Publisher.FindReplace.MatchKashida
ms.assetid: ec2b5fa0-0549-b5c2-d8b9-666be1cbe193
ms.date: 06/08/2017
ms.openlocfilehash: 2522d2be2e5051f2ec35a71f12e98673b27643d0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacematchkashida-property-publisher"></a>Свойство FindReplace.MatchKashida (издатель)

Задает или возвращает значение **типа Boolean** представляющее ли операция поиска будет соответствовать кашиды. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MatchKashida**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство не может быть доступно в зависимости от языка, включен в операционной системе. Значение по умолчанию — **False**.

Возвращает ** отказано в доступе ** Если арабский не включено.


## <a name="example"></a>Пример

В этом примере выполняется поиск первого появления слово «» арабского языка в документе в формате соответствия кашиды.


```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "" 
 .MatchKashida = True 
 .Execute 
End With 

```

В этом примере исходя из предыдущей за исключением того, что кашиды не совпадать. Таким образом слова «» или «», оба доступны из-за кашиды будет игнорироваться.




```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "مــــحـــمـــــد" 
 .MatchKashida = False 
 .Execute 
End With 

```


