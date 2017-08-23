---
title: "Свойство Page.ReaderSpread (издатель)"
keywords: vbapb10.chm393238
f1_keywords: vbapb10.chm393238
ms.prod: publisher
api_name: Publisher.Page.ReaderSpread
ms.assetid: 32823d2d-4bcd-a5a6-1ad1-ca1035d4fdea
ms.date: 06/08/2017
ms.openlocfilehash: 98fbda8383dc35f911c478a38386badb7a321227
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagereaderspread-property-publisher"></a>Свойство Page.ReaderSpread (издатель)

Возвращает объект **[ReaderSpread](readerspread-object-publisher.md)** , представляющий распространения чтения на указанной странице.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReaderSpread**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

ReaderSpread


## <a name="example"></a>Пример

В этом примере проверяется, если ширина для указанной странице чтения включает в себя меньше, чем две страницы. Если это так, он изменяется ширина для включения две страницы чтения.


```vb
Sub SetFacingPages() 
 With ActiveDocument.Pages(2).ReaderSpread 
 If .PageCount < 2 Then _ 
 ActiveDocument.ViewTwoPageSpread = True 
 End With 
End Sub
```


