---
title: "Метод Font.Duplicate (издатель)"
keywords: vbapb10.chm5373992
f1_keywords: vbapb10.chm5373992
ms.prod: publisher
api_name: Publisher.Font.Duplicate
ms.assetid: 26ae64bc-036e-5c19-cbac-99f11da7fb60
ms.date: 06/08/2017
ms.openlocfilehash: 72b996b8de044d19a51d325153c2356624f205b7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontduplicate-method-publisher"></a>Метод Font.Duplicate (издатель)

Создает копию на указанный объект **[шрифта](font-object-publisher.md)** и возвращает новый объект **Font** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Дублирующиеся**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Font


## <a name="example"></a>Пример

В следующем примере дублирует форматирование информации из диапазона текст в фигуре одно на странице символов, один из активных публикации и применяется к диапазон текста в форме двух.


```vb
Dim fntTemp As Font 
 
With ActiveDocument.Pages(1) 
 Set fntTemp = _ 
 .Shapes(1).TextFrame.TextRange.Font.Duplicate 
 .Shapes(2).TextFrame.TextRange.Font = fntTemp 
End With
```


