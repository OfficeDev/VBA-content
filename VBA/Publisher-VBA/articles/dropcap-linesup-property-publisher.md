---
title: "Свойство DropCap.LinesUp (издатель)"
keywords: vbapb10.chm5505031
f1_keywords: vbapb10.chm5505031
ms.prod: publisher
api_name: Publisher.DropCap.LinesUp
ms.assetid: 97bf3fc1-2203-d916-0c2d-352260c279fe
ms.date: 06/08/2017
ms.openlocfilehash: efa66ff97962ae7cfab7422e6f81d5fee6193b45
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcaplinesup-property-publisher"></a>Свойство DropCap.LinesUp (издатель)

Возвращает или задает типа **Long** , представляющее номер строки возникновения потерянных заглавной буквы выше строки текста, на котором она существует. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LinesUp**

 переменная _expression_A, представляет собой объект- **буквицу** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере создается настраиваемых буквицы пять строк высокой и вызывает его две строки над строкой, на котором она существует.


```vb
Sub RaisedDropCap() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 With .DropCap 
 .Size = 5 
 .LinesUp = 2 
 End With 
 End With 
End Sub
```


