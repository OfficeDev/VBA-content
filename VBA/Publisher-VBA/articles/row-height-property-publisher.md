---
title: "Свойство Row.Height (издатель)"
keywords: vbapb10.chm4849667
f1_keywords: vbapb10.chm4849667
ms.prod: publisher
api_name: Publisher.Row.Height
ms.assetid: fd243edc-1521-bd67-a365-2c4685ee5037
ms.date: 06/08/2017
ms.openlocfilehash: 9419394737f46489fd414900e643a6591b8e0910
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rowheight-property-publisher"></a>Свойство Row.Height (издатель)

Возвращает или задает **Variant** , который представляет высоту (в точках) строки указанной таблицы или фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Высота**

 переменная _expression_A, представляет собой объект- **строку** .


## <a name="remarks"></a>Заметки

Допустимые значения для свойства **Height** зависит от размера рабочей области приложения и позиции объекта в рабочей области. По центру объектов на размер страницы не баннер свойство **Height** может быть 0,0-50,0 дюйма. По центру объектов на размер заголовка страницы свойство **Height** может быть 0.0 для 241.0 дюйма.


## <a name="example"></a>Пример

В этом примере создается новая таблица и задает высоту и ширину второй строк и столбцов, соответственно.


```vb
Sub SetRowHeightColumnWidth() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=3, _ 
 NumColumns:=3, Left:=80, Top:=80, Width:=400, Height:=12).Table 
 .Rows(2).Height = 72 
 .Columns(2).Width = 72 
 End With 
End Sub
```


