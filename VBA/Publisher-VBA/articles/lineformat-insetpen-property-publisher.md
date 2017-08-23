---
title: "Свойство LineFormat.InsetPen (издатель)"
keywords: vbapb10.chm3408148
f1_keywords: vbapb10.chm3408148
ms.prod: publisher
api_name: Publisher.LineFormat.InsetPen
ms.assetid: 955b152d-517f-b5fa-6e23-765ddeb41d46
ms.date: 06/08/2017
ms.openlocfilehash: f148a26d672ede3e795aad7bdd9d1b13d01f08b3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatinsetpen-property-publisher"></a>Свойство LineFormat.InsetPen (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, указанного фигуры отображаются ли линии внутри его границы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsetPen**

 переменная _expression_A, представляющий объект **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Ошибка возникает при попытке этому свойству присвоено значение **msoTrue** для любого автофигуры Microsoft Office, которая не поддерживает задание направления рисования.

Значение свойства **InsetPen** для таблиц всегда является **msoTrue**; При попытке установить свойство любые другие значения приводит к ошибке.

Значение свойства **InsetPen** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Линии непосредственно по границам указанного фигуры.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Линии внутри границы указанного фигуры.|

## <a name="example"></a>Пример

В следующем примере добавляется два прямоугольника страницу один из активных публикации первый с его линий внутри его границ, а второй — с помощью его линий на его границы.


```vb
Dim shpNew As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=150, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen = msoTrue 
 End With 
 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=300, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen = msoFalse 
 End With 
End With
```


