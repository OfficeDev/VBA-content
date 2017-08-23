---
title: "Свойство LineFormat.Style (издатель)"
keywords: vbapb10.chm3408144
f1_keywords: vbapb10.chm3408144
ms.prod: publisher
api_name: Publisher.LineFormat.Style
ms.assetid: 3826eb43-b90e-e24b-31d5-8d9eddd3ed4e
ms.date: 06/08/2017
ms.openlocfilehash: 004ab0354fa754fdff7174b0563e5bb47db75b3c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatstyle-property-publisher"></a>Свойство LineFormat.Style (издатель)

Возвращает или задает константой **MsoLineStyle** , представляющий стиль строку, чтобы применить к фигуры или границы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Стиль**

 переменная _expression_A, представляет собой объект- **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoLineStyle


## <a name="remarks"></a>Заметки

Значение свойства **Style** может иметь одно из **MsoLineStyle** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



| **msoLineSingle**|| **msoLineStyleMixed**|| **msoLineThickBetweenThin**|| **msoLineThickThin**|| **msoLineThinThick**|| **msoLineThinThin**|

## <a name="example"></a>Пример

В этом примере добавляется новый фигура и устанавливаются свойства линии для фигуры.


```vb
Sub SetLineStyle() 
 With ActiveDocument.Pages(1).Shapes.AddShape(msoShapeRectangle, _ 
 Left:=72, Top:=140, Width:=200, Height:=100) 
 .Rotation = 120 
 With .Line 
 .Weight = 5 
 .DashStyle = msoLineDashDotDot 
 .Style = msoLineThickBetweenThin 
 End With 
 End With 
End Sub
```


