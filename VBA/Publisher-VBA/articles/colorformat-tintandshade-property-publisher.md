---
title: "Свойство ColorFormat.TintAndShade (издатель)"
keywords: vbapb10.chm2555912
f1_keywords: vbapb10.chm2555912
ms.prod: publisher
api_name: Publisher.ColorFormat.TintAndShade
ms.assetid: 1c4897e0-ac55-08a8-8c43-dbd25d097ecc
ms.date: 06/08/2017
ms.openlocfilehash: ab8c87eaeffcf81637d226e8760e7e0c55657f7e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformattintandshade-property-publisher"></a>Свойство ColorFormat.TintAndShade (издатель)

Возвращает или задает **единого** , представляющий добавляемого освещения или затемнения цвета указанного фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TintAndShade**

 переменная _expression_A, представляет собой объект- **ColorFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

От -1 (самый темный) можно ввести номер 1 (очень светлый) для свойства **TintAndShade** нуль (0), нейтральный.


## <a name="example"></a>Пример

В этом примере создается новая форма в активном документе, задает цвет заливки и осветляет цветом.


```vb
Sub NewTintedShape() 
 Dim shpHeart As Shape 
 Set shpHeart = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=150, _ 
 Top:=150, Width:=250, Height:=250) 
 With shpHeart.Fill.ForeColor 
 .CMYK.SetCMYK Cyan:=255, Magenta:=28, Yellow:=0, Black:=0 
 .TintAndShade = 0.3 
 End With 
End Sub
```


