---
title: "Свойство ColorCMYK.Black (издатель)"
keywords: vbapb10.chm2621442
f1_keywords: vbapb10.chm2621442
ms.prod: publisher
api_name: Publisher.ColorCMYK.Black
ms.assetid: f404ee43-45e1-6c6d-75cc-b868fdedcd86
ms.date: 06/08/2017
ms.openlocfilehash: 7d23349ff055dab9087a26875be879cdc2ff8e0b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorcmykblack-property-publisher"></a>Свойство ColorCMYK.Black (издатель)

Задает или возвращает значение типа **Long** , представляющее компонент черного цвета CMYK. Значение может быть любое число в диапазоне от 0 до 255. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Черный**

 переменная _expression_A, представляет собой объект- **ColorCMYK** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере создается два новых фигур и затем показана цвет заливки CMYK для одной формы и значения CMYK вторую фигуру на те же значения CMYK.


```vb
Sub ReturnAndSetCMYK() 
 Dim lngCyan As Long 
 Dim lngMagenta As Long 
 Dim lngYellow As Long 
 Dim lngBlack As Long 
 Dim shpHeart As Shape 
 Dim shpStar As Shape 
 
 Set shpHeart = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeHeart, Left:=100, _ 
 Top:=100, Width:=100, Height:=100) 
 Set shpStar = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=100, Width:=150, Height:=150) 
 
 With shpHeart.Fill.ForeColor.CMYK 
 .SetCMYK 10, 80, 200, 30 
 lngCyan = .Cyan 
 lngMagenta = .Magenta 
 lngYellow = .Yellow 
 lngBlack = .Black 
 End With 
 
 'Sets new shape to current shape's CMYK colors 
 shpStar.Fill.ForeColor.CMYK.SetCMYK _ 
 Cyan:=lngCyan, Magenta:=lngMagenta, _ 
 Yellow:=lngYellow, Black:=lngBlack 
End Sub
```


