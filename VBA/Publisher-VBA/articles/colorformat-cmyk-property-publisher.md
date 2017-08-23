---
title: "Свойство ColorFormat.CMYK (издатель)"
keywords: vbapb10.chm2555907
f1_keywords: vbapb10.chm2555907
ms.prod: publisher
api_name: Publisher.ColorFormat.CMYK
ms.assetid: 28d7ad65-c63c-3b11-3ecc-c77a1a586b84
ms.date: 06/08/2017
ms.openlocfilehash: c78ea82073600070dc45b6769ba26bbfa3f4980f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformatcmyk-property-publisher"></a>Свойство ColorFormat.CMYK (издатель)

Возвращает объект **ColorCMYK** , который представляет свойства цвета CMYK.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Формат CMYK**

 переменная _expression_A, представляет собой объект- **ColorFormat** .


### <a name="return-value"></a>Возвращаемое значение

ColorCMYK


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


