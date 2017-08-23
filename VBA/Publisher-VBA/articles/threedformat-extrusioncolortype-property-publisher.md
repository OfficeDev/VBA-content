---
title: "Свойство ThreeDFormat.ExtrusionColorType (издатель)"
keywords: vbapb10.chm3801346
f1_keywords: vbapb10.chm3801346
ms.prod: publisher
api_name: Publisher.ThreeDFormat.ExtrusionColorType
ms.assetid: 5abd895d-0cf3-985d-537e-e45d02f8a852
ms.date: 06/08/2017
ms.openlocfilehash: a850644e6ba238dfb14b10680afc211d027b5507
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatextrusioncolortype-property-publisher"></a>Свойство ThreeDFormat.ExtrusionColorType (издатель)

Возвращает или задает константой **MsoExtrusionColorType** , указывающее, является ли цвет объемной фигуры основано на вытянутый фигуры заливки (лицевой из изменяется) и устанавливается автоматически при изменении заливки фигуры или ли цвет объемной фигуры не зависит от заливки фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ExtrusionColorType**

 переменная _expression_A, представляющий объект **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoExtrusionColorType


## <a name="remarks"></a>Заметки

Значение свойства **ExtrusionColorType** может иметь одно из ** [MsoExtrusionColorType](http://msdn.microsoft.com/library/6acf7f2b-3d7b-15e3-f468-7dcb20865dc1%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

Если первой фигуры в активной публикации цвет автоматического объемной фигуры, в этом примере позволяет выбирать желтый цвет.


```vb
With ActiveDocument.Pages(1).Shapes(1).ThreeD 
    If .ExtrusionColorType = msoExtrusionColorAutomatic Then 
        .ExtrusionColor.RGB = RGB(240, 235, 16) 
    End If 
End With 

```


