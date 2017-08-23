---
title: "Свойство ThreeDFormat.PresetMaterial (издатель)"
keywords: vbapb10.chm3801351
f1_keywords: vbapb10.chm3801351
ms.prod: publisher
api_name: Publisher.ThreeDFormat.PresetMaterial
ms.assetid: 5f12fb22-f596-0d59-1f02-63ce8d4bd927
ms.date: 06/08/2017
ms.openlocfilehash: 696c6288c8d848d1fb56914cb16cfebdc3a5db1c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatpresetmaterial-property-publisher"></a>Свойство ThreeDFormat.PresetMaterial (издатель)

Возвращает или задает константой **MsoPresetMaterial** , представляющий материал поверхности придания объема. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetMaterial**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetMaterial


## <a name="remarks"></a>Заметки

Значение свойства **PresetMaterial** может иметь одно из ** [MsoPresetMaterial](http://msdn.microsoft.com/library/4cf62ef4-f6c8-eb0c-1dfd-569aafca16c0%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере указывает поверхности придания объема для фигуры одно в активной публикации каркас. В данном примере для работы указанного фигуры должен быть объемной фигуры.


```vb
Sub SetExtrusionMaterial() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .PresetMaterial = msoMaterialWireFrame 
 End With 
End Sub
```


