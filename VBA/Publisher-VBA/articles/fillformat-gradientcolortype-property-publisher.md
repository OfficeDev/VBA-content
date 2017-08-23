---
title: "Свойство FillFormat.GradientColorType (издатель)"
keywords: vbapb10.chm2359554
f1_keywords: vbapb10.chm2359554
ms.prod: publisher
api_name: Publisher.FillFormat.GradientColorType
ms.assetid: b0866675-4bc4-5e82-780d-8bae4b7d16ef
ms.date: 06/08/2017
ms.openlocfilehash: e223fe955af45db80fb9e90350200d5eb0868b2f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatgradientcolortype-property-publisher"></a>Свойство FillFormat.GradientColorType (издатель)

Возвращает константу **MsoGradientColorType** , указывающий тип градиента для указанного заполнения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GradientColorType**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoGradientColorType


## <a name="remarks"></a>Заметки

Используйте метод [OneColorGradient](fillformat-onecolorgradient-method-publisher.md), [PresetGradient](fillformat-presetgradient-method-publisher.md)или **[TwoColorGradient](fillformat-twocolorgradient-method-publisher.md)** , чтобы задать тип градиента для заполнения.

Значение свойства **GradientColorType** может иметь одно из ** [MsoGradientColorType](http://msdn.microsoft.com/library/0940fc83-d089-8b1d-dcf1-73773d0e21c5%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере изменяется заливки для всех фигур на первой странице active публикации, имеющие градиентной заливки двух цветов для предварительно градиентной заливки.


```vb
Dim shpLoop As Shape 
 
' Loop through collection of shapes. 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 ' Test for two-color gradient. 
 If .GradientColorType = msoGradientTwoColors Then 
 ' Apply a preset gradient. 
 .PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientBrass 
 End If 
 End With 
Next shpLoop 

```


