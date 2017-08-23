---
title: "Свойство TextEffectFormat.RotatedChars (издатель)"
keywords: vbapb10.chm3735817
f1_keywords: vbapb10.chm3735817
ms.prod: publisher
api_name: Publisher.TextEffectFormat.RotatedChars
ms.assetid: 47566497-7b78-65dc-48d9-26b2e4245d31
ms.date: 06/08/2017
ms.openlocfilehash: 9772eb6c984df60680215d8f395b6d56459f57b7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatrotatedchars-property-publisher"></a>Свойство TextEffectFormat.RotatedChars (издатель)

 **msoTrue** при символов в указанном WordArt вращение 90 градусов относительно WordArt ограничивающего фигуры. **msoFalse** , если символы в указанном WordArt сохранить исходная ориентация относительно ограничивающего фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RotatedChars**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Если объект WordArt горизонтальный текст, для свойства **RotatedChars** значение **True,** поворот знаки 90 градусов против. Если объект WordArt вертикальной, для свойства **RotatedChars** значение **False** поворот знаки 90 градусов часовой. Используйте метод **[ToggleVerticalText](texteffectformat-toggleverticaltext-method-publisher.md)** для переключения между горизонтальных и вертикальных текста.

Метод **[Отразить](shape-flip-method-publisher.md)** и **[Вращение](shape-rotation-property-publisher.md)** свойство объекта **[Shape](shape-object-publisher.md)** **RotatedChars** свойство и метод **ToggleVerticalText** объекта **[TextEffectFormat](texteffectformat-object-publisher.md)** все влияет на ориентация символов и направление потока текста в объект **фигуры** , представляющий WordArt. Может потребоваться проверить, узнайте, как объединить эффекты этих свойств и методов для получения результатов, который будет.


## <a name="example"></a>Пример

В этом примере добавляется объект WordArt, который содержит текст «Test» для активной публикации и поворот знаки 90 градусов против.


```vb
Sub CreateFormatWordArt() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test", FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=10, Top:=10) 
 .TextEffect.RotatedChars = msoTrue 
 End With 
End Sub
```


