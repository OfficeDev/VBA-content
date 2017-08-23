---
title: "Свойство TextEffectFormat.PresetTextEffect (издатель)"
keywords: vbapb10.chm3735816
f1_keywords: vbapb10.chm3735816
ms.prod: publisher
api_name: Publisher.TextEffectFormat.PresetTextEffect
ms.assetid: d7ef0995-4560-fea0-b98d-03c8e0c8e65e
ms.date: 06/08/2017
ms.openlocfilehash: b50f96364e3565cd5829961c68dc3ed43ec370e8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatpresettexteffect-property-publisher"></a>Свойство TextEffectFormat.PresetTextEffect (издатель)

Возвращает или задает константой **MsoPresetTextEffect** , представляющий стиль указанного WordArt. Значения для этого свойства соответствуют форматов в диалоговом окне **Коллекция WordArt** нумерованные слева направо, сверху вниз. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetTextEffect**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetTextEffect


## <a name="remarks"></a>Заметки

Значение свойства **PresetTextEffect** может иметь одно из ** [MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере задается эффекта стиля текста для первой фигуры на первой странице active публикации. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.


```vb
Sub ChangeTextEffect() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = msoTextEffect Then 
 .TextEffect.PresetTextEffect = msoTextEffect1 
 End If 
 End With 
End Sub
```


