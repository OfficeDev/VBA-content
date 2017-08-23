---
title: "Свойство PictureFormat.TransparentBackground (издатель)"
keywords: vbapb10.chm3604744
f1_keywords: vbapb10.chm3604744
ms.prod: publisher
api_name: Publisher.PictureFormat.TransparentBackground
ms.assetid: 0a78b579-92bf-36e6-22f6-3ca0a48f5b5a
ms.date: 06/08/2017
ms.openlocfilehash: 961293cb4074d2e4fd56c85e112a5451fec389b7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformattransparentbackground-property-publisher"></a>Свойство PictureFormat.TransparentBackground (издатель)

Указывает, отображение прозрачной части указанного изображения, которые определены как прозрачный. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TransparentBackground**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **TransparentBackground** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Части рисунка, цвет которого является прозрачность цвета не отображаются прозрачной.|
| **msoTriStateMixed**|Возвращает только значение, указывающее, сочетание **msoTrue** и **msoFalse** для указанных объектов...|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**| Части рисунка, цвет которого является прозрачность цвет прозрачным.|
Свойство **[прозрачного цвета](pictureformat-transparencycolor-property-publisher.md)** прозрачный цвет.

Это свойство применяется только к растровых изображений.

Если необходимо иметь возможность видеть через прозрачный частей рисунка до объекты за изображение, необходимо присвоить **mso False**свойство **[Visible](fillformat-visible-property-publisher.md)** объекта **[FillFormat](fillformat-object-publisher.md)** рисунков. Если изображение имеет прозрачный цвет, свойство **Visible** объекта **FillFormat** изображение задано значение **msoTrue**, отображается с помощью прозрачный цвет заливки рисунка, но объектов, изображение, скрываются.


## <a name="example"></a>Пример

В этом примере задается цвет синий как прозрачный для фигуры одно в активной публикации. Для обеспечения работы примера фигуры один должен быть растрового изображения.


```vb
With ActiveDocument.Pages(1).Shapes(1) 
 
 With .PictureFormat 
 .TransparentBackground = msoTrue 
 ' RGB(0, 0, 255) is the color blue. 
 .TransparencyColor = RGB(0, 0, 255) 
 End With 
 
 .Fill.Visible = False 
 
End With 

```


