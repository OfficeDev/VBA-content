---
title: "Метод Shapes.AddLabel (издатель)"
keywords: vbapb10.chm2162707
f1_keywords: vbapb10.chm2162707
ms.prod: publisher
api_name: Publisher.Shapes.AddLabel
ms.assetid: 5a803aa2-d37f-6da1-7d8b-58ee2dcd8146
ms.date: 06/08/2017
ms.openlocfilehash: 3cf3fac4f868b594ded4581ddc2a3019298efe90
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddlabel-method-publisher"></a>Метод Shapes.AddLabel (издатель)

Добавляет новый объект **[фигуры](shape-object-publisher.md)** , представляющее текстовой метки для указанной коллекции **[фигур](shapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddLabel** ( **_Ориентация_**, **_слева_**, **_Top_**, **_Ширина_**, **_Высота_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Ориентация|Обязательное свойство.| **PbTextOrientation**|Ориентация метку.|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющей текстовой метки.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющей текстовой метки.|
|Width|Обязательное свойство.| **Variant**|Ширина формы, представляющее текстовой метки.|
|Height|Обязательное свойство.| **Variant**|Высота фигуры, представляющей текстовой метки.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Аргументы слева, Top, ширину и высоту числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Параметр ориентации может иметь одно из следующих констант **PbTextOrientation** .



| **pbTextOrientationHorizontal**| Горизонтальная текстовой метки для языков слева направо. | | **pbTextOrientationRightToLeft**| Горизонтальная текстовой метки для языков для письма справа налево. | | **pbTextOrientationVerticalEastAsia**| Вертикальная текстовой метки для языков Восточной Азии. |

## <a name="example"></a>Пример

В следующем примере добавляется новой метки горизонтальный текст для первой страницы active публикации.


```vb
Dim shpLabel As Shape 
 
Set shpLabel = ActiveDocument.Pages(1).Shapes.AddLabel _ 
 (Orientation:=pbTextOrientationHorizontal, _ 
 Left:=144, Top:=144, _ 
 Width:=72, Height:=18)
```


