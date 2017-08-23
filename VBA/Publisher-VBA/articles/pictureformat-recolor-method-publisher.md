---
title: "Метод PictureFormat.Recolor (издатель)"
keywords: vbapb10.chm3604793
f1_keywords: vbapb10.chm3604793
ms.prod: publisher
api_name: Publisher.PictureFormat.Recolor
ms.assetid: 42bc2280-b6d0-862a-7118-38ec1513b9c7
ms.date: 06/08/2017
ms.openlocfilehash: 7814f2afaf49174ff67f5c451bca81aa26e1f620
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatrecolor-method-publisher"></a>Метод PictureFormat.Recolor (издатель)

Изменение цвета изображения в публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Изменение цвета** ( **_Цвет_**, **_LeaveBlackPartsBlack_**)

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Цвет|Обязательное свойство.| **ColorFormat**|Цвет, используемый для Перекрашивание.|
|LeaveBlackPartsBlack|Обязательное свойство.| **MsoTriState**| **Значение true,** Если все части исходного изображения были черный цвет следует оставить черные.|

## <a name="remarks"></a>Заметки

**Изменение цвета** , соответствует параметры, доступные в диалоговом окне **Перекрашивание рисунков** . (В меню **Формат** выберите пункт **изображение**и нажмите кнопку **изменить цвет**)


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **Перекрашивание** для изменения цвета изображения. Он recolors первой фигуры в коллекции **фигур** на первой странице публикации. После выполнения кода с помощью метода **[RestoreOriginalColors](pictureformat-restoreoriginalcolors-method-publisher.md)** можно восстановить исходные цвета.

В данном примере для работы фигуры к перекрашиванию значения изображения или объекта, который представляет изображение.




```vb
Public Sub Recolor_Example() 
 
 Dim pubPictureFormat As Publisher.PictureFormat 
 Dim pubShape As Publisher.Shape 
 Dim pubColorFormat As Publisher.ColorFormat 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 Set pubPictureFormat = pubShape.PictureFormat 
 Set pubColorFormat = pubShape.Fill.BackColor 
 
 pubPictureFormat.Recolor pubColorFormat, msoTrue 
 
End Sub
```


