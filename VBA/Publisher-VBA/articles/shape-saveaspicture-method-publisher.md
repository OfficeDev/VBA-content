---
title: "Метод Shape.SaveAsPicture (издатель)"
keywords: vbapb10.chm2228375
f1_keywords: vbapb10.chm2228375
ms.prod: publisher
api_name: Publisher.Shape.SaveAsPicture
ms.assetid: 2cc18a83-b947-ca8c-eab4-71a03b79b82b
ms.date: 06/08/2017
ms.openlocfilehash: ee46ba06b9ea275566371772f64070e664f153ff
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaveaspicture-method-publisher"></a>Метод Shape.SaveAsPicture (издатель)

Сохраняет одну как файл рисунка.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SaveAsPicture** ( **_Имя файла_**, **_pbResolution_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Путь и имя нового файла изображения, который требуется создать. Рисунок сохраняется в формат графики определяется по расширению имени файла (например, JPG или GIF) укажите.|
|pbResolution|Необязательный| **PbPictureResolution**|Разрешение, в которой будут рисунок сохраняется. Возможные значения см.|

## <a name="remarks"></a>Заметки

Возможные значения для параметра pbResolution объявляются в перечислении **[PbPictureResolution](pbpictureresolution-enumeration-publisher.md)** в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **SaveAsPicture** для сохранения первой фигуры в коллекции фигур на первой странице активная публикация качестве рисунков JPG-файла.

Перед запуском этого кода замените _filename.jpg_ допустимое имя файла и путь к папке на компьютере, где у вас есть разрешение на сохранение файлов.




```vb
Public Sub SaveAsPicture_Example() 
 
 ThisDocument.Pages(1).Shapes(1).SaveAsPicture "filename.jpg" 
 
End Sub
```


