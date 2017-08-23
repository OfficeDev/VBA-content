---
title: "Метод ShapeRange.SaveAsPicture (издатель)"
keywords: vbapb10.chm2294050
f1_keywords: vbapb10.chm2294050
ms.prod: publisher
api_name: Publisher.ShapeRange.SaveAsPicture
ms.assetid: 0be9b741-8f11-a386-313b-231a3269883a
ms.date: 06/08/2017
ms.openlocfilehash: 3d01b533f5822e9303523fe5d0a299f69ccc54f7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangesaveaspicture-method-publisher"></a>Метод ShapeRange.SaveAsPicture (издатель)

Сохраняет диапазон из одного или нескольких фигур в файле.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SaveAsPicture** ( **_Имя файла_**, **_pbResolution_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Путь и имя нового файла изображения, который требуется создать. Рисунок сохраняется в формат графики определяется по расширению имени файла (например, JPG или GIF) укажите.|
|pbResolution|Необязательный| **PbPictureResolution**|Разрешение, в которой будут рисунок сохраняется. Возможные значения см.|

## <a name="remarks"></a>Заметки

Возможные значения для параметра pbResolution объявляются в перечислении **[PbPictureResolution](pbpictureresolution-enumeration-publisher.md)** в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **SaveAsPicture** для сохранения всех фигур на первой странице активная публикация качестве рисунков JPG-файла.

Перед запуском этого кода замените _filename.jpg_ допустимое имя файла и путь к папке на компьютере, где у вас есть разрешение на сохранение файлов.




```vb
Public Sub SaveAsPicture_Example() 
 
 Dim pubShapeRange As Publisher.ShapeRange 
 Set pubShapeRange = ThisDocument.Pages(1).Shapes.Range 
 
 pubShapeRange.SaveAsPicture "filename.jpg" 
 
End Sub
```


