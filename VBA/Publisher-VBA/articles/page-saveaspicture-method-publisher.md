---
title: "Метод Page.SaveAsPicture (издатель)"
keywords: vbapb10.chm393272
f1_keywords: vbapb10.chm393272
ms.prod: publisher
api_name: Publisher.Page.SaveAsPicture
ms.assetid: 9b118126-e072-9516-9863-14ea60264f01
ms.date: 06/08/2017
ms.openlocfilehash: 82e9f02ffd4ab02d3176dcc589fabe728e8a4690
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesaveaspicture-method-publisher"></a>Метод Page.SaveAsPicture (издатель)

Сохраняет страницу как файл рисунка.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SaveAsPicture** ( **_Имя файла_**, **_pbResolution_**)

 переменная _expression_A, представляющий объект **Page** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Путь и имя нового файла изображения, который требуется создать. Рисунок сохраняется в формат графики определяется по расширению имени файла (например, JPG или GIF) укажите.|
|pbResolution|Необязательный| **PbPictureResolution**|Разрешение, в которой будут рисунок сохраняется. Возможные значения см.|

## <a name="remarks"></a>Заметки

Возможные значения для параметра pbResolution объявляются в перечислении **[PbPictureResolution](pbpictureresolution-enumeration-publisher.md)** в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **SaveAsPicture** для сохранения первой страницы публикации, активных качестве рисунков JPG-файла.

Перед запуском этого кода замените _filename.jpg_ допустимое имя файла и путь к папке на компьютере, где у вас есть разрешение на сохранение файлов.




```vb
Public Sub SaveAsPicture_Example() 
 
 ThisDocument.Pages(1).SaveAsPicture "filename.jpg" 
 
End Sub
```


