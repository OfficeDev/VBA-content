---
title: "Метод FillFormat.UserPicture (издатель)"
keywords: vbapb10.chm2359319
f1_keywords: vbapb10.chm2359319
ms.prod: publisher
api_name: Publisher.FillFormat.UserPicture
ms.assetid: b1eaf724-42b4-657f-4d88-bc8547664893
ms.date: 06/08/2017
ms.openlocfilehash: 20c87215ec53dc604afd243d1ca3f6faa66bca67
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatuserpicture-method-publisher"></a>Метод FillFormat.UserPicture (издатель)

Заполняет указанные форму одно большое изображение.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UserPicture** ( **_PictureFile_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PictureFile|Обязательное свойство.| **String**|Имя файла изображения.|

## <a name="remarks"></a>Заметки

Для заливки фигуры с небольшой заголовков изображения, используйте метод **[UserTextured](fillformat-usertextured-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере добавляется два прямоугольника active публикацию. Область в левой части заполняется одно большое изображение с; прямоугольник справа заполняется множество небольших заголовков и то же изображение. (Обратите внимание на то, что действительный путь к файлу для работы этого примера необходимо заменить PathToFile.)


```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserPicture PictureFile:="PathToFile" 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserTextured TextureFile:="PathToFile" 
End With 

```


