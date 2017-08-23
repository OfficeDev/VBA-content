---
title: "Свойство PictureFormat.FileSize (издатель)"
keywords: vbapb10.chm3604757
f1_keywords: vbapb10.chm3604757
ms.prod: publisher
api_name: Publisher.PictureFormat.FileSize
ms.assetid: 8bad7bc0-7381-9bd8-3db8-5841e41ccb34
ms.date: 06/08/2017
ms.openlocfilehash: 230a196706401a004e253134c299b6542ca65d0c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatfilesize-property-publisher"></a>Свойство PictureFormat.FileSize (издатель)

Возвращает значение типа **Long** , представляющее, в байтах, размер изображения или объекта OLE, как оно отображается в указанной публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Размер файла**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Свойство **[OriginalFileSize](pictureformat-originalfilesize-property-publisher.md)** для определения размера связанный файл, если связанного рисунка или объекта OLE.

Чтобы определить, является ли фигура представляет связанного рисунка, используйте свойство **[Type](shape-type-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** или свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** .


## <a name="example"></a>Пример

В следующем примере проверяется каждого изображения в активной публикации и печатает свойства выбранного изображения для изображений, которые связаны.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Original File Size: " &; .OriginalFileSize &; " bytes" 
 Debug.Print "File size in publication: " &; .FileSize &; " bytes" 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


