---
title: "Свойство PictureFormat.OriginalResolution (издатель)"
keywords: vbapb10.chm3604776
f1_keywords: vbapb10.chm3604776
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalResolution
ms.assetid: 0cb7ee4e-3eb8-baee-6535-d936e3c5f05c
ms.date: 06/08/2017
ms.openlocfilehash: d054be6b2ac9f58fd19b349539735736c0c27f37
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalresolution-property-publisher"></a>Свойство PictureFormat.OriginalResolution (издатель)

Возвращает значение типа **времени** , представляющий, точек на дюйм (т/д), решение, изначально сканируемые связанного рисунка. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalResolution**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанных рисунков. Возвращает значение «Отказано в разрешении» для фигуры, представляющие внедренные или вставлять рисунки.

Чтобы определить, является ли фигура представляет связанного рисунка, используйте свойство **[Type](shape-type-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** или свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** .

Свойство **[EffectiveResolution](pictureformat-effectiveresolution-property-publisher.md)** используется для определения разрешения, в котором этот рисунок или объект OLE печатает в указанный документ.


## <a name="example"></a>Пример

В следующем примере проверяется каждого изображения в активной публикации и возвращает свойства выбранного изображения для изображений, которые связаны.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Resolution in Publication: " &; .EffectiveResolution &; " dpi" 
 Debug.Print "Original Resolution: " &; .OriginalResolution &; " dpi" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


