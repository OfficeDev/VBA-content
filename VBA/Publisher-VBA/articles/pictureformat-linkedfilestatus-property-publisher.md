---
title: "Свойство PictureFormat.LinkedFileStatus (издатель)"
keywords: vbapb10.chm3604787
f1_keywords: vbapb10.chm3604787
ms.prod: publisher
api_name: Publisher.PictureFormat.LinkedFileStatus
ms.assetid: 43ddffe3-9cc3-b102-c5e8-80f26f63849c
ms.date: 06/08/2017
ms.openlocfilehash: edc92477d0c48e21b898dac8019334e23322e90c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatlinkedfilestatus-property-publisher"></a>Свойство PictureFormat.LinkedFileStatus (издатель)

Возвращает константу **PbLinkedFileStatus** , указывающий состояние файла, связанных с указанным рисунков. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LinkedFileStatus**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbLinkedFileStatus


## <a name="remarks"></a>Заметки

Это свойство применяется только к файлам связанного рисунка. Возвращает «Отказано в разрешении» для фигуры, представляющие внедренных или вставленного изображения.

Используйте один из следующих свойств для определения, является ли фигура представляет связанного рисунка:


-  Свойство **[Type](shape-type-property-publisher.md)** объекта **[фигуры](shape-object-publisher.md)**
    
- Свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)**
    


Значение свойства **LinkedFileStatus** может иметь одно из **[PbLinkedFileStatus](pblinkedfilestatus-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующий пример создает список связанных рисунков в активной публикации, для которого не удается найти связанные файлы.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .LinkedFileStatus = pbLinkedFileMissing Then 
 Debug.Print .Filename 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 

```


