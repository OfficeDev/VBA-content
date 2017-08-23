---
title: "Метод Plates.FindPlateByInkName (издатель)"
keywords: vbapb10.chm2818053
f1_keywords: vbapb10.chm2818053
ms.prod: publisher
api_name: Publisher.Plates.FindPlateByInkName
ms.assetid: 4ebbc826-468b-7cd7-806e-056e4cbb488c
ms.date: 06/08/2017
ms.openlocfilehash: 50dcd61361f6a44b93f281f70ae89008bff6b9cf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="platesfindplatebyinkname-method-publisher"></a>Метод Plates.FindPlateByInkName (издатель)

Возвращает объект **формы** , который представляет формы имя указанного рукописного ввода.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindPlateByInkName** ( **_InkName_**)

 _expression_An выражение, возвращающее объект **формы** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|InkName|Обязательное свойство.| **PbInkName**|Указывает форму для возврата.|

### <a name="return-value"></a>Возвращаемое значение

Формы


## <a name="remarks"></a>Заметки

Параметр InkName может иметь одно из ** [PbInkName](http://msdn.microsoft.com/library/69e335b8-40b8-c984-84b6-64073a8ed7ab%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Publisher.

Процесс цвета назначены разные номера индекса в коллекции **формы** , чем в коллекции **PrintablePlates** . Используйте метод **FindPlateByInkName** , чтобы убедиться, что доступ к желаемую объекту **формы** или **PrintablePlate** .


## <a name="example"></a>Пример

Следующий пример возвращает свойства для формы, представляющее третий цвет, определенных для активной публикации.


```vb
Sub ListPlatePropertiesByInkName() 
Dim pplPlate As Plate 
 
 Set pplPlate = ActiveDocument.Plates.FindPlateByInkName(pbInkNameSpot3) 
 
 With pplPlate 
 Debug.Print "Plate Name: " &; .Name 
 Debug.Print "Index: " &; .Index 
 Debug.Print "Ink Name: " &; .InkName 
 Debug.Print "Color: " &; .Color 
 Debug.Print "Luminance: " &; .Luminance 
 Debug.Print "In Use?: " &; .InUse 
 End With 
End Sub
```


