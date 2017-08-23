---
title: "Свойство Options.PathForPictures (издатель)"
keywords: vbapb10.chm1048596
f1_keywords: vbapb10.chm1048596
ms.prod: publisher
api_name: Publisher.Options.PathForPictures
ms.assetid: e66c8c86-f049-0f32-0a0d-60fd37470708
ms.date: 06/08/2017
ms.openlocfilehash: 4fd6be2884b089875a54f64afa34624f8609bdbe
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionspathforpictures-property-publisher"></a>Свойство Options.PathForPictures (издатель)

Возвращает **строку** , представляющую путь по умолчанию для файлов рисунков. Чтение.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PathForPictures**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере устанавливает путь по умолчанию для файлов рисунков в строку, а затем использует строку пути для добавления указанного файла для активной публикации. (Обратите внимание на то, имя файла, заменены допустимое имя файла для работы этого примера).


```vb
Sub InsertNewPicture() 
 Dim strPicPath As String 
 
 strPicPath = Options.PathForPictures 
 
 ActiveDocument.Pages(1).Shapes.AddPicture FileName:=strPicPath _ 
 &; "Filename", LinktoFile:=msoFalse, _ 
 SaveWithDocument:=msoTrue, Left:=50, Top:=50, Height:=200 
 
End Sub
```


