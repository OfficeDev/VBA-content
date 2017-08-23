---
title: "Метод PictureFormat.Replace (издатель)"
keywords: vbapb10.chm3604786
f1_keywords: vbapb10.chm3604786
ms.prod: publisher
api_name: Publisher.PictureFormat.Replace
ms.assetid: b2bce79a-5c46-1473-601d-a4a25176edeb
ms.date: 06/08/2017
ms.openlocfilehash: 77e5cdf9a24657536db79ac711a87f517dac30d2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatreplace-method-publisher"></a>Метод PictureFormat.Replace (издатель)

Заменяет указанный рисунок. Возвращает **значение Nothing**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Замена** ( **_Pathname_**, **_InsertAs_**)

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя пути|Обязательное свойство.| **String**|Имя и путь к файлу, с которым вы хотите заменить указанный рисунок.|
|InsertAs|Необязательный| **PbPictureInsertAs**|Так, в которой будет вставлено в документ файл рисунка: связанные или внедренные.|

## <a name="remarks"></a>Заметки

Используйте метод **Replace** для обновления связанных рисунков файлы, которые были изменены с момента они были добавлены в документ. Используйте свойство **[LinkedFileStatus](pictureformat-linkedfilestatus-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** для определения, были ли изменены связанного рисунка.

Параметр InsertAs может иметь одно из следующих **PbPictureInsertAs** константы, описанные в библиотеке типов, Microsoft Publisher. значение по умолчанию — **pbPictureInsertAsOriginalState**.



| **pbPictureInsertAsEmbedded**|| **pbPictureInsertAsLinked**|| **pbPictureInsertAsOriginalState**|

## <a name="example"></a>Пример

В следующем примере заменяется каждого вхождения определенного изображения в активной публикации на другой рисунок.


```vb
Sub ReplaceLogo() 
 
Dim pgLoop As Page 
Dim shpLoop As Shape 
Dim strExistingArtName As String 
Dim strReplaceArtName As String 
 
 
strExistingArtName = "C:\path\logo 1.bmp" 
strReplaceArtName = "C:\path\logo 2.bmp" 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .Filename = strExistingArtName Then 
 .Replace (strReplaceArtName) 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
Next pgLoop 
 
End Sub
```

В этом примере проверяется каждого связанного рисунка, чтобы определить, если связанный файл был изменен с момента его был вставлен в публикацию. Если Да, изображение обновляется, заменив самого файла.




```vb
Sub UpdateModifiedLinkedPictures() 
 
Dim pgLoop As Page 
Dim shpLoop As Shape 
Dim strPictureName As String 
 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .LinkedFileStatus = pbLinkedFileModified Then 
 strPictureName = .Filename 
 .Replace (strPictureName) 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```


