---
title: "Метод PictureFormat.ReplaceEx (издатель)"
keywords: vbapb10.chm3604808
f1_keywords: vbapb10.chm3604808
ms.prod: publisher
api_name: Publisher.PictureFormat.ReplaceEx
ms.assetid: 0f1b9eaf-51b6-ae21-518f-55663184ab87
ms.date: 06/08/2017
ms.openlocfilehash: 081703ec3927e5d8564ae4b8bdd27b0ec542bd98
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatreplaceex-method-publisher"></a>Метод PictureFormat.ReplaceEx (издатель)

Заменяет указанный рисунок, при необходимости Подгонка замещающего рисунка в элементе frame или заполнение фрейма. Возвращает значение nothing.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReplaceEx** ( _Pathname_, _InsertAs_ _под размер_)

 переменная _expression_A, представляет собой объект- [PictureFormat](pictureformat-object-publisher.md) .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя пути|Обязательное свойство.| **String**|Имя и путь к файлу, с которым вы хотите заменить указанный рисунок.|
|InsertAs|Необязательный| **[PbPictureInsertAs](pbpictureinsertas-enumeration-publisher.md)**|Так, в которой будет вставлено в документ файл рисунка: связанные или внедренные.|
|Подобрать|Необязательный| **[pbPictureInsertFit](pbpictureinsertfit-enumeration-publisher.md)**|Является ли вставленное изображение Вписать в элементе frame или Заливка рамки.|

## <a name="remarks"></a>Заметки

Параметр _InsertAs_ может иметь одно из следующих **PbPictureInsertAs** константы, описанные в библиотеке типов, Microsoft Publisher. Значение по умолчанию — **pbPictureInsertAsOriginalState**.



| **pbPictureInsertAsEmbedded**|| **pbPictureInsertAsLinked**|| **pbPictureInsertAsOriginalState**|

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **ReplaceEx** замена всех изображений в публикации на другой рисунок. В этом примере вписать замещающего рисунка в кадрах предыдущей изображения, но **pbFill** вместо **pbFit** можно использовать, если вместо этого заполняемых кадров. В этом примере также исключает изображения на главных страницах.

Прежде чем запустить этот макрос, замените replacementPicturePath путь к изображению, которые вы хотите использовать в качестве замены.




```vb
Public Sub ReplaceEx_Example()
    
    Dim pubPage As Page
    Dim pubShape As Shape
    Dim strReplacePicturePath As String
    
    strReplacePicturePath = replacementPicturePath
    
    For Each pubPage In ActiveDocument.Pages
        
        For Each pubShape In pubPage.Shapes
            
            If pubShape.Type = pbPicture Then

                pubShape.PictureFormat.ReplaceEx strReplacePicturePath, pbPictureInsertAsOriginalState, pbFit

            End If
        
        Next pubShape
        
    Next pubPage
            
End Sub
```


