Attribute VB_Name = "CommentImage"
Sub CommentImage()
	Dim imagen
	imagen = Application.GetOpenFilename(" , *.jpg*;*.png*;*.bmp*")

	If (imagen <> False) Then
		With ActiveCell
			If Not (ActiveCell.Comment Is Nothing) Then
				ActiveCell.Comment.Delete
			End If

			.AddComment
			.Comment.Shape.Width = 854 'Width		16:9 Aspect resolution
			.Comment.Shape.Height = 480 'Height		2.25 times smaller than 1080p

			ActiveCell.Comment.Shape.Fill.UserPicture imagen
		End With
	End If
End Sub
