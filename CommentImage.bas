Attribute VB_Name = "CommentImage"
Sub CommentImage()
Attribute CommentImage.VB_ProcData.VB_Invoke_Func = "s\n14"
	Dim imagen
	imagen = Application.GetOpenFilename("File (*.jpg*;*.png*;*.bmp*), *.jpg*;*.png*;*.bmp*")

	'Dialog Control
	If imagen = False Then 
		Exit Sub
	End If
	
	With ActiveCell
		If Not (ActiveCell.Comment Is Nothing) Then
			ActiveCell.Comment.Delete
		End If
		
		.AddComment
		.Comment.Shape.Width = 854 'Width		16:9 Aspect resolution
		.Comment.Shape.Height = 480 'Height		2.25 times smaller than 1080p
		
		ActiveCell.Comment.Shape.Fill.UserPicture imagen
	End With
End Sub
