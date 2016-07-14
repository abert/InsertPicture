'by Andy Bertagnoli
'xlxbert@gmail.com
'www.xlxbert.com
'a function to insert images into cells with a formula call.

Public Function ipx(PictureFileName As String, TargetCell As Range)

    Dim tc As String

    tc = TargetCell.Address

    Dim t As Double, l As Double, w As Double, h As Double, PictureFileNamec As String, hgth As Integer, wdth As Integer
     'THIS IS IMPORTANT!!!!
     'CHANGE "C:\IMG\" to the location of your images
     PictureFileNamec = "C:\img\" & PictureFileName & ".jpg"
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    If Dir(PictureFileNamec) = "" Then Exit Function
    
    Set p = ActiveSheet.Pictures.Insert(PictureFileNamec)
    
    With p
       hgth = .Height
       wdth = .Width
       .Name = tc
    End With
    
    Set p = Nothing
    
    For Each Shape In ActiveSheet.Shapes
        If Shape.Name = tc Then
            Shape.Delete
        End If
    Next
     
   Set pix = ActiveSheet.Shapes.AddPicture(PictureFileNamec, False, True, 100, 100, wdth, hgth)
    pix.Placement = xlMoveAndSize
    pix.ControlFormat.PrintObject = True

    With TargetCell
            t = .Top
            l = .Left
            
    End With
    

    With pix
        .Top = t
        .Left = l
        .LockAspectRatio = True
        .Height = TargetCell.RowHeight
        .Name = tc
    End With
    
    If pix.Width > TargetCell.Width Then
        With pix
            .Top = t
            .Left = l
            .Width = TargetCell.Width
        End With
    End If
   
    Set pix = Nothing
    
    
 End Function
