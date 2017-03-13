'�ռ���
Sub ReColor()
    Dim sld As Slide
    Dim sh As Shape
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            Call ReColorSH(sh)
        Next
    Next

    ActivePresentation.ExtraColors.Add RGB(Red:=255, Green:=255, Blue:=255)
    If ActivePresentation.HasTitleMaster Then
        With ActivePresentation.TitleMaster.Background
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 0#
            .Fill.Solid
        End With
    End If
    With ActivePresentation.SlideMaster.Background
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0#
        .Fill.Solid
    End With
    With ActivePresentation.Slides.Range
        .FollowMasterBackground = msoTrue
        .DisplayMasterShapes = msoFalse
    End With

End Sub
  
Function ReColorSH(sh As Shape)
    Dim ssh As Shape
    If sh.Type = msoGroup Then ' when the shape itself is a group
        For Each ssh In sh.GroupItems
        Call ReColorSH(ssh)  ' the recursion
        Next
        '�ı乫ʽ�����ֵ���ɫΪ��ɫ����֪�������Ϊ������ɫ
        ElseIf sh.Type = msoEmbeddedOLEObject Then ' recolor the equation
   If Left(sh.OLEFormat.ProgID, 8) = "Equation" Then
                sh.PictureFormat.ColorType = msoPictureBlackAndWhite
                sh.PictureFormat.Brightness = 0
                sh.PictureFormat.Contrast = 1
                'sh.Fill.Visible = msoFalse
   End If
        '�ı��ı��������ֵ���ɫ�����Լ��趨
        ElseIf sh.HasTextFrame Then
            ' /* ��ǰ�õ�Ƭ�еĵ�ǰ��״�����ı�. */
            If sh.TextFrame.HasText Then
                ' �����ı�����е��ı�.
                Set trng = sh.TextFrame.TextRange
                ' /* �����ı�����е�ÿһ���ַ�. */
                For i = 1 To trng.Characters.Count
                    ' �����������޸�Ϊԭ������ɫֵ (��ɫ).
                    'If trng.Characters(i).Font.Color = vbWhite Then
                        ' �����������޸�ΪҪ�滻����ɫֵ (��ɫ).
                        trng.Characters(i).Font.Color = vbBlack
                    'End If
                Next
            End If
    End If
End Function