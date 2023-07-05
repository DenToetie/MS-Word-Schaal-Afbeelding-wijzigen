Attribute VB_Name = "Scale_Picture"
Option Explicit

Sub Show_Scale_Form()
   ScaleForm.Show vbModeless
End Sub

Sub ScalePics_Plus5()
Dim SHP As Word.Shape
Dim ISHP As Word.InlineShape
Dim HOOGTEbegin As Long
Dim HOOGTEorg As Long
    With Word.Selection
        Select Case .Type
        Case wdSelectionInlineShape
            Set ISHP = .Range.InlineShapes(1)
            ISHP.LockAspectRatio = msoTrue
            HOOGTEbegin = ISHP.Height
            ISHP.ScaleHeight = 100
            HOOGTEorg = ISHP.Height
        Case wdSelectionShape
            Set SHP = .Range.ShapeRange(1)
            SHP.LockAspectRatio = msoTrue
            HOOGTEbegin = SHP.Height
            SHP.ScaleHeight 1, msoTrue
            HOOGTEorg = SHP.Height
        End Select
    End With
    Call Run_ScalePics(Round((HOOGTEbegin / HOOGTEorg * 100) + 5))
    '§ keer terug naar Word
    Application.Activate
End Sub

Sub ScalePics_Min5()
Dim SHP As Word.Shape
Dim ISHP As Word.InlineShape
Dim HOOGTEbegin As Long
Dim HOOGTEorg As Long
    With Word.Selection
        Select Case .Type
        Case wdSelectionInlineShape
            Set ISHP = .Range.InlineShapes(1)
            ISHP.LockAspectRatio = msoTrue
            HOOGTEbegin = ISHP.Height
            ISHP.ScaleHeight = 100
            HOOGTEorg = ISHP.Height
        Case wdSelectionShape
            Set SHP = .Range.ShapeRange(1)
            SHP.LockAspectRatio = msoTrue
            HOOGTEbegin = SHP.Height
            SHP.ScaleHeight 1, msoTrue
            HOOGTEorg = SHP.Height
        End Select
    End With
    Call Run_ScalePics(Round((HOOGTEbegin / HOOGTEorg * 100) - 5))
    '§ keer terug naar Word
    Application.Activate
End Sub

Sub ScalePics_10()
    Call Run_ScalePics(10)
End Sub

Sub ScalePics_20()
    Call Run_ScalePics(20)
End Sub

Sub ScalePics_30()
    Call Run_ScalePics(30)
End Sub

Sub ScalePics_40()
    Call Run_ScalePics(40)
End Sub

Sub ScalePics_50()
    Call Run_ScalePics(50)
End Sub

Sub ScalePics_60()
    Call Run_ScalePics(60)
End Sub

Sub ScalePics_70()
    Call Run_ScalePics(70)
End Sub

Sub ScalePics_80()
    Call Run_ScalePics(80)
End Sub

Sub ScalePics_90()
    Call Run_ScalePics(90)
End Sub

Sub ScalePics_100()
    Call Run_ScalePics(100)
End Sub

Sub Run_ScalePics(PROCENT As Single)
Dim SHP As Word.Shape
Dim ISHP As Word.InlineShape
    With Word.Selection
        Select Case .Type
        Case wdSelectionInlineShape
            Set ISHP = .Range.InlineShapes(1)
            ISHP.LockAspectRatio = msoTrue
            ISHP.ScaleHeight = PROCENT
        Case wdSelectionShape
            Set SHP = .Range.ShapeRange(1)
            SHP.LockAspectRatio = msoTrue
            SHP.ScaleHeight PROCENT / 100, msoTrue
        End Select
    End With
    '§ keer terug naar Word
    Application.Activate
End Sub

