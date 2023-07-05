VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScaleForm 
   Caption         =   "Scale Picture"
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   OleObjectBlob   =   "ScaleForm 03.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScaleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComZoom10_Click()
    Call Run_ScalePics(10)
End Sub

Private Sub ComZoom20_Click()
    Call Run_ScalePics(20)
End Sub

Private Sub ComZoom30_Click()
    Call Run_ScalePics(30)
End Sub

Private Sub ComZoom40_Click()
    Call Run_ScalePics(40)
End Sub

Private Sub ComZoom50_Click()
    Call Run_ScalePics(50)
End Sub

Private Sub ComZoom60_Click()
    Call Run_ScalePics(60)
End Sub

Private Sub ComZoom70_Click()
    Call Run_ScalePics(70)
End Sub

Private Sub ComZoom80_Click()
    Call Run_ScalePics(80)
End Sub

Private Sub ComZoom90_Click()
    Call Run_ScalePics(90)
End Sub

Private Sub ComZoom100_Click()
    Call Run_ScalePics(100)
End Sub

Private Sub ComZoomMin5_Click()
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
End Sub

Private Sub ComZoomPlus5_Click()
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
End Sub


'§ ====== AAN ZETTEN ALS de "Scale_Picture.bas" niet aanwezig is =====

'Sub Run_ScalePics(PROCENT As Single)
'Dim SHP As Word.Shape
'Dim ISHP As Word.InlineShape
'    With Word.Selection
'        Select Case .Type
'        Case wdSelectionInlineShape
'            Set ISHP = .Range.InlineShapes(1)
'            ISHP.LockAspectRatio = msoTrue
'            ISHP.ScaleHeight = PROCENT
'        Case wdSelectionShape
'            Set SHP = .Range.ShapeRange(1)
'            SHP.LockAspectRatio = msoTrue
'            SHP.ScaleHeight PROCENT / 100, msoTrue
'        End Select
'    End With
'End Sub

'§ ====== AAN ZETTEN ALS de "Scale_Picture.bas" niet aanwezig is =====



