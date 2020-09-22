VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Region Test"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "The lazy mans guide to Windows regions  - please vote if this helps"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      Height          =   2055
      Left            =   120
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Move your mouse over parts of the house, you can also click a region to do a hit test."
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AllAreas As New Areas

Dim bOver As Boolean
Private Sub Form_Load()
Set AllAreas = New Areas
DoTest

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo EH

Dim I As Integer
Dim LRGN As Long

LRGN = IsInRegion(AllAreas, x, y)
For I = 1 To AllAreas.Count
If AllAreas(I).AreaNumber = LRGN Then
        
        If AllAreas(I).AreaSelected = False Then
            AllAreas(I).AreaSelected = True
            InvertRgn AllAreas.ParentHDC, AllAreas(I).AreaNumber
            List1.AddItem "Mouse_Over: " & AllAreas(I).AreaName, 0
        End If
Else
        If AllAreas(I).AreaSelected = True Then
            AllAreas(I).AreaSelected = False
            InvertRgn AllAreas.ParentHDC, AllAreas(I).AreaNumber
        End If
End If
Next I

Exit Sub
EH:
 'do nothing
Exit Sub
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim I As Integer
Dim LRGN As Long
LRGN = IsInRegion(AllAreas, x, y)
If LRGN <> 0 Then
    For I = 1 To AllAreas.Count
    If AllAreas(I).AreaNumber = LRGN Then
        List1.AddItem "Mouse_Click: " & AllAreas(I).AreaName, 0
    End If
    Next I
End If
End Sub

Private Sub Form_Paint()
'painting with the API is very fast
Me.Cls
Dim I As Integer
For I = 1 To AllAreas.Count
If AllAreas(I).AreaState = 0 Then
    PaintARGN AllAreas.ParentHDC, AllAreas(I).AreaNumber, AllAreas(I).AreaPen, AllAreas(I).AreaBrush
Else
    PaintARGN AllAreas.ParentHDC, AllAreas(I).AreaNumber, AllAreas(I).AreaPen, AllAreas(I).AreaAlertBrush
End If
        
Next I

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Be sure to call clear all..
'otherwise the regions stay in memory after your app closes
'and that means bad things can happen in memory
AllAreas.ClearAll
End Sub

Private Sub DoTest()
'below is for the test
''''''''''''''''''''''''''''''''''
Dim RC As RECT
Dim bOK As Boolean

AllAreas.ParentHDC = Me.Hdc

With RC
.Left = 14
.Top = 96
.Right = 105
.Bottom = 230
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Kitchen", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

With RC
.Left = 113
.Top = 30
.Right = 253
.Bottom = 229
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Living Room", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

With RC
.Left = 260
.Top = 30
.Right = 318
.Bottom = 108
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Guest Bedroom", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

With RC
.Left = 325
.Top = 30
.Right = 384
.Bottom = 108
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Home Office", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

With RC
.Left = 391
.Top = 30
.Right = 470
.Bottom = 138
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Master Bedroom", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)


With RC
.Left = 391
.Top = 144
.Right = 470
.Bottom = 164
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Small Bathroom", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

With RC
.Left = 352
.Top = 170
.Right = 470
.Bottom = 230
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Ready Room", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

With RC
.Left = 352
.Top = 236
.Right = 471
.Bottom = 379
End With
'Adding a rectangle is easy....
bOK = AddRGNRectangle(AllAreas, RC, "Garage", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)

'add a circle just to prove we can

With RC
.Left = 10
.Top = 10
.Right = 90
.Bottom = 90
End With
'Adding a rectangle is easy....
bOK = AddRGNElliptic(AllAreas, RC, "Swimming Pool", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)


'add a polygon
Dim Z(1 To 6) As POINTAPI
nCount = 7
Z(1).x = 260
Z(1).y = 146
Z(2).x = 258
Z(2).y = 229
Z(3).x = 345
Z(3).y = 229
Z(4).x = 346
Z(4).y = 167
Z(5).x = 324
Z(5).y = 146
Z(6).x = 260
Z(6).y = 146
'send an array of X,Y positions to the function, and the count
bOK = AddRGNPoly(AllAreas, Z, UBound(Z), "Big Bathroom", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)


'Add a polygon
Dim P(1 To 7) As POINTAPI
nCount = 7
P(1).x = 260
P(1).y = 116
P(2).x = 259
P(2).y = 139
P(3).x = 324
P(3).y = 140
P(4).x = 351
P(4).y = 164
P(5).x = 384
P(5).y = 164
P(6).x = 384
P(6).y = 115
P(7).x = 260
P(7).y = 116
'send an array of X,Y positions to the function, and the count
bOK = AddRGNPoly(AllAreas, P, UBound(P), "Hallway", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED)
End Sub

