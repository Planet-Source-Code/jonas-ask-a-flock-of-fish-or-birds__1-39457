Attribute VB_Name = "publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Type Person
 Act As Boolean
 X As Currency
 Y As Currency
 tX As Integer
 tY As Integer
 Atr As Currency
 AtTarget As Boolean
End Type


Public Men(1 To 30) As Person

Public MaxSpeed As Boolean



Public Sub PaintBoard()
    FrmMain.Cls
    For A = 1 To UBound(Men)
        If A = 1 Then FrmMain.Line (Men(A).X, Men(A).Y)-(Men(A).tX, Men(A).tY)
        FrmMain.Line (Men(A).X, Men(A).Y)-Step(5, 5), IIf(A = 1, vbBlue, 0), BF
    Next A
    
End Sub

Public Function FindAng(X, Y, Sx, Sy, RAD) 'convert as set of coordinates to the angle between them (given standard VB scales)
Const Pi = 3.14159265358979
Dim Ang As Currency
    If Y - Sy = 0 Then
        Ang = IIf(X >= Sx, 0, Pi)
    Else
        Ang = Atn((X - Sx) / (Y - Sy))
        Ang = Ang + (Pi / 2)
        If Y >= Sy Then
            Ang = Pi + Ang
        End If
    End If
    If RAD = 0 Then Ang = Ang / Pi * 180
    FindAng = Ang
End Function
Public Sub DoStuffMen()
Dim A As Integer
    For A = 1 To UBound(Men)
    With Men(A)
        .AtTarget = False
        If .X + 1 > .tX And .X - 1 < .tX Then
        If .Y + 1 > .tY And .Y - 1 < .tY Then
            .X = .tX
            .Y = .tY
            .AtTarget = True
        End If
        End If
        If A <> 1 Then
            .tX = Men(1).X
            .tY = Men(1).Y
        End If
        If .AtTarget = True Then
            If Rnd > 0.97 Then
                .tX = Rnd * FrmMain.ScaleWidth
                .tY = Rnd * FrmMain.ScaleHeight
            End If
        End If
        MoveMan A
    End With
    Next A
End Sub
Public Sub MoveMan(A As Integer)
Dim TempX As Currency, TempY As Currency
    K = 2

    With Men(A)
        TempX = 0
        TempY = 0
        If .AtTarget = False Then
            Ang = FindAng(.X, .Y, .tX, .tY, 1)
            TempX = -Cos(Ang) * 2
            TempY = Sin(Ang) * 2
        End If
        
        'Get influence vectors
        For b = 1 To UBound(Men)
            Magn = 0
            If b <> A Then
                r = Sqr((Men(A).X - Men(b).X) ^ 2 + (Men(A).Y - Men(b).Y) ^ 2)
                If r < 50 Then
                    Ang = FindAng(Men(A).X, Men(A).Y, Men(b).X, Men(b).Y, 1)
                    If r > 0 Then
                        Magn = (K * (Men(A).Atr * Men(b).Atr) / r ^ 2)
                    End If
                                    
                    TempX = TempX + (Cos(Ang) * Magn)
                    TempY = TempY - (Sin(Ang) * Magn)
                End If
            End If
        Next b
        
        
        If TempX < -2 Then TempX = -2
        If TempX > 2 Then TempX = 2
        If TempY < -2 Then TempY = -2
        If TempY > 2 Then TempY = 2
        .X = .X + TempX
        .Y = .Y + TempY
    End With
End Sub

