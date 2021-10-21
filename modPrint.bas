Attribute VB_Name = "modPrint"

Option Explicit
Public prnCopies As Integer

'//Print ONLY front cover
Public Sub PrintFront()
'Variables
Dim TopMargin As Integer
Dim LeftMargin As Integer
Dim Width As Integer
Dim Height As Integer
        
    'assigns values (in twips) to variables
    TopMargin = 1440
    LeftMargin = 1440
    Width = 6900
    Height = 6900
          
    Printer.Orientation = 1 'Prints in protrait
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
   
    Printer.PaintPicture frmMain!imgFront.Picture, _
        LeftMargin, TopMargin, Width, Height
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub

'//Print ONLY back cover
Public Sub PrintBack()
'Variables
Dim bTopMargin As Integer
Dim bLeftMargin As Integer
Dim bWidth As Integer
Dim bHeight As Integer
    
    'assigns values (in twips) to variables
    bTopMargin = 1440
    bLeftMargin = 1440
    bWidth = 8530
    bHeight = 6700
    
    Printer.Orientation = 1 'Prints in Portrait
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgBack.Picture, bLeftMargin, _
        bTopMargin, bWidth, bHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print ONLY front and inside (separate) covers
Public Sub PrintFrontAndInside()
'Variables
Dim FrontTopMargin As Integer
Dim FrontLeftMargin As Integer
Dim FrontWidth As Integer
Dim FrontHeight As Integer
Dim InsideTopMargin As Integer
Dim InsideLeftMargin As Integer
Dim InsideWidth As Integer
Dim InsideHeight As Integer
    
    'assigns values (in twips) to variables
    FrontTopMargin = 1000
    FrontLeftMargin = 6900 + 1010
    FrontWidth = 6900
    FrontHeight = 6900
    
    'assigns values (in twips) to variables
    InsideTopMargin = 1000
    InsideLeftMargin = 1000
    InsideWidth = 6900
    InsideHeight = 6900
    
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFrontS.Picture, FrontLeftMargin, _
    FrontTopMargin, FrontWidth, FrontHeight
    
    Printer.PaintPicture frmMain!imgInsideS.Picture, InsideLeftMargin, _
        InsideTopMargin, InsideWidth, InsideHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print Front and Inside (Whole) cover
Public Sub PrintWhole()
'Variables
Dim WholeTopMargin As Integer
Dim WholeLeftMargin As Integer
Dim WholeWidth As Integer
Dim WholeHeight As Integer
    
    'assigns values (in twips) to variables
    WholeTopMargin = 1440
    WholeLeftMargin = 1440
    WholeWidth = 13800
    WholeHeight = 6900
    
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFront_Inside_W.Picture, WholeLeftMargin, _
    WholeTopMargin, WholeWidth, WholeHeight

    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print ONLY Front and Back covers
Public Sub PrintFrontAndBack()
'Variables
Dim FrontTopMargin As Integer
Dim FrontLeftMargin As Integer
Dim FrontWidth As Integer
Dim FrontHeight As Integer
Dim BackTopMargin As Integer
Dim BackLeftMargin As Integer
Dim BackWidth As Integer
Dim BackHeight As Integer

    'assigns values (in twips) to variables
    FrontTopMargin = 1000
    FrontLeftMargin = 1440 + 720
    FrontWidth = 6900
    FrontHeight = 6900
    
    'assigns values (in twips) to variables
    BackTopMargin = 6900 + 1440
    BackLeftMargin = 1440
    BackWidth = 8530
    BackHeight = 6700
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFront.Picture, FrontLeftMargin, _
    FrontTopMargin, FrontWidth, FrontHeight
    
    Printer.PaintPicture frmMain!imgBack.Picture, BackLeftMargin, _
        BackTopMargin, BackWidth, BackHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print Front and Inside (Separate) and Back covers
Public Sub PrintFrontsAndInsideSAndBack()
'Fron and Inside Variables
Dim FrontTopMargin As Integer
Dim FrontLeftMargin As Integer
Dim FrontWidth As Integer
Dim FrontHeight As Integer
Dim InsideTopMargin As Integer
Dim InsideLeftMargin As Integer
Dim InsideWidth As Integer
Dim InsideHeight As Integer
'Back Variables
Dim bTopMargin As Integer
Dim bLeftMargin As Integer
Dim bWidth As Integer
Dim bHeight As Integer
 
    
    'assigns values (in twips) to variables
    FrontTopMargin = 1000
    FrontLeftMargin = 6900 + 1010
    FrontWidth = 6900
    FrontHeight = 6900
    
    'assigns values (in twips) to variables
    InsideTopMargin = 1000
    InsideLeftMargin = 1000
    InsideWidth = 6900
    InsideHeight = 6900
    
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFrontS.Picture, FrontLeftMargin, _
    FrontTopMargin, FrontWidth, FrontHeight
    
    Printer.PaintPicture frmMain!imgInsideS.Picture, InsideLeftMargin, _
        InsideTopMargin, InsideWidth, InsideHeight
    
    Printer.NewPage 'Begins new page for back cover
        
    'assigns values (in twips) to variables
    bTopMargin = 1440
    bLeftMargin = 1440
    bWidth = 8530
    bHeight = 6700
    
    Printer.Orientation = 1 'Prints in Portrait
    
    Printer.PaintPicture frmMain!imgBack.Picture, bLeftMargin, _
        bTopMargin, bWidth, bHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
    
End Sub


'//Print Front and Inside (whole) and Back covers
Public Sub PrintWholeAndBack()
'Whole Variables
Dim WholeTopMargin As Integer
Dim WholeLeftMargin As Integer
Dim WholeWidth As Integer
Dim WholeHeight As Integer
    
'Back Variables
Dim bTopMargin As Integer
Dim bLeftMargin As Integer
Dim bWidth As Integer
Dim bHeight As Integer
    
    
    'assigns values (in twips) to variables
    WholeTopMargin = 1440
    WholeLeftMargin = 1440
    WholeWidth = 13800
    WholeHeight = 6900
        
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFront_Inside_W.Picture, WholeLeftMargin, _
    WholeTopMargin, WholeWidth, WholeHeight
    
    Printer.NewPage 'Begins new page for back cover
    
    'assigns values (in twips) to variables
    bTopMargin = 1440
    bLeftMargin = 1440
    bWidth = 8530
    bHeight = 6700
    
    Printer.Orientation = 1 'Prints in Portrait
    
    Printer.PaintPicture frmMain!imgBack.Picture, bLeftMargin, _
        bTopMargin, bWidth, bHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
    
End Sub

