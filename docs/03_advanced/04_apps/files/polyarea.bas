Attribute VB_Name = "Module2"
Option Explicit

' gets the area of the selected polyline and prints it in the polyline
 
' select the polyline
' add to selection set (by running macro)
' get area of polyline
' find middle of bounding box
' print the area in the middle of the bounding box
 
Option Explicit
Dim pline As AcadEntity
Dim plinearea As Double
Dim middlePoint(0 To 2) As Double
 
Public Sub subGetAreaPline()
 
Dim pfSS As AcadSelectionSet
Dim polyline1 As AcadEntity
Dim n As Integer
 

Set pfSS = ThisDrawing.SelectionSets.Add("junk2")
pfSS.Clear
pfSS.SelectOnScreen
 
For n = 0 To pfSS.Count - 1
 
Set polyline1 = pfSS.Item(n)
 

getArea polyline1
middle polyline1
print_area
 
Next n
 
pfSS.Clear
pfSS.Delete
 
End Sub
 
Private Sub getArea(pline As AcadEntity)
 
plinearea = Round(pline.area, 4)
 
End Sub
 
Private Sub middle(pline As AcadEntity)
 
Dim minExt As Variant
Dim maxExt As Variant
 
'return the bounding box for the line and return the minimum
'and maximum extents of the box in the minExt and maxExt variables.
pline.GetBoundingBox minExt, maxExt
 
'caluculates the middle of the bounding box
 
middlePoint(0) = ((maxExt(0) - minExt(0)) / 2) + minExt(0)
middlePoint(1) = ((maxExt(1) - minExt(1)) / 2) + minExt(1)
middlePoint(2) = ((maxExt(2) - minExt(2)) / 2) + minExt(2)
 
End Sub
 
Private Sub print_area()
 
Dim textObj As AcadText
Dim areaString As String
Dim height As Double
 
areaString = plinearea
height = 1#
 
Set textObj = ThisDrawing.ModelSpace.AddText(areaString, middlePoint, height)
 
End Sub
 

