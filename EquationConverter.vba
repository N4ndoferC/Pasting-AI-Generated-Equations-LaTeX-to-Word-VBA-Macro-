' MIT License

' Copyright (c) 2026 N4ndoferC

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Sub ConvertirEcuaciones()
    Dim rInicio As Range
    Dim rFin As Range
    Dim rEcuacion As Range
    Dim textoFormula As String
    
    ' Congelar pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------
    ' FASE 1: PROCESAR DOBLES DÓLARES ($$)
    ' --------------------------------------------------
    Set rInicio = ActiveDocument.Range(0, 0)
    
    Do
        With rInicio.Find
            .ClearFormatting
            .Text = "$$"
            .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With
        
        If Not rInicio.Find.Execute Then Exit Do
        
        Set rFin = ActiveDocument.Range(rInicio.End, ActiveDocument.Content.End)
        With rFin.Find
            .ClearFormatting
            .Text = "$$"
            .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With
        
        If Not rFin.Find.Execute Then Exit Do
        
        Set rEcuacion = ActiveDocument.Range(rInicio.Start, rFin.End)
        textoFormula = ActiveDocument.Range(rInicio.End, rFin.Start).Text
        
        textoFormula = Replace(textoFormula, Chr(13), " ")
        textoFormula = Replace(textoFormula, Chr(11), " ")
        textoFormula = Replace(textoFormula, Chr(10), " ")
        textoFormula = Replace(textoFormula, Chr(9), " ")
        
        If Len(Trim(textoFormula)) > 0 Then
            rEcuacion.Text = Trim(textoFormula)
            rEcuacion.Select
            Selection.OMaths.Add Range:=Selection.Range
            
            On Error Resume Next 
            Selection.OMaths.BuildUp
            On Error GoTo 0
            
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Set rInicio = ActiveDocument.Range(Selection.End, ActiveDocument.Content.End)
        Else
            rEcuacion.Text = ""
            Set rInicio = ActiveDocument.Range(rEcuacion.End, ActiveDocument.Content.End)
        End If
    Loop
    
    ' --------------------------------------------------
    ' FASE 2: PROCESAR DÓLARES SIMPLES ($)
    ' --------------------------------------------------
    Set rInicio = ActiveDocument.Range(0, 0)
    
    Do
        With rInicio.Find
            .ClearFormatting
            .Text = "$"
            .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With
        
        If Not rInicio.Find.Execute Then Exit Do
        
        Set rFin = ActiveDocument.Range(rInicio.End, ActiveDocument.Content.End)
        With rFin.Find
            .ClearFormatting
            .Text = "$"
            .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With
        
        If Not rFin.Find.Execute Then Exit Do
        
        Set rEcuacion = ActiveDocument.Range(rInicio.Start, rFin.End)
        textoFormula = ActiveDocument.Range(rInicio.End, rFin.Start).Text
        
        textoFormula = Replace(textoFormula, Chr(13), " ")
        textoFormula = Replace(textoFormula, Chr(11), " ")
        textoFormula = Replace(textoFormula, Chr(10), " ")
        textoFormula = Replace(textoFormula, Chr(9), " ")
        
        If Len(Trim(textoFormula)) > 0 Then
            rEcuacion.Text = Trim(textoFormula)
            rEcuacion.Select
            Selection.OMaths.Add Range:=Selection.Range
            
            On Error Resume Next
            Selection.OMaths.BuildUp
            On Error GoTo 0
            
            Set rInicio = ActiveDocument.Range(Selection.End, ActiveDocument.Content.End)
        Else
            rEcuacion.Text = ""
            Set rInicio = ActiveDocument.Range(rEcuacion.End, ActiveDocument.Content.End)
        End If
    Loop
    
    Application.ScreenUpdating = True
    MsgBox "¡Ecuaciones convertidas con éxito!", vbInformation
End Sub
