Option Explicit

Sub PanelizePCB()

    Dim pcbLength As Double
    Dim pcbWidth As Double
    Dim spacing As Double
    Dim panelLength As Double
    Dim panelWidth As Double
    Dim maxPCBs As Integer
    Dim maxRotated As Integer
    Dim ws As Worksheet

    ' Use active sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Step 1: Ask user for dimensions
    ws.Range("A1").Value = "PCB Length (mm):"
    ws.Range("A2").Value = "PCB Width (mm):"
    ws.Range("A3").Value = "Panel Length (mm):"
    ws.Range("A4").Value = "Panel Width (mm):"
    ws.Range("A5").Value = "Spacing (mm):"
    ws.Range("A5").Value = 2 ' Default value

    pcbLength = ws.Range("B1").Value
    pcbWidth = ws.Range("B2").Value
    panelLength = ws.Range("B3").Value
    panelWidth = ws.Range("B4").Value
    spacing = ws.Range("B5").Value

    ' Ensure width is the smaller dimension
    If pcbWidth > pcbLength Then
        Dim temp As Double
        temp = pcbLength
        pcbLength = pcbWidth
        pcbWidth = temp
        ws.Range("B1").Value = pcbLength
        ws.Range("B2").Value = pcbWidth
    End If

    ' Step 2: Compute max PCBs without rotation
    maxPCBs = Int((panelWidth + spacing) / (pcbWidth + spacing)) * Int((panelLength + spacing) / (pcbLength + spacing))

    ' Step 3: Compute max PCBs with rotation
    maxRotated = Int((panelWidth + spacing) / (pcbLength + spacing)) * Int((panelLength + spacing) / (pcbWidth + spacing))
    If maxRotated > maxPCBs Then
        maxPCBs = maxRotated
    End If
    Range("D5").Value = maxPCBs & " Max pcb without rotation"
    

    ' Step 4: Mixed calculation
    Dim maxMixed As Integer
    Dim maxLines As Integer
    Dim currentTotal As Integer
    Dim i As Integer
    Dim pcbPerLineNoRotate As Integer
    Dim pcbPerLineRotate As Integer
    Dim remainingSpace As Double
    remainingSpace = panelLength - pcbLength - spacing
    i = 1
    
    maxMixed = maxPCBs
    pcbPerLineNoRotate = Int((panelWidth + spacing) / (pcbWidth + spacing))
    pcbPerLineRotate = Int((panelWidth + spacing) / (pcbLength + spacing))
    Range("D10").Value = pcbPerLineNoRotate & " number of pcb per line without rotation"
    Range("D11").Value = pcbPerLineRotate & " number of pcb per line with rotation"
        
    Do
        Range("D13").Value = "Remaining space after " & i & " loop iteration: " & remainingSpace
        currentTotal = (i * pcbPerLineNoRotate) + pcbPerLineRotate * Int((remainingSpace + spacing) / (pcbWidth + spacing))
        remainingSpace = remainingSpace - pcbLength - spacing
        Range("E13").Value = "Remaining space: " & remainingSpace
        
        Range("D2").Value = i & " lines without rotation"
        Range("D14").Value = "Total on this loop " & i & " " & currentTotal
        If currentTotal > maxMixed Then
            Range("D4").Value = currentTotal & " current total"
            maxMixed = currentTotal
        End If
        i = i + 1
    Loop While remainingSpace + spacing >= pcbWidth + spacing

    maxPCBs = maxMixed

    ws.Range("A7").Value = "Maximum PCBs you can place:"
    ws.Range("B7").Value = maxPCBs

End Sub

