VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim oData As Matrix
Dim vSampleNames As Variant
Dim vVariableNames As Variant
Dim oAverages As Vector
Dim oStdevs As Vector

Property Let Data(inData As Variant)
    Set oData = New Matrix
    oData.Mat = inData
End Property

Property Let SampleNames(inData As Variant)
    vSampleNames = inData
    ReDim Preserve vSampleNames(1 To Me.NumberOfSamples, 1 To 1)
End Property

Property Let VariableNames(inData As Variant)
    vVariableNames = inData
    vVariableNames = WorksheetFunction.Transpose(vVariableNames)
    ReDim Preserve vVariableNames(1 To Me.NumberOfVariables, 1 To 1)
End Property

Property Get Data() As Matrix
    Set Data = oData
End Property

Property Get NumberOfSamples()
    NumberOfSamples = oData.M
End Property

Property Get NumberOfVariables()
    NumberOfVariables = oData.N
End Property

Property Get Averages() As Vector
    Set Averages = oAverages
End Property

Property Get Stdevs() As Vector
    Set Stdevs = oStdevs
End Property

Property Get SampleNames()
    SampleNames = vSampleNames
End Property

Property Get VariableNames()
    VariableNames = vVariableNames
End Property

Public Sub CalcAves()
    iVariables = oData.N
    ReDim vAverages(1 To iVariables, 1 To 1)
    
    Dim i As Long
    
    For i = 1 To iVariables
        vAverages(i, 1) = oData.GetColumn(i).Average
    Next i
    Set oAverages = New Vector
    oAverages.Vec = vAverages
End Sub

Public Sub CalcStdevs()
    iVariables = oData.N
    ReDim vStdevs(1 To iVariables, 1 To 1)
    
    Dim i As Long
    
    For i = 1 To iVariables
        vStdevs(i, 1) = oData.GetColumn(i).Stdev
    Next i
    Set oStdevs = New Vector
    oStdevs.Vec = vStdevs
End Sub
