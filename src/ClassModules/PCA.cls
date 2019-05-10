Dim oDataset As Dataset
Dim oNewDataset As Dataset
Dim oNormalData As Matrix
Dim oNormalNewData As Matrix
Dim oCorrelation As Matrix
Dim oEigen As Eigen
Dim oScores As Matrix
Dim bScale As Boolean
Dim bCenter As Boolean

Property Set Data(inData As Dataset)
    Set oDataset = inData
    oDataset.CalcAves
    oDataset.CalcStdevs
End Property

Property Set NewData(inData As Dataset)
    Set oNewDataset = inData
    If bCenter Or bScale Then Normalize (True)
End Property

Property Let PCACenter(inSet As Boolean)
    bCenter = inSet
End Property

Property Let PCAScale(inSet As Boolean)
    bScale = inSet
End Property

Property Get NormalData()
    NormalData = oNormalData.Mat
End Property

Property Get Correlation()
    Correlation = oCorrelation.Mat
End Property

Property Get EigenVec()
    EigenVec = oEigen.EigenVec.Mat
End Property

Property Get EigenVal()
    EigenVal = oEigen.EigenVal.Vec
End Property

Property Get Scores()
    Scores = oScores.Mat
End Property

Property Get Loading() As Matrix
    Set Loading = oEigen.EigenVec
End Property

' Function: Normalize
' Normalise the dataset in the ways sepcified in the object
' The NewData parameter chooses which dataset to normalize.
Public Sub Normalize(Optional NewData = False)
    Dim NewDataSet As Dataset
    If NewData Then Set NewDataSet = oNewDataset Else Set NewDataSet = oDataset
    iSamples = NewDataSet.NumberOfSamples
    
    Dim oOnesColumn As Matrix
    Set oOnesColumn = ScalarMatrix(1, iSamples, 1)
    Dim oResult, oStdevs As Matrix
    Set oResult = NewDataSet.Data
    If bCenter Then
        Set oResult = oResult.Subtract(oOnesColumn.Multiply(oDataset.Averages.toRowMatrix()))
    End If
    If bScale Then
        Set oStdevs = DiagonalMatrix(oDataset.Stdevs)
        Set oResult = oResult.Multiply(oStdevs.Inverse())
    End If
    
    If NewData Then
        Set oNormalNewData = oResult
    Else
        Set oNormalData = oResult
    End If
End Sub

' Function: CorrelationMatrix
' Create either the Covariance or correlation matrix from a dataset
Public Sub CorrelationMatrix()
    iSamples = oDataset.NumberOfSamples
    Dim oCorrelationt As Matrix
    Set oCorrelationt = oNormalData.Transpose()
    Set oCorrelation = oCorrelationt.Multiply(oNormalData).ScalarDivide(iSamples - 1)
End Sub

Public Sub Run()
    If bCenter Or bScale Then Normalize
    CorrelationMatrix
    Set oEigen = New Eigen
    Set oEigen.Matrix = oCorrelation
    oEigen.JK
    Transform
End Sub

Public Sub Transform(Optional iNum As Integer)
    Set oScores = oNormalData.Multiply(oEigen.EigenVec)
End Sub

Public Sub PRESS()
    Dim i As Integer
    Dim iSample As Integer
    iSample = vData.M
    Dim oRow As Matrix
    Dim oErrorMatrix As Matrix
    For i = 1 To iVariables
        For j = 1 To iSample
            Dim InnerPCA As New PCA
            InnerPCA.Data = vData.ExcludeRow(j)
            InnerPCA.Run
            Set oRow = InnerPCA.TransformNewData(vData.GetRow(j).Vec, i)
            Set oPress = oPress.AugmentBelow(oRow)
        Next j
        Set oErrorMatrix = oNormalData.Subtract(oPress)
        vPress(i) = Math.sqr(WorksheetFunction.SumSq(oErrorMatrix.Mat))
    Next i
End Sub

Public Function getRESS() As Vector
    Dim vRsq As Variant
    vRsq = getCumRsq().Vec
    Inertia = oEigen.EigenVal.Sum * (oDataset.NumberOfSamples - 1)
    For i = 1 To oDataset.NumberOfVariables
        vRsq(i, 1) = (1 - vRsq(i, 1)) * Inertia
    Next i
    Dim oResult As New Vector
    oResult.Vec = vRsq
    Set getRESS = oResult
End Function

Public Function getRsq() As Vector
    Set getRsq = oEigen.EigenVal.ScalarDivide(oEigen.EigenVal.Sum)
End Function

Public Function getCumRsq() As Vector
    Dim vRsq As Variant
    vRsq = getRsq().Vec
    For i = 2 To oDataset.NumberOfVariables
        vRsq(i, 1) = vRsq(i, 1) + vRsq(i - 1, 1)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vRsq
    Set getCumRsq = oResult
End Function

Public Function getDmodX()
    ' if inDataset is Nothing
    '   inDataset = oDataset
    ' else
    '   Check that inDataset.numberOfVariables = oDataset.NumberOfVariables
    '   ' The degrees of freedom in the scaling factor need to be changed in this case
    '   Normalize incoming data
    ' end if
    Dim oDmodX As Matrix
    Set oDmodX = ScalarMatrix("NA", oDataset.NumberOfSamples, oDataset.NumberOfVariables)
    ReDim vDmodX(1 To oDataset.NumberOfSamples, 1 To oDataset.NumberOfVariables)
    vRESS = getRESS()
    Ao = 1   ' if data not centered Ao = 0
    Dim ReducedScores, ReducedLoading, Residuals As Matrix
    Dim SSE As Vector
    ' If supplying new data, can go out to all variables
    For i = 1 To oDataset.NumberOfVariables - 1
      ' Select i loadings and i scores
        Set ReducedScores = oScores.Submatrix(1, oDataset.NumberOfSamples, 1, i)
        Set ReducedLoading = oEigen.EigenVec.Submatrix(1, oDataset.NumberOfVariables, 1, i)
        Set Residuals = oNormalData.Subtract(ReducedScores.Multiply(ReducedLoading.Transpose()))
        Set SSE = Residuals.Multiply(Residuals.Transpose()).GetDiagonalElements()
        Scaling_Factor = vRESS(i, 1) / (oDataset.NumberOfSamples - i - Ao)
        Set oDmodX = oDmodX.ReplaceColumn(i, SSE.ScalarDivide(Scaling_Factor).sqr())
    Next i
    Set getDmodX = oDmodX
End Function

Public Function getQ(Optional bUseNewdata = False) As Matrix
    Dim oIdentity, oLoadings, oStdevs As Matrix
    Dim oScaledX As New Matrix
    Dim oAverages As New Matrix
    If bUseNewdata Then Set oScaledX = oNormalNewData Else Set oScaledX = oNormalData

    Dim vResult As Variant
    Dim oResult As Matrix
    Set oResult = oScaledX.ScalarMultiply(0)
    ReDim vResult(1 To 1, 1 To oDataset.NumberOfVariables)
    Set oIdentity = Identity(oDataset.NumberOfVariables)
    For i = 1 To oDataset.NumberOfVariables
        Set oLoadings = oEigen.EigenVec.Submatrix(1, oDataset.NumberOfVariables, 1, i)
        Set oResult = oResult.ReplaceColumn(i, oScaledX.Multiply(oIdentity.Subtract(oLoadings.Multiply(oLoadings.Transpose()))).Multiply(oScaledX.Transpose()).GetDiagonalElements())
    Next i
    Set getQ = oResult
End Function

' Function: getT2
' Get the T2 distances of the model or new data
'
' Parameters:
'   bUseNewData - True if the new data dataset object is to be used
Public Function getT2(Optional bUseNewdata = False) As Matrix
    nSamp = oDataset.NumberOfSamples
    nVar = oDataset.NumberOfVariables
    Dim oResult, oInverseEvals As Matrix
    Dim oReducedInverseEvals As Matrix
    Dim oReducedScores As Matrix
    Dim ithT2 As Vector
    Set oResult = ScalarMatrix(0, nSamp, nVar)
    Set oInverseEvals = DiagonalMatrix(oEigen.EigenVal).Inverse()
    For i = 1 To nVar
        Set oReducedInverseEvals = oInverseEvals.Submatrix(1, i, 1, i)
        Set oReducedScores = oScores.Submatrix(1, nSamp, 1, i)
        Set ithT2 = oReducedScores.Multiply(oReducedInverseEvals).Multiply(oReducedScores.Transpose()).GetDiagonalElements()
        Set oResult = oResult.ReplaceColumn(i, ithT2)
    Next i
    Set getT2 = oResult
End Function

Public Function getCritT2(Optional alpha = 0.05) As Vector
    nSamp = oDataset.NumberOfSamples
    nVar = oDataset.NumberOfVariables
    Dim vResult As Variant
    ReDim vResult(1 To nVar, 1 To 1)
    For i = 1 To nVar
        vResult(i, 1) = i * (nSamp - 1) * WorksheetFunction.F_Inv(1 - alpha, i, nSamp - i) / (nSamp - i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set getCritT2 = oResult
End Function

Public Function getResiduals(nPcs) As Matrix
    'Given a number of primary components
    nVar = oDataset.NumberOfVariables
    Dim oSelectedloadings As Matrix
    Set oSelectedloadings = Loadings.Submatrix(1, nVar, 1, nPcs)
    Set getResiduals = oDataset.Data.Multiply(Identity(nVar).Subtract(oSelectedloadings.Multiply(oSelectedloadings.Transpose())))
End Function

Public Function OutputModelData()
    nVar = oDataset.NumberOfVariables
    nSamp = oDataset.NumberOfSamples
    Dim oOutput As Matrix
    Set oOutput = ScalarMatrix("", nVar + nSamp + 14, nVar * 4 + 7)
    Dim oSampleNames As New Matrix
    oSampleNames.Mat = oDataset.SampleNames
    
    Dim oVariableNames As New Matrix
    oVariableNames.Mat = oDataset.VariableNames
    
    Dim vPC As Variant
    ReDim vPC(1 To oDataset.NumberOfVariables, 1 To 1)
    For i = 1 To oDataset.NumberOfVariables
        vPC(i, 1) = "PC" & i
    Next i
    
    Dim oPC As New Matrix
    oPC.Mat = vPC

    ' Position Elements
    ' Top Row
    Set oOutput = oOutput.Embed(1, 1, ScalarMatrix("Data", 1, 1))
    Set oOutput = oOutput.Embed(1, 2, oVariableNames.Transpose())
    Set oOutput = oOutput.Embed(1, nVar + 3, ScalarMatrix("Component Score", 1, 1))
    Set oOutput = oOutput.Embed(1, nVar + 4, oPC.Transpose())
    Set oOutput = oOutput.Embed(1, nVar * 2 + 5, ScalarMatrix("Q Residuals", 1, 1))
    Set oOutput = oOutput.Embed(1, nVar * 2 + 6, oPC.Transpose())
    Set oOutput = oOutput.Embed(1, nVar * 3 + 7, ScalarMatrix("T² Distance", 1, 1))
    Set oOutput = oOutput.Embed(1, nVar * 3 + 8, oPC.Transpose())
    
    ' Top block of data
    Set oOutput = oOutput.Embed(2, 2, oNormalData)
    Set oOutput = oOutput.Embed(2, nVar + 3, oSampleNames)
    Set oOutput = oOutput.Embed(2, nVar + 4, oScores)
    Set oOutput = oOutput.Embed(2, nVar * 2 + 5, oSampleNames)
    Set oOutput = oOutput.Embed(2, nVar * 2 + 6, getQ())
    Set oOutput = oOutput.Embed(2, nVar * 3 + 7, oSampleNames)
    Set oOutput = oOutput.Embed(2, nVar * 3 + 8, getT2())
    
    ' Second data block
    Set oOutput = oOutput.Embed(nSamp + 3, 2, oPC.Transpose())
    Set oOutput = oOutput.Embed(nSamp + 3, nVar + 3, ScalarMatrix("Var/Covar", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + 3, nVar + 4, oVariableNames.Transpose())
    Set oOutput = oOutput.Embed(nSamp + 4, 2, oEigen.EigenVec)
    Set oOutput = oOutput.Embed(nSamp + 4, nVar + 3, oVariableNames)
    Set oOutput = oOutput.Embed(nSamp + 4, nVar + 4, oCorrelation)
    
    'Third Data Block
    Set oOutput = oOutput.Embed(nSamp + nVar + 5, 2, oPC.Transpose())
    Set oOutput = oOutput.Embed(nSamp + nVar + 6, 2, RowMatrix(oEigen.EigenVal))
    Set oOutput = oOutput.Embed(nSamp + nVar + 7, 2, RowMatrix(getRsq()))
    Set oOutput = oOutput.Embed(nSamp + nVar + 8, 2, RowMatrix(getCumRsq()))
    Set oOutput = oOutput.Embed(nSamp + nVar + 9, 2, RowMatrix(getRESS()))
    Set oOutput = oOutput.Embed(nSamp + nVar + 13, 2, RowMatrix(getCritT2()))
    
    ' Left Column
    Set oOutput = oOutput.Embed(2, 1, oSampleNames)
    Set oOutput = oOutput.Embed(nSamp + 3, 1, ScalarMatrix("Loading", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + 4, 1, oVariableNames)
    Set oOutput = oOutput.Embed(nSamp + nVar + 6, 1, ScalarMatrix("Eval", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 7, 1, ScalarMatrix("R²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 8, 1, ScalarMatrix("CumR²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 9, 1, ScalarMatrix("RESS", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 10, 1, ScalarMatrix("PRESS", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 11, 1, ScalarMatrix("Q²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 12, 1, ScalarMatrix("CumQ²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 13, 1, ScalarMatrix("critT²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nVar + 14, 1, ScalarMatrix("critQ", 1, 1))
    
    OutputModelData = oOutput.Mat
End Function
