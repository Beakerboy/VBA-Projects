
End SubDim oDataset As Dataset
Dim oNewDataset As Dataset
Dim oNormalData As Matrix
Dim oNormalNewData As Matrix
Dim oCorrelation As Matrix
Dim oEigen As Eigen
Dim oTransformed As Matrix
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

Property Get TransformedData()
    TransformedData = oTransformed.Mat
End Property

' Function: Normalize
' Normalise the dataset
'
' Parameters:
'  Cov - If true, do not divide each column by its standard deviation
'
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
    Set oTransformed = oNormalData.Multiply(oEigen.EigenVec)
End Sub

' Function TransformNewData
' Use the loading matrix to transform new data
' Parameters:
'  inData - New Data
'  iNum - Number of PCs to use
Public Function TransformNewData(inData As Variant, Optional iNum As Integer = 0)
    'Normalize inData
    If iNum = 0 Then
        iNum = vData.N
        Set oLoadings = oEigenVec
    Else
        Set oLoadings = oEigenVec.Submatrix(oDataset.NumberOfSamples, iNum)
    End If

    Dim oInData As New Matrix
    oInData.Mat = inData
    Set TransformNewData = oInData.Multiply(oLoadings)
End Function

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

Public Function getRESS()
    Dim vRsq As Variant
    vRsq = getCumRsq()
    Inertia = oEigen.EigenVal.Sum * (oDataset.NumberOfSamples - 1)
    For i = 1 To oDataset.NumberOfVariables
        vRsq(i, 1) = (1 - vRsq(i, 1)) * Inertia
    Next i
    getRESS = vRsq
End Function

Public Function getRsq()
    getRsq = oEigen.EigenVal.ScalarDivide(oEigen.EigenVal.Sum).Vec
End Function

Public Function getCumRsq()
    Dim vRsq As Variant
    vRsq = getRsq()
    For i = 2 To oDataset.NumberOfVariables
        vRsq(i, 1) = vRsq(i, 1) + vRsq(i - 1, 1)
    Next i
    getCumRsq = vRsq
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
        Set ReducedScores = oTransformed.Submatrix(1, oDataset.NumberOfSamples, 1, i)
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

Public Function getResiduals(nPcs) As Matrix
    'Given a number of primary components
    nVar = oDataset.NumberOfVariables
    Dim oSelectedloadings As Matrix
    Set oSelectedloadings = Loadings.Submatrix(1, nVar, 1, nPcs)
    Set getResiduals = oDataset.Data.Multiply(Identity(nVar).Subtract(oSelectedloadings.Multiply(oSelectedloadings.Transpose())))
End Function

Public Function OutputModelData()

    Dim oSampleNames As New Matrix
    oSampleNames.Mat = oDataset.SampleNames
    
    Dim vZeroRow As Variant
    ReDim vZeroRow(1 To 1, 1 To 1)
    
    Dim oZeroRow As Matrix
    Set oZeroRow = New Matrix
    oZeroRow.Mat = vZeroRow
    
    Dim oVariableNames As New Matrix
    oVariableNames.Mat = oDataset.VariableNames

    Dim oOutput As Matrix
    Set oOutput = oNormalData
    
    Dim oZeroColumn As New Matrix
    
    Set oOutput = oOutput.AugmentRight(ScalarMatrix("", oOutput.M, 1))
    Set oOutput = oOutput.AugmentRight(oSampleNames)
    Set oOutput = oOutput.AugmentRight(oTransformed)
    Set oOutput = oOutput.AugmentRight(ScalarMatrix("", oOutput.M, 1))
    Set oOutput = oOutput.AugmentRight(getQ())
    Dim vPC As Variant
    ReDim vPC(1 To oDataset.NumberOfVariables, 1 To 1)
    For i = 1 To oDataset.NumberOfVariables
        vPC(i, 1) = "PC" & i
    Next i
    
    Dim oPC As New Matrix
    oPC.Mat = vPC
    
    Dim oTitleRow As Matrix
    Set oTitleRow = oPC.Transpose()
    Set oTitleRow = oTitleRow.AugmentRight(ScalarMatrix("", 1, 2))
    Set oTitleRow = oTitleRow.AugmentRight(oVariableNames.Transpose())
    Set oTitleRow = oTitleRow.AugmentRight(ScalarMatrix("", 1, oDataset.NumberOfVariables + 1))
    Set oOutput = oOutput.AugmentBelow(ScalarMatrix("", 1, oOutput.N))
    Set oOutput = oOutput.AugmentBelow(oTitleRow)
    
    Set oZeroRow = ScalarMatrix("", 1, oOutput.N)
    
    Dim NextBlock As Matrix
    Set NextBlock = oEigen.EigenVec
    
    Set NextBlock = NextBlock.AugmentRight(ScalarMatrix("", NextBlock.M, 1))
    Set NextBlock = NextBlock.AugmentRight(oVariableNames)
    Set NextBlock = NextBlock.AugmentRight(oCorrelation)
    Set NextBlock = NextBlock.AugmentRight(ScalarMatrix("", NextBlock.M, oOutput.N - NextBlock.N))

    
    Set oOutput = oOutput.AugmentBelow(NextBlock)
    Set oOutput = oOutput.AugmentBelow(oZeroRow)
    Dim vEvalRow() As Variant
    vEvalRow = oEigen.EigenVal.Vec
    
    Dim oEvalRow As New Matrix
    oEvalRow.Mat = vEvalRow
    
    Set oZeroRow = ScalarMatrix("", oOutput.N - oEvalRow.M, 1)
    
    Dim vRsqRow() As Variant
    vRsqRow = Me.getRsq()
    
    Dim oRsqRow As New Matrix
    oRsqRow.Mat = vRsqRow
    
    Dim vCumRsqRow() As Variant
    vCumRsqRow = Me.getCumRsq()
    Dim oCumRsqRow As New Matrix
    oCumRsqRow.Mat = vCumRsqRow
    
    Dim vRESSRow() As Variant
    vRESSRow = Me.getRESS()
    Dim oRESSRow As New Matrix
    oRESSRow.Mat = vRESSRow
    
    Set oEvalRow = oEvalRow.AugmentBelow(oZeroRow).Transpose()
    Set oRsqRow = oRsqRow.AugmentBelow(oZeroRow).Transpose()
    Set oCumRsqRow = oCumRsqRow.AugmentBelow(oZeroRow).Transpose()
    Set oRESSRow = oRESSRow.AugmentBelow(oZeroRow).Transpose()
    Set oOutput = oOutput.AugmentBelow(oEvalRow)
    Set oOutput = oOutput.AugmentBelow(oRsqRow)
    Set oOutput = oOutput.AugmentBelow(oCumRsqRow)
    Set oOutput = oOutput.AugmentBelow(oRESSRow)
    
    Dim zeroElement As Matrix
    Set zeroElement = ScalarMatrix("", 1, 1)
    
    Set oSampleNames = oSampleNames.AugmentBelow(ScalarMatrix("", 2, 1)).AugmentBelow(oVariableNames).AugmentBelow(zeroElement).AugmentBelow(ScalarMatrix("Eval", 1, 1)).AugmentBelow(ScalarMatrix("R²", 1, 1)).AugmentBelow(ScalarMatrix("CumR²", 1, 1)).AugmentBelow(ScalarMatrix("RESS", 1, 1))
    Set oOutput = oSampleNames.AugmentRight(oOutput)
    
    Dim oTopRow As Matrix
    Set oTopRow = zeroElement.AugmentRight(oVariableNames.Transpose).AugmentRight(zeroElement).AugmentRight(zeroElement).AugmentRight(oPC.Transpose).AugmentRight(zeroElement).AugmentRight(oPC.Transpose)
    Set oOutput = oTopRow.AugmentBelow(oOutput)
    OutputModelData = oOutput.Mat
    
End Function

