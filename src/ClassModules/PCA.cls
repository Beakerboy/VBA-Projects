' Class: PCA
' Perform Primary Component Analysis on a set of data
'
' Version: 19091301

Dim oDataset As Dataset
Dim oNewDataset As Dataset
Dim oNormalData As Matrix
Dim oNormalNewData As Matrix
Dim oCorrelation As Matrix
Dim oEigen As Eigen
Dim oScores As Matrix
Dim bScale As Boolean
Dim bCenter As Boolean
Dim bCV As Boolean
Dim dInertia As Double

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

' Property: CrossValidate
' If set to true, calculate crossvalidated T², Q-residuals, and PRESS
Property Let CrossValidate(bSet As Boolean)
    bCV = bSet
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
    Dim Divisor As Double
    Set oCorrelationt = oNormalData.Transpose()
    Divisor = iSamples - 1
    Set oCorrelation = oCorrelationt.Multiply(oNormalData).ScalarDivide(Divisor)
End Sub

Public Sub Run()
    Normalize
    dInertia = oNormalData.frobeniousNorm()
    CorrelationMatrix
    Set oEigen = New Eigen
    Set oEigen.Matrix = oCorrelation
    oEigen.JK
    Transform
    If bCV Then RunCV
End Sub


Public Sub RunCV()

End Sub

Public Sub Transform(Optional iNum As Integer)
    Set oScores = oNormalData.Multiply(oEigen.EigenVec)
End Sub

Public Function getPRESS() As Vector
    nSamp = oDataset.NumberOfSamples
    nvar = oDataset.NumberOfVariables
    Dim oResults As Matrix
    
    Set oResults = ScalarMatrix(0, nSamp, nvar)
    Dim pressPCA As New PCA
    Dim oTrainingDataset As New Dataset
    Dim oTestingDataset As New Dataset
    With pressPCA
        .PCACenter = True
        .PCAScale = True
    End With
    For i = 1 To nSamp - 1
        oTrainingDataset.Data = oDataset.Data.ExcludeRow(i).Mat
        Set pressPCA.Data = oTrainingDataset
        pressPCA.Run
        oTestingDataset.Data = oDataset.Data.GetRow(i).toRowMatrix().Mat
        Set pressPCA.NewData = oTestingDataset
        Set oResults = oResults.ReplaceRow(i, pressPCA.getQ(True).GetRow(1))
    Next i
    
    Dim vResultsVector As Variant
    ReDim vResultsVector(1 To nvar, 1 To 1)
    For i = 1 To nvar
        vResultsVector(i, 1) = oResults.GetColumn(i).Sum()
    Next i
    
    Dim oResultsVector As New Vector
    oResultsVector.Vec = vResultsVector
    Set getPRESS = oResultsVector
End Function

Public Function getRESS() As Vector
    Dim vRsq As Variant
    vRsq = getCumRsq().Vec
    Inertia = oEigen.EigenVal.Sum * (oDataset.NumberOfSamples - 1)
    For i = 1 To oDataset.NumberOfVariables
        vRsq(i, 1) = (1 - vRsq(i, 1)) * dInertia
    Next i
    Dim oResult As New Vector
    oResult.Vec = vRsq
    Set getRESS = oResult
End Function

Public Function getRsq() As Vector
    Set getRsq = oEigen.EigenVal.ScalarDivide(oCorrelation.Trace())
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

' Function: getQ
' Calculate the Q residuals for each sample, for all model complexities
' For each sample (i), the Q residual is
' Qi = ei' * ei = xi' (I - PP')xi
Public Function getQ(Optional bUseNewdata = False) As Matrix
    Dim oIdentity, oLoadings, oStdevs As Matrix
    Dim oScaledX As New Matrix
    Dim oAverages As New Matrix
    If bUseNewdata Then Set oScaledX = oNormalNewData Else Set oScaledX = oNormalData

    Dim vResult As Variant
    Dim oResult As Matrix
    
    ' Initialize oResult to be a matrix with the same shape as oSacledX
    Set oResult = oScaledX.ScalarMultiply(0)
    nvar = oDataset.NumberOfVariables
    ReDim vResult(1 To 1, 1 To nvar)
    Set oIdentity = Identity(nvar)
    For i = 1 To oDataset.NumberOfVariables
        Set oLoadings = oEigen.EigenVec.Submatrix(1, nvar, 1, i)
        Set oResult = oResult.ReplaceColumn(i, oScaledX.Multiply(oIdentity.Subtract(oLoadings.Multiply(oLoadings.Transpose()))).Multiply(oScaledX.Transpose()).GetDiagonalElements())
    Next i
    Set getQ = oResult
End Function

' Function: getT2
' Get the T2 distances of the model or new data
' T²i = sum,j=1 to r(ti,j / Eval,j) = ti * diag(eval)^-1 * ti'
'
' Parameters:
'   bUseNewData - True if the new data dataset object is to be used
Public Function getT2(Optional bUseNewdata = False) As Matrix
    nSamp = oDataset.NumberOfSamples
    nvar = oDataset.NumberOfVariables
    Dim oResult, oInverseEvals As Matrix
    Dim oReducedInverseEvals As Matrix
    Dim oReducedScores As Matrix
    Dim ithT2 As Vector
    Set oResult = ScalarMatrix(0, nSamp, nvar)
    Set oInverseEvals = DiagonalMatrix(oEigen.EigenVal).Inverse()
    For i = 1 To nvar
        Set oReducedInverseEvals = oInverseEvals.Submatrix(1, i, 1, i)
        Set oReducedScores = oScores.Submatrix(1, nSamp, 1, i)
        Set ithT2 = oReducedScores.Multiply(oReducedInverseEvals).Multiply(oReducedScores.Transpose()).GetDiagonalElements()
        Set oResult = oResult.ReplaceColumn(i, ithT2)
    Next i
    Set getT2 = oResult
End Function

Public Function getCritT2(Optional alpha = 0.05) As Vector
    nSamp = oDataset.NumberOfSamples
    nvar = oDataset.NumberOfVariables
    Dim vResult As Variant
    ReDim vResult(1 To nvar, 1 To 1)
    For i = 1 To nvar
        vResult(i, 1) = i * (nSamp - 1) * WorksheetFunction.F_Inv(1 - alpha, i, nSamp - i) / (nSamp - i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set getCritT2 = oResult
End Function

Public Function getCritQ(Optional alpha = 0.05) As Vector
    nSamp = oDataset.NumberOfSamples
    nvar = oDataset.NumberOfVariables
    Dim vResult As Variant
    ReDim vResult(1 To nvar, 1 To 1)
    Dim eval_subset As Vector
    Dim eval_subset2 As Vector
    Dim eval_subset3 As Vector
    For i = 1 To nvar - 1
        
        Set eval_subset = oEigen.EigenVal.subVector(i + 1)
        Set eval_subset2 = eval_subset.pow(2)
        Set eval_subset3 = eval_subset.pow(3)
        t1 = eval_subset.Sum()
        t2 = eval_subset2.Sum()
        t3 = eval_subset3.Sum()
        h0 = 1 - 2 * t1 * t3 / 3 / t2 / t2
        If h0 < 0.001 Then
            h0 = 0.001
        End If
        ca = WorksheetFunction.Norm_S_Inv(1 - alpha)
        h1 = ca * Math.sqr(2 * t2 * h0 * h0) / t1
        h2 = t2 * h0 * (h0 - 1) / t1 / t1
        vResult(i, 1) = t1 * (1 + h1 + h2) ^ (1 / h0)
    Next i
    vResult(nvar, 1) = 0
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set getCritQ = oResult
End Function

Public Function getResiduals(nPcs) As Matrix
    'Given a number of primary components
    nvar = oDataset.NumberOfVariables
    Dim oSelectedloadings As Matrix
    Set oSelectedloadings = Loadings.Submatrix(1, nvar, 1, nPcs)
    Set getResiduals = oDataset.Data.Multiply(Identity(nvar).Subtract(oSelectedloadings.Multiply(oSelectedloadings.Transpose())))
End Function

Public Function OutputModelData()
    nvar = oDataset.NumberOfVariables
    nSamp = oDataset.NumberOfSamples
    Dim oOutput As Matrix
    Set oOutput = ScalarMatrix("", nvar + nSamp + 17, nvar * 4 + 7)
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
    Set oOutput = oOutput.Embed(1, 1, ScalarMatrix("Variables", 1, 1))
    Set oOutput = oOutput.Embed(1, 2, ScalarMatrix("Samples", 1, 1))
    Set oOutput = oOutput.Embed(1, 3, ScalarMatrix("Inertia", 1, 1))
    
    'Summary Data
    Set oOutput = oOutput.Embed(2, 1, ScalarMatrix(oDataset.NumberOfVariables, 1, 1))
    Set oOutput = oOutput.Embed(2, 2, ScalarMatrix(oDataset.NumberOfSamples, 1, 1))
    Set oOutput = oOutput.Embed(2, 3, ScalarMatrix(dInertia, 1, 1))
    
    'Top Data
    Set oOutput = oOutput.Embed(4, 1, ScalarMatrix("Data", 1, 1))
    Set oOutput = oOutput.Embed(4, 2, oVariableNames.Transpose())
    Set oOutput = oOutput.Embed(4, nvar + 3, ScalarMatrix("Component Score", 1, 1))
    Set oOutput = oOutput.Embed(4, nvar + 4, oPC.Transpose())
    Set oOutput = oOutput.Embed(4, nvar * 2 + 5, ScalarMatrix("Q Residuals", 1, 1))
    Set oOutput = oOutput.Embed(4, nvar * 2 + 6, oPC.Transpose())
    Set oOutput = oOutput.Embed(4, nvar * 3 + 7, ScalarMatrix("T² Distance", 1, 1))
    Set oOutput = oOutput.Embed(4, nvar * 3 + 8, oPC.Transpose())
    
    ' Top block of data
    Set oOutput = oOutput.Embed(5, 2, oNormalData)
    Set oOutput = oOutput.Embed(5, nvar + 3, oSampleNames)
    Set oOutput = oOutput.Embed(5, nvar + 4, oScores)
    Set oOutput = oOutput.Embed(5, nvar * 2 + 5, oSampleNames)
    Set oOutput = oOutput.Embed(5, nvar * 2 + 6, getQ())
    Set oOutput = oOutput.Embed(5, nvar * 3 + 7, oSampleNames)
    Set oOutput = oOutput.Embed(5, nvar * 3 + 8, getT2())
    
    ' Second data block
    Set oOutput = oOutput.Embed(nSamp + 6, 2, oPC.Transpose())
    Set oOutput = oOutput.Embed(nSamp + 6, nvar + 3, ScalarMatrix("Var/Covar", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + 6, nvar + 4, oVariableNames.Transpose())
    Set oOutput = oOutput.Embed(nSamp + 7, 2, oEigen.EigenVec)
    Set oOutput = oOutput.Embed(nSamp + 7, nvar + 3, oVariableNames)
    Set oOutput = oOutput.Embed(nSamp + 7, nvar + 4, oCorrelation)
    'If bCV Then Set oOutput = oOutput.Embed(nSamp + 7, 2 * nVar + 5, getQCV())
    'If bCV Then Set oOutput = oOutput.Embed(nSamp + 7, 3 * nVar + 8, getT2CV())
    
    'Third Data Block
    Set oOutput = oOutput.Embed(nSamp + nvar + 8, 2, oPC.Transpose())
    Set oOutput = oOutput.Embed(nSamp + nvar + 9, 2, RowMatrix(oEigen.EigenVal))
    Set oOutput = oOutput.Embed(nSamp + nvar + 10, 2, RowMatrix(getRsq()))
    Set oOutput = oOutput.Embed(nSamp + nvar + 11, 2, RowMatrix(getCumRsq()))
    Set oOutput = oOutput.Embed(nSamp + nvar + 12, 2, RowMatrix(getRESS()))
    If bCV Then Set oOutput = oOutput.Embed(nSamp + nvar + 13, 2, RowMatrix(getPRESS()))
    Set oOutput = oOutput.Embed(nSamp + nvar + 16, 2, RowMatrix(getCritT2()))
    Set oOutput = oOutput.Embed(nSamp + nvar + 17, 2, RowMatrix(getCritQ()))
    
    ' Left Column
    Set oOutput = oOutput.Embed(5, 1, oSampleNames)
    Set oOutput = oOutput.Embed(nSamp + 6, 1, ScalarMatrix("Loading", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + 7, 1, oVariableNames)
    Set oOutput = oOutput.Embed(nSamp + nvar + 9, 1, ScalarMatrix("Eval", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 10, 1, ScalarMatrix("R²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 11, 1, ScalarMatrix("CumR²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 12, 1, ScalarMatrix("RESS", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 13, 1, ScalarMatrix("PRESS", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 14, 1, ScalarMatrix("Q²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 15, 1, ScalarMatrix("CumQ²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 16, 1, ScalarMatrix("critT²", 1, 1))
    Set oOutput = oOutput.Embed(nSamp + nvar + 17, 1, ScalarMatrix("critQ", 1, 1))
    
    OutputModelData = oOutput.Mat
End Function
