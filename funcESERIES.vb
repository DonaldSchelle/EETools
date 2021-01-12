Imports ExcelDna.Integration

Public Module funcESERIES

    'EXCEL/PUBLIC FUNCTIONS
    '----------------------

    'Function to calculate the closest E-series value
    <ExcelFunction(Name:="ESERIES",
                   Category:="Engineering-ESERIES",
                   Description:="Calculates closest standard E-series value",
                   HelpTopic:="http://github.com/DonaldSchelle/EETools")>
    Public Function ESERIES(
        <ExcelArgument(Name:="dValue", Description:="Value to convert")> dValue As Double,
        <ExcelArgument(Name:="sSeries", Description:="Desired series (E3-E192)")> sSeries As String,
        <ExcelArgument(Name:="[iRound]", Description:="Rounding Type: -1 = Next Lowest, (0 = Closest), +1 = Next Highest")> Optional iRound As Integer = 0,
        <ExcelArgument(Name:="[iCalcType]", Description:="Method for determining offset: (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic")> Optional iCalcType As Integer = 0,
        <ExcelArgument(Name:="[dMinLimit]", Description:="Minimum Returned Value (default: 1)")> Optional dMinLimit As Double = 1,
        <ExcelArgument(Name:="[dMaxLimit]", Description:="Maximum Returned Value (default: 1,000,000)")> Optional dMaxLimit As Double = 1000000)

        'Function to return closest standard resistor value
        'dValue = target value (must be greater than 0)
        'sSeries (String) = series to be used
        '[iRound] = Rounding option 1 = next highest, (0 = closest), -1 = next lowest
        '[iCalcType] = Used to determine closest value (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic
        '[dMinLimit] = Minimum returned value (1)
        '[dMaxLimit] = Maximum returned value (1000000)

        'Initialize Variables
        Dim resValue As Double  'Resistor E-Series Value
        Dim minValue As Double  'Minimum E-Series Limit
        Dim maxValue As Double  'Maximum E-Series Limit
        Dim eValues() As Double 'Array that holds selected E-Series array

        'Initialize optional parameters if empty
        If iRound = vbEmpty Then iRound = 0
        If iCalcType = vbEmpty Then iCalcType = 0
        If dMinLimit = vbEmpty Then dMinLimit = 1
        If dMaxLimit = vbEmpty Then dMaxLimit = 1000000

        'Error checking of input values
        If dValue <= 0 Then Return ExcelError.ExcelErrorValue
        If (iRound < -1) Or (iRound > +1) Then Return ExcelError.ExcelErrorValue
        If (iCalcType < 0) Or (iCalcType > 3) Then Return ExcelError.ExcelErrorValue
        If dMinLimit <= 0 Then Return ExcelError.ExcelErrorValue
        If dMaxLimit <= 0 Then Return ExcelError.ExcelErrorValue

        'Correct for string values with lower case letters (i.e. "e24" vs "E24") and error check
        sSeries = sSeries.ToUpper
        Select Case (sSeries)
            Case "E3" : eValues = E3
            Case "E6" : eValues = E6
            Case "E12" : eValues = E12
            Case "E24" : eValues = E24
            Case "E48" : eValues = E48
            Case "E96" : eValues = E96
            Case "E192" : eValues = E192
            Case Else
                'An invalid string was entered.   Return appropriate Excel error code
                Return ExcelError.ExcelErrorName
        End Select

        'Find E-Series resistor
        resValue = simpleESERIES(dValue, eValues, iRound, iCalcType)
        If resValue = 0 Then Return ExcelError.ExcelErrorNA     'Error Something went wrong converting value

        'Find eseries value of min/max Limits
        minValue = simpleESERIES(dMinLimit, eValues, +1, iCalcType)
        maxValue = simpleESERIES(dMaxLimit, eValues, -1, iCalcType)
        'Check for errors converting min/max value, and input errors
        If (minValue = 0) Or (maxValue = 0) Or (maxValue <= minValue) Then Return ExcelError.ExcelErrorValue

        'check calcualted E-series value against min/max values
        Select Case resValue
            Case < minValue         'Resistor value is below minimum
                Return minValue
            Case > maxValue         'Resistor value is above maximum
                Return maxValue
            Case Else               'Resistor value is within limits
                Return resValue
        End Select

    End Function

    'Function to calculate closest E-series Resistor combination
    <ExcelFunction(Name:="ESeriesResistorPair",
                   Category:="Engineering-ESERIES",
                   Description:="Calculates two resistor values that combine to the closest target value",
                   HelpTopic:="http://github.com/DonaldSchelle/EETools")>
    Public Function ESeriesResistorPair(
        <ExcelArgument(Name:="dTarget", Description:="Target resistance value.  Must be greater than 0.")> dTarget As Double,
        <ExcelArgument(Name:="sSeries", Description:="Desired E-series (E3-E192)")> sSeries As String,
        <ExcelArgument(Name:="iComboType", Description:="Combination Type.  0 = Series, 1 = Parallel")> iComboType As Integer,
        <ExcelArgument(Name:="iReturnValue", Description:="Return Value: 0 = Array, 1 = Resistor #1, 2 = Resistor #2, 3 = Remainder")> iReturnValue As Integer,
        <ExcelArgument(Name:="[iRound]", Description:="Return Value Rounding: -1 = Force Lower, (0 = Closest), +1 = Force Higher")> Optional iRound As Integer = 0,
        <ExcelArgument(Name:="[iCalcType]", Description:="Method for determining offset: (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic")> Optional iCalcType As Integer = 0,
        <ExcelArgument(Name:="[dMatchType]", Description:="Matching Criteria: -1 = Comprehensive/Best Match, (0 = Quick Match), >0 = Preferred Value Match")> Optional dMatchType As Double = 0,
        <ExcelArgument(Name:="[dCompMinLimit]", Description:="Minimum Individual Component Value (1)")> Optional dCompMinLimit As Double = 1,
        <ExcelArgument(Name:="[dCompMaxLimit]", Description:="Maximum Individual Component Value (1,000,000)")> Optional dCompMaxLimit As Double = 1000000)

        'Function to return closest E-series values that combine into the closet value to dTarget
        'dTarget = target value (must be greater than 0)
        'sSeries(Of String) = E-Series To be used
        'iComboType = Series(0)/Parallel(1) combination
        'iReturnValue = Return value from function, array or chosen single value
        '[iRound] = Rounding option 1 = next highest, (0 = closest), -1 = next lowest
        '[iCalcType] = Used to determine closest value (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic
        '[dMatchType] = How the matching is done: (-1 = find best pair), 0 = stop at first perfect match, >0 = find pair with value closest to dMatchType
        '[dCompMinLimit] = Minimum Component Value (1)
        '[dCompMaxLimit] = Maximum Component Value (1,000,000)
        'Limits are to ensure that returned values are practical (i.e. no 1uOhm or 10 GOhm resistors).

        'Initialize Variables
        Dim resValues(2) As Double              'Internal Calculation Array
        Dim ResArrayValues(2, 2) As Double      'Return array
        Dim eValues() As Double                 'Array that holds selected E-Series array

        'Initialize optional parameters if empty
        If iRound = vbEmpty Then iRound = 0
        If iCalcType = vbEmpty Then iCalcType = 0
        If dMatchType = vbEmpty Then dMatchType = 0
        If dCompMinLimit = vbEmpty Then dCompMinLimit = 1
        If dCompMaxLimit = vbEmpty Then dCompMaxLimit = 1000000

        'Error checking of input values
        If dTarget <= 0 Then Return ExcelError.ExcelErrorValue
        If (iRound < -1) Or (iRound > +1) Then Return ExcelError.ExcelErrorValue
        If (iCalcType < 0) Or (iCalcType > 3) Then Return ExcelError.ExcelErrorValue
        If dMatchType < -1 Then Return ExcelError.ExcelErrorValue
        If dCompMinLimit <= 0 Then Return ExcelError.ExcelErrorValue
        If dCompMaxLimit <= 0 Then Return ExcelError.ExcelErrorValue
        If dCompMaxLimit <= dCompMinLimit Then Return ExcelError.ExcelErrorValue

        sSeries = sSeries.ToUpper    'Correct for string values with lower case letters (i.e. "e24" vs "E24")
        Select Case sSeries
            Case "E3" : eValues = E3
            Case "E6" : eValues = E6
            Case "E12" : eValues = E12
            Case "E24" : eValues = E24
            Case "E48" : eValues = E48
            Case "E96" : eValues = E96
            Case "E192" : eValues = E192
            Case Else
                Return ExcelError.ExcelErrorName
        End Select

        'These variables are checked below
        'If (iComboType < 0) Or (iComboType > 1) Then Return ExcelError.ExcelErrorValue  'Not needed as it's done below
        'If (iReturnValue < 0) Or (iReturnValue > 2) Then Return ExcelError.ExcelErrorValue  'Not needed as it's done below

        'Determine ComboType (Series or Parallel)
        Select Case iComboType
            Case 0     'Series Combination
                resValues = ComboSeriesResistor(dTarget, eValues, iRound, iCalcType, dMatchType, dCompMinLimit, dCompMaxLimit)
                'Calculate Remainder
                resValues(2) = dTarget - resValues(0) - resValues(1)

            Case 1      'Parallel Combination
                resValues = ComboParallelResistor(dTarget, eValues, iRound, iCalcType, dMatchType, dCompMinLimit, dCompMaxLimit)
                'Calculate Remainder
                resValues(2) = dTarget - (resValues(0) * resValues(1)) /
                                         (resValues(0) + resValues(1))

            Case Else
                Return ExcelError.ExcelErrorValue   'Invalid combo type
        End Select

        'Check if return values are valid
        If resValues(0) = 0 Then Return ExcelError.ExcelErrorNA     'No suitable value found

        'Determine return value
        Select Case iReturnValue
            Case 0                                  'Return Array of Values
                'Structure array
                ResArrayValues(0, 0) = resValues(0)
                ResArrayValues(1, 0) = resValues(1)
                ResArrayValues(0, 1) = resValues(1)
                ResArrayValues(2, 0) = resValues(2)
                ResArrayValues(0, 2) = resValues(2)
                Return ResArrayValues
            Case 1                                  'Return value for Resistor #1
                Return resValues(0)
            Case 2                                  'Return value for Resistor #2
                Return resValues(1)
            Case 3                                  'Return value for Remainder
                Return resValues(2)
            Case Else
                Return ExcelError.ExcelErrorValue   'Invalid return type
        End Select

    End Function

    'Function to calculate the closest E-series values To generate a precise ratio based on a resistor configuration
    <ExcelFunction(Name:="ESeriesResistorRatio",
                   Category:="Engineering-ESERIES",
                   Description:="Calculates E-Series divider values to achieve closest ratio",
                   HelpTopic:="http://github.com/DonaldSchelle/EETools")>
    Public Function ESeriesResistorRatio(
        <ExcelArgument(Name:="dNumerator", Description:="Numerator")> dNumerator As Double,
        <ExcelArgument(Name:="dDenominator", Description:="Denominator")> dDemoninator As Double,
        <ExcelArgument(Name:="sSeries", Description:="Desired E-series (E3-E192)")> sSeries As String,
        <ExcelArgument(Name:="iRatioType", Description:="RatioType.  0 = Simple Ratio, 1 = Voltage Divider (dRatio > 1)")> iRatioType As Integer,
        <ExcelArgument(Name:="iReturnValue", Description:="Return Value: 0 = Array, 1 = Primary, 2 = Secondary, 3 = Tertiary, 4 = Ratio Offset")> iReturnValue As Integer,
        <ExcelArgument(Name:="[iRound]", Description:="Rounding: -1 = closest lower ratio, (0 = closest ratio), +1 = closest higher ratio")> Optional iRound As Integer = 0,
        <ExcelArgument(Name:="[iCalcType]", Description:="Method for determining offset: (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic")> Optional iCalcType As Integer = 0,
        <ExcelArgument(Name:="[iElements]", Description:="(0 = two components), three components: series pair (1 = bottom, 2 = top), parallel pair (3 = bottom, 4 = top)")> Optional iElements As Integer = 0,
        <ExcelArgument(Name:="[dMatchType]", Description:="Thevenin Matching: -1 = Match towards middle of range, (0 = Quick Match), >0 = Match to Preferred Value")> Optional dMatchType As Double = 0,
        <ExcelArgument(Name:="[dThevMinLimit]", Description:="Minimum Thevinen Resistance of network (default: 1,000)")> Optional dTotalMinLimit As Double = 1000,
        <ExcelArgument(Name:="[dThevMaxLimit]", Description:="Maximum Thevinen Resistance of network (default: 100,000)")> Optional dTotalMaxLimit As Double = 100000,
        <ExcelArgument(Name:="[dCompMinLimit]", Description:="Minimum Individual Component Value (default: 1)")> Optional dCompMinLimit As Double = 1,
        <ExcelArgument(Name:="[dCompMaxLimit]", Description:="Maximum Individual Component Value (default: 1,000,000)")> Optional dCompMaxLimit As Double = 1000000)

        'Function to return closest E-series values to generate a precise ratio
        'Function modeled somewhat to this: http://jansson.us/resistors.html#ratioCalc
        'dRatio = Target ratio
        'iRatioType:  0 = simple ratio, 1 = voltage divider ratio (dRatio > 1)
        'sSeries = E-series to be used
        'iReturnValue:  0 = Return upper component value, 1 = Return lower component value, 2 = Return lower secondary component value
        '[iRound] = 1 = closest higher ratio, (0 = closest), -1 = closest lower ratio
        '[iCalcType] = Used to determine closest value (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic
        '[iElements] = (0 = two components), 1 = three components (series pair bottom), 2 = three components (parallel pair bottom)
        '[dMatchType] = How series/parallel matching is done: -1 = find best pair, 0 = stop at first perfect match, >0 = find pair with value closest to dMatchType
        '[dTotalMinLimit] = Minimum sum of components in network (1,000)
        '[dTotalMaxLimit] = Maximum sum of components in network (100,000)
        '[dCompMinLimit] = Minimum returned component value (1)
        '[dCompMaxLimit] = Maximum returned component value (1000000)
        'Limits are to ensure that all returned value are within practical limits (i.e. no 1uOhm or 10 GOhm resistors please).


        'Initialize Variables
        Dim resValues(3) As Double
        Dim ResArrayValues(3, 3) As Double      'Return array
        Dim dRatio As Double
        Dim eValues() As Double                 'Array that holds selected E-Series array

        'Initialize optional parameters if empty
        If iRound = vbEmpty Then iRound = 0
        If iCalcType = vbEmpty Then iCalcType = 0
        If iElements = vbEmpty Then iElements = 0
        If dMatchType = vbEmpty Then dMatchType = 0
        If dTotalMinLimit = vbEmpty Then dTotalMinLimit = 1000
        If dTotalMaxLimit = vbEmpty Then dTotalMaxLimit = 100000
        If dCompMinLimit = vbEmpty Then dCompMinLimit = 1
        If dCompMaxLimit = vbEmpty Then dCompMaxLimit = 1000000

        'Error checking of input values
        If dNumerator <= 0 Then Return ExcelError.ExcelErrorValue
        If dDemoninator <= 0 Then Return ExcelError.ExcelErrorValue

        dRatio = dNumerator / dDemoninator

        'Correct for Ratio Type and error check input value
        If (iRatioType = 1) And (dRatio > 1) Then      'Voltage divider ratio type
            'Ratio correction calculation
            dRatio -= 1
        ElseIf (iRatioType = 0) And (dRatio > 0) Then  'Simple ratio type (do nothing)
        Else   'Either iRatioType is not valid, or dValueRatio is out-of-bounds
            Return ExcelError.ExcelErrorValue
        End If

        If (iRound < -1) Or (iRound > +1) Then Return ExcelError.ExcelErrorValue
        If (iCalcType < 0) Or (iCalcType > 3) Then Return ExcelError.ExcelErrorValue
        If (iElements < 0) Or (iElements > 4) Then Return ExcelError.ExcelErrorValue
        If dMatchType < -1 Then Return ExcelError.ExcelErrorValue

        'If (iReturnValue < 0) Or (iReturnValue > 1) Then Return ExcelError.ExcelErrorValue  'Not needed as it's done below
        'If (iComboType < 0) Or (iComboType > 1) Then Return ExcelError.ExcelErrorValue  'Not needed as it's done below

        sSeries = sSeries.ToUpper    'Correct for string values with lower case letters (i.e. "e24" vs "E24")
        Select Case sSeries
            Case "E3" : eValues = E3
            Case "E6" : eValues = E6
            Case "E12" : eValues = E12
            Case "E24" : eValues = E24
            Case "E48" : eValues = E48
            Case "E96" : eValues = E96
            Case "E192" : eValues = E192
            Case Else
                Return ExcelError.ExcelErrorName
        End Select

        If dTotalMinLimit <= 0 Then Return ExcelError.ExcelErrorValue
        If dTotalMaxLimit <= 0 Then Return ExcelError.ExcelErrorValue
        If dCompMinLimit <= 0 Then Return ExcelError.ExcelErrorValue
        If dCompMaxLimit <= 0 Then Return ExcelError.ExcelErrorValue


        'There's no way to be certain that the best two matching values are chosen
        'unless we try all of them.    Let's do that, and create an array with columns
        'Column names: Upper Value, Lower Value, Deviation

        resValues = RatioCalculator(dRatio, eValues,
                                    iRound, iCalcType, iElements, dMatchType,
                                    dTotalMinLimit, dTotalMaxLimit,
                                    dCompMinLimit, dCompMaxLimit)


        'Check if return values are valid
        If resValues(0) = 0 Then Return ExcelError.ExcelErrorValue  'Results are invalid, function failed

        'Determine return value
        Select Case iReturnValue
            Case 0                                  'Return Array of Values
                'Structure array for Primary/Secondary/Tertiary values
                ResArrayValues(0, 0) = resValues(0)
                ResArrayValues(1, 0) = resValues(1)
                ResArrayValues(0, 1) = resValues(1)
                ResArrayValues(2, 0) = resValues(2)
                ResArrayValues(0, 2) = resValues(2)
                'Calculate Ratio Offset
                Select Case iElements
                    Case 0, 1, 3    'Do nothing Ratio is correct
                    Case 2, 4       'invert calculated ratio
                        resValues(5) = 1 / resValues(5)
                End Select

                'Determine what the returned error value should be
                ResArrayValues(3, 0) = OffsetValue(resValues(5), dRatio, iCalcType)
                ResArrayValues(0, 3) = ResArrayValues(3, 0)     'Copy value for horizontal part of array

                Return ResArrayValues
            Case 1 'Resistor (Primary)
                Return resValues(0)
            Case 2 'Resistor (Secondary)
                Return resValues(1)
            Case 3 'Resistor (Tertiary)
                Return resValues(2)
            Case 4 'Ratio Offset
                'Calculate Ratio Offset
                Select Case iElements
                    Case 0, 1, 3    'Do nothing Ratio is correct
                    Case 2, 4       'invert calculated ratio
                        resValues(5) = 1 / resValues(5)
                End Select
                Return resValues(5) - dRatio
            Case Else
                Return ExcelError.ExcelErrorValue   'Invalid return type
        End Select

    End Function

    'PRIVATE FUNCTIONS
    '-----------------

    'Returns ideal resistor values for ratio application
    Private Function RatioCalculator(dRatio As Double, eValues() As Double,
                                     iRound As Integer, iCalcType As Integer, iElements As Integer, dMatchType As Double,
                                     dTotalMinLimit As Double, dTotalMaxLimit As Double,
                                     dCompMinLimit As Double, dCompMaxLimit As Double)

        'Returns E-series values to generate a precise resistor ratio

        'dRatio = Target ratio
        'eValues() = table of E-Series values for matching
        '[iRound] = 1 = closest higher ratio, (0 = closest), -1 = closest lower ratio
        '[iCalcType] = Used to determine closest ratio value (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic
        '[iElements] = (0 = two components), three components: series pair (1 = bottom, 2 = top), parallel pair (3 = bottom, 4 = top)
        '[dMatchType] = How series/parallel matching is done: -1 = find best pair, 0 = stop at first perfect match, >0 = find pair with value closest to dMatchType
        '[dTotalMinLimit] = Minimum sum of components in network 
        '[dTotalMaxLimit] = Maximum sum of components in network 
        '[dCompMinLimit] = Minimum returned component value 
        '[dCompMaxLimit] = Maximum returned component value 

        'Returns {0, 0, 0} if anything goes wrong

        'Setup Variables that we're going to use
        Dim tmpValues(5, 1) As Double       'Holds calculated E-Series Values
        Dim rtnValues(5) As Double          'Return values
        Dim comboValues(2) As Double        'Temporary return value storage for series/parallel calculations
        Dim iOrder As Integer               'Order index counter
        Dim iOrderStart As Integer          'Order start point
        Dim iIndex As Integer               'Index counter
        Dim rangeMiddle As Double           'Middle of range network resistance
        Dim calcThev As Double              'Calculated Thevenin Resistance

        Dim dPrimaryMin As Double           'Primary component minimum value
        Dim dPrimaryMax As Double           'Primary component maximum value

        Dim dValueOrder As Double           'Order of magnitude calculated value

        'Initialize return values
        rtnValues(0) = 0                    'Primary resistor
        rtnValues(1) = 0                    'Secondary resistor
        rtnValues(2) = 0                    'Tertiary resistor
        rtnValues(3) = 10 ^ 10              'ABS Error between target ratio and calculated ratio (Initialize to 10 G)
        rtnValues(4) = 10 ^ 10              'ABS Error between MatchType value and closest resistor (Initialize to 10G)
        rtnValues(5) = 0                    'Calculated Ratio using E-Series Components

        'Initialize values
        rangeMiddle = dTotalMinLimit + ((dTotalMaxLimit - dTotalMinLimit) / 2)

        'Calculate Minimum Ideal Primary/Secondary Values
        Select Case iElements
            Case 0, 1, 3    'Calculated component on bottom
                'Primary component is on top. Do nothing
            Case 2, 4       'Calculated component on top
                'Primary component is on bottom.   
                dRatio = 1 / dRatio     'Invert Ratio
                iRound *= -1            'Invert Rounding
        End Select
        dPrimaryMin = dTotalMinLimit * (dRatio / (dRatio + 1))  'Calculate minimum primary component value
        dPrimaryMax = dTotalMaxLimit * (dRatio / (dRatio + 1))  'Calculate maximum primary component value

        'Moving forward, the ratio inversion means that all calculations can be thought of as the primary
        'component is on the top.

        'Calculate dPrimaryMin order of magnitude
        dValueOrder = Math.Floor(Math.Log10(dPrimaryMin))   'Calculate order of magnitude of dPrimaryMin
        iOrderStart = dValueOrder - 2                       'Figure out starting order for searching on E-Series table

        'MATCH ROUTINE
        '-------------
        'Loop through E-Series values and find closest pairing value using calculations and brute-force
        iOrder = iOrderStart          'Initialize iOrder
        Do
            For iIndex = 0 To (eValues.Length - 1)
                'Find the next e-Series primary value for tesr.   Note, for the first value, the routine will 
                'loop through all E-series possibilities from dPrimaryMin to dPrimaryMax ensuring total coverage

                tmpValues(0, 0) = eValues(iIndex) * (10 ^ iOrder)       'First E-Series Value
                If tmpValues(0, 0) > dPrimaryMax Then Return rtnValues  'Exit once all values are tried
                tmpValues(0, 1) = tmpValues(0, 0)                       'First E-Series Value
                tmpValues(1, 0) = tmpValues(0, 0) / dRatio              'Calculated ideal Secondary Value

                'Find the next component #2 e-series value using calculation.  
                'tmpValues(x,0) always creates a lower/equal ratio.   tmpValues(x,1) always creates a higher/equal ratio
                Select Case iElements
                    Case 0  'Simple Case with single secondary resistor (Primary on top, Secondary on Bottom)
                        tmpValues(1, 1) = simpleESERIES(tmpValues(1, 0), eValues, -1)   'Secondary Value (lower value = higher ratio)
                        tmpValues(2, 1) = 0                                             'Tertiary Value = 0
                        tmpValues(1, 0) = simpleESERIES(tmpValues(1, 0), eValues, +1)   'Secondary Value (higher value = lower ratio)
                        tmpValues(2, 0) = 0                                             'Tertiary Value = 0

                        'Calculate E-Series Generated Ratio
                        tmpValues(5, 0) = tmpValues(0, 0) / tmpValues(1, 0)
                        tmpValues(5, 1) = tmpValues(0, 1) / tmpValues(1, 1)

                    Case 1, 2 'Series resistor case 
                        'Secondary Value (lower value = higher ratio)
                        comboValues = ComboSeriesResistor(tmpValues(1, 0), eValues, -1, iCalcType, -1, dCompMinLimit, dCompMaxLimit)
                        tmpValues(1, 1) = comboValues(0)    'Store secondary value
                        tmpValues(2, 1) = comboValues(1)    'Store tertiary value
                        'Secondary Value (higher value = lower ratio)
                        comboValues = ComboSeriesResistor(tmpValues(1, 0), eValues, +1, iCalcType, -1, dCompMinLimit, dCompMaxLimit)
                        tmpValues(1, 0) = comboValues(0)    'Store secondary value
                        tmpValues(2, 0) = comboValues(1)    'Store tertiary value
                        'Calculate E-Series Generated Ratio
                        tmpValues(5, 0) = tmpValues(0, 0) / (tmpValues(1, 0) + tmpValues(2, 0))
                        tmpValues(5, 1) = tmpValues(0, 1) / (tmpValues(1, 1) + tmpValues(2, 1))

                    Case 3, 4 'Parallel resistor case 
                        'Secondary Value (lower value = higher ratio)
                        comboValues = ComboParallelResistor(tmpValues(1, 0), eValues, -1, iCalcType, -1, dCompMinLimit, dCompMaxLimit)
                        tmpValues(1, 1) = comboValues(0)    'Store secondary value
                        tmpValues(2, 1) = comboValues(1)    'Store tertiary value
                        'Secondary Value (higher value = lower ratio)
                        comboValues = ComboParallelResistor(tmpValues(1, 0), eValues, +1, iCalcType, -1, dCompMinLimit, dCompMaxLimit)
                        tmpValues(1, 0) = comboValues(0)    'Store secondary value
                        tmpValues(2, 0) = comboValues(1)    'Store tertiary value
                        'Calculate E-Series Generated Ratio
                        tmpValues(5, 0) = tmpValues(0, 0) / ((tmpValues(1, 0) * tmpValues(2, 0)) / (tmpValues(1, 0) + tmpValues(2, 0)))
                        tmpValues(5, 1) = tmpValues(0, 1) / ((tmpValues(1, 1) * tmpValues(2, 1)) / (tmpValues(1, 1) + tmpValues(2, 1)))

                End Select

                'Calculate Offsets
                tmpValues(3, 0) = OffsetValue(tmpValues(5, 0), dRatio, iCalcType)
                tmpValues(3, 1) = OffsetValue(tmpValues(5, 1), dRatio, iCalcType)

                'See if either one is a better match
                If tmpValues(3, 0) > rtnValues(3) And tmpValues(3, 1) > rtnValues(3) Then Continue For   'Both are bad.   Try next set

                'One of the values is equal to or better than our existing return values (MUST check both)
                For CheckCounter = 0 To 1
                    'Eliminate Bad Values
                    If tmpValues(3, CheckCounter) > rtnValues(3) Then Continue For                                                      'No match, check next set of values
                    If (tmpValues(0, CheckCounter) < dCompMinLimit) Or (tmpValues(0, CheckCounter) > dCompMaxLimit) Then Continue For   'Out of Range
                    If (tmpValues(1, CheckCounter) < dCompMinLimit) Or (tmpValues(1, CheckCounter) > dCompMaxLimit) Then Continue For   'Out of Range
                    'Check Rounding
                    Select Case iRound
                        Case -1 : If (tmpValues(5, CheckCounter) > dRatio) Then Continue For                                            'Doesn't meet Rounding
                        Case +1 : If (tmpValues(5, CheckCounter) < dRatio) Then Continue For                                            'Doesn't meet Rounding
                    End Select

                    'Check Tertiary component in Range (if required), and calculated Thevinin Resistance
                    Select Case iElements
                        Case 0      'Simple two-element case
                            calcThev = tmpValues(0, CheckCounter) + tmpValues(1, CheckCounter)
                        Case 1, 2   'Series three-element case
                            If (tmpValues(2, CheckCounter) < dCompMinLimit) Or (tmpValues(2, CheckCounter) > dCompMaxLimit) Then Continue For   'Out of Range
                            calcThev = tmpValues(0, CheckCounter) + tmpValues(1, CheckCounter) + tmpValues(2, CheckCounter)
                        Case 3, 4   'Parallel three-element case
                            If (tmpValues(2, CheckCounter) < dCompMinLimit) Or (tmpValues(2, CheckCounter) > dCompMaxLimit) Then Continue For   'Out of Range
                            calcThev = tmpValues(0, CheckCounter) + ((tmpValues(1, CheckCounter) * tmpValues(2, CheckCounter)) /
                                                                     (tmpValues(1, CheckCounter) + tmpValues(2, CheckCounter)))
                    End Select

                    If (calcThev < dTotalMinLimit) Or (calcThev > dTotalMaxLimit) Then Continue For                                     'Thevinen Resistance Out of Range

                    'Values are equal/better, within limits, and meet rounding criteria.   Need to Calculate Secondary Parameters
                    Select Case dMatchType
                        Case -1, 0  'Prefer Thevenin resistance as close to the middle of the range as possible
                            tmpValues(4, CheckCounter) = Math.Abs(rangeMiddle - calcThev)
                        Case > 1    'Prefer Thevenin resistance as close as possible to dMatchType
                            tmpValues(4, CheckCounter) = Math.Abs(dMatchType - calcThev)
                    End Select

                    'Equal/Better Match Found
                    If tmpValues(3, CheckCounter) < rtnValues(3) Then       'Better match found.   Use this always
                        rtnValues(0) = tmpValues(0, CheckCounter)
                        rtnValues(1) = tmpValues(1, CheckCounter)
                        rtnValues(2) = tmpValues(2, CheckCounter)
                        rtnValues(3) = tmpValues(3, CheckCounter)
                        rtnValues(4) = tmpValues(4, CheckCounter)
                        rtnValues(5) = tmpValues(5, CheckCounter)

                    Else                                                    'Equal match found
                        'Check Secondary Parameters
                        If tmpValues(4, CheckCounter) < rtnValues(4) Then   'Better match based on secondary parameters found.
                            rtnValues(0) = tmpValues(0, CheckCounter)
                            rtnValues(1) = tmpValues(1, CheckCounter)
                            rtnValues(2) = tmpValues(2, CheckCounter)
                            rtnValues(3) = tmpValues(3, CheckCounter)
                            rtnValues(4) = tmpValues(4, CheckCounter)
                            rtnValues(5) = tmpValues(5, CheckCounter)
                        End If
                    End If
                Next CheckCounter

                'Check for Perfect Match / dMatchType = 0 (quick match, take the first perfect match)
                If (rtnValues(3) = 0) And (dMatchType = 0) Then Return rtnValues

            Next iIndex
            iOrder += 1 'Increment
        Loop

        'Return values
        Return rtnValues
    End Function

    'Returns closest E-Series resistor value
    Private Function simpleESERIES(dValue As Double, eValues() As Double, iRound As Integer, Optional iCalcType As Integer = 0) As Double

        'Function to return closest standard resistor value
        'dValue = target value (must be greater than 0)
        'eValues() = table of E-Series values for matching
        'iRound = Rounding option 1 = next highest, 0 = closest (algebraic), -1 = next lowest
        'iCalcType:  0 returns closest algebraic value, 1 = percent, 2 = percent difference, 3 = logarithmic

        'Function return 0 (zero) if anything goes wrong

        'Setup Variables that we're going to use
        Dim dValueOrder As Double       'Order of magnitude calculated from incoming dValue
        Dim dValueShifted As Double     'dValue shifted by an integer order of magnitude to match evalues table

        Dim iIndex As Integer           'Index counter used for linear/binary search
        Dim iIndexJump As Integer       'Index counter jump value used for binary search

        'Shift dValue to ensure 3 digits before decimal point. 
        dValueOrder = Math.Floor(Math.Log10(dValue))        'Calculate order of magnitude
        dValueShifted = dValue * (10 ^ (2 - dValueOrder))   'Shift decimal point to match eSeries tables

        If dValueShifted < 100 Then                         'Capture/correct rounding case for 99.99999999999
            dValueOrder -= 1
            dValueShifted = dValue * (10 ^ (2 - dValueOrder))
        End If

        'At this point OrderValue is magnitude corrected version of dValue
        'Find E-series index of closest lower value
        iIndex = 0                              'Initialize Index Pointer
        iIndexJump = (eValues.Length - 1) / 2   'Initialize Index Jump Counter

        'Use Binary Search to find closest lower value in chosen array
        Do Until dValueShifted >= eValues(iIndex) And dValueShifted < eValues(iIndex + 1)
            'No match found
            iIndex += iIndexJump                        'Setup next Index search location
            'Setup the next Index Jump amount
            If dValueShifted >= eValues(iIndex) Then
                iIndexJump = Math.Abs(iIndexJump / 2)   'iIndexJump is positive
            Else
                iIndexJump = -Math.Abs(iIndexJump / 2)  'iIndexJump is negative
            End If
        Loop

        'Check if exact match
        If dValueShifted = eValues(iIndex) Then
            'Exact match found
            Return (eValues(iIndex) * (10 ^ (dValueOrder - 2)))
        Else
            'No exact match found, figure out rounding
            Select Case iRound
                Case 0                          'Return Closest Value based in iCalcType
                    If OffsetValue(eValues(iIndex), dValueShifted, iCalcType) < OffsetValue(eValues(iIndex + 1), dValueShifted, iCalcType) Then
                        'Lower value is closer
                        Return (eValues(iIndex) * (10 ^ (dValueOrder - 2)))
                    Else
                        'Higher value is closer or exactly the same distance from lower value
                        Return (eValues(iIndex + 1) * (10 ^ (dValueOrder - 2)))
                    End If
                Case -1                         'Return Next Lower Value
                    Return (eValues(iIndex) * (10 ^ (dValueOrder - 2)))
                Case +1                         'Return Next Higher Value
                    Return (eValues(iIndex + 1) * (10 ^ (dValueOrder - 2)))
                Case Else                       'iRound is out of range and we should return an error
                    Return 0
            End Select
        End If

    End Function

    'Returns closest parallel combination of two resistors to target value
    Private Function ComboParallelResistor(dTarget As Double, eValues() As Double,
                                           iRound As Integer, iCalcType As Integer, dMatchType As Double,
                                           dMinLimit As Double, dMaxLimit As Double) As Double()
        'Returns two parallel values for target resistor

        'dTarget = target value (must be greater than 0)
        'eValues() = table of E-Series values for matching
        'iRound = Rounding option 1 = next highest, (0 = closest), -1 = next lowest
        'iCalcType = Used to determine closest value (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic
        'dMatchType = Details how the matching pair should be found.
        'dMxxLimit (optional): Limit returned calculated component values

        'Returns {0, 0, 0} if anything goes wrong

        'The function starts at 1 or the minimum limit order and tries every single 
        'e-series combination, continuously recording if the tried combination yields
        'a more accurate result than the top contender. 

        'Initialize Variables
        Dim tmpValues(3, 1) As Double       'Holds calculated E-Series Values
        Dim rtnValues(3) As Double          'Return values
        Dim iOrder As Integer               'Order index counter
        Dim iOrderEnd As Integer          'Order end point
        Dim iIndex As Integer               'Index counter
        Dim CheckCounter As Integer

        Dim dValueOrder As Double           'Order of magnitude calculated value

        'Initialize return values
        rtnValues(0) = 0                    'First found resistor
        rtnValues(1) = 0                    'Second found resistor
        rtnValues(2) = 10 ^ 10              'Error between target and combo of two resistors (Initialize to 10 GOhm)
        rtnValues(3) = 10 ^ 10              'Error between MatchType value and closest resistor (Initialize to 10G)

        'Calculate dMaxLimit order of magnitude
        dValueOrder = Math.Floor(Math.Log10(dMaxLimit))   'Calculate order of magnitude on dMinLimit
        iOrderEnd = dValueOrder - 2                       'Figure out starting order for searching on E-Series table

        'MATCH ROUTINE
        '-------------
        'Loop through E-Series values and find closest pairing value using calculations or brute-force
        iOrder = iOrderEnd          'Initialize iOrder
        Do
            For iIndex = (eValues.Length - 1) To 0 Step -1  'Count backwards
                'Find the next component #1 e-Series value.   Not, for the first value, the routine will 

                'loop through all E-series possibilities from dMinLimit to dTarget ensuring total coverage
                tmpValues(0, 0) = eValues(iIndex) * (10 ^ iOrder)   'First E-Series Value
                If tmpValues(0, 0) <= dTarget Then Return rtnValues 'Exit once all values are tried
                tmpValues(0, 1) = tmpValues(0, 0)                   'First E-Series Value

                'Calculate Secondary values
                tmpValues(1, 0) = (tmpValues(0, 0) * dTarget) / (tmpValues(0, 0) - dTarget)     'Calculated ideal second Value
                'tmpValues(1, 0) = dTarget - tmpValues(0, 0)                     'Calculated ideal Secondary Value
                tmpValues(1, 1) = simpleESERIES(tmpValues(1, 0), eValues, +1)   'Closest Secondary E-Series Value (higher)
                tmpValues(1, 0) = simpleESERIES(tmpValues(1, 0), eValues, -1)   'Closest Secondary E-Series Value (lower)
                'Calculate both offset values of combo from target
                tmpValues(2, 0) = OffsetValue((tmpValues(0, 0) * tmpValues(1, 0)) / (tmpValues(0, 0) + tmpValues(1, 0)),
                                              dTarget, iCalcType)                                            'Lower Resistor Offset
                tmpValues(2, 1) = OffsetValue((tmpValues(0, 1) * tmpValues(1, 1)) / (tmpValues(0, 1) + tmpValues(1, 1)),
                                              dTarget, iCalcType)                                             'Upper Resistor Offset


                'See if either one is a better match
                If tmpValues(2, 0) > rtnValues(2) And tmpValues(2, 1) > rtnValues(2) Then Continue For   'Both are bad.   Try next set

                'One of the values is equal to or better than our existing return values
                For CheckCounter = 0 To 1
                    'Eliminate Bad Matches
                    If tmpValues(2, CheckCounter) > rtnValues(2) Then Continue For     'No match, check next set of values
                    If (tmpValues(0, CheckCounter) < dMinLimit) Or (tmpValues(0, CheckCounter) > dMaxLimit) Or
                       (tmpValues(1, CheckCounter) < dMinLimit) Or (tmpValues(1, CheckCounter) > dMaxLimit) Then Continue For       'Out of Range

                    Select Case iRound
                        Case -1
                            If ((tmpValues(0, CheckCounter) * tmpValues(1, CheckCounter)) /
                                (tmpValues(0, CheckCounter) + tmpValues(1, CheckCounter)) > dTarget) Then Continue For              'Doesn't meet Rounding
                        Case +1
                            If ((tmpValues(0, CheckCounter) * tmpValues(1, CheckCounter)) /
                                (tmpValues(0, CheckCounter) + tmpValues(1, CheckCounter)) < dTarget) Then Continue For              'Doesn't meet Rounding
                    End Select

                    'Values are equal/better, within limits, and meet rounding criteria.   Need to Calculate Secondary Parameters
                    Select Case dMatchType
                        Case -1, 0  'Prefer values that are as close to equal as possible
                            tmpValues(3, CheckCounter) = Math.Abs(tmpValues(0, CheckCounter) - tmpValues(1, CheckCounter))
                        Case > 1    'Prefer one value be as close as possible to dMatchType
                            tmpValues(3, CheckCounter) = Math.Min(
                                                  OffsetValue(tmpValues(0, CheckCounter), dMatchType, iCalcType),
                                                  OffsetValue(tmpValues(1, CheckCounter), dMatchType, iCalcType))
                    End Select

                    'Equal/Better Match Found
                    If tmpValues(2, CheckCounter) < rtnValues(2) Then       'Better match found.   Use this always
                        rtnValues(0) = tmpValues(0, CheckCounter)
                        rtnValues(1) = tmpValues(1, CheckCounter)
                        rtnValues(2) = tmpValues(2, CheckCounter)
                        rtnValues(3) = tmpValues(3, CheckCounter)
                    Else                                                        'Equal match found
                        'Check Secondary Parameters
                        If tmpValues(3, CheckCounter) < rtnValues(3) Then   'Better match based on secondary parameters found.
                            rtnValues(0) = tmpValues(0, CheckCounter)
                            rtnValues(1) = tmpValues(1, CheckCounter)
                            rtnValues(2) = tmpValues(2, CheckCounter)
                            rtnValues(3) = tmpValues(3, CheckCounter)
                        End If
                    End If
                Next CheckCounter

                'Check for Perfect Match and quick matching
                If (rtnValues(2) = 0) And (dMatchType = 0) Then Return rtnValues

            Next iIndex
            iOrder -= 1 'Decrement
        Loop

        'Return values
        Return rtnValues

    End Function

    'Returns closest series combination of two resistors to target value
    Private Function ComboSeriesResistor(dTarget As Double, eValues() As Double,
                                         iRound As Integer, iCalcType As Integer, dMatchType As Double,
                                         dMinLimit As Double, dMaxLimit As Double) As Double()
        'Returns two parallel values for target resistor

        'dTarget = target value (must be greater than 0)
        'eValues() = table of E-Series values for matching
        'iRound = Rounding option 1 = next highest, (0 = closest), -1 = next lowest
        'iCalcType = Used to determine closest value (0 = algebraic), 1 = percent, 2 = percent difference, 3 = logarithmic
        'dMatchType = Details how the matching pair should be found.
        'dMxxLimit (optional): Limit returned calculated component values

        'Returns {0, 0, 0} if anything goes wrong

        'The function starts at 1 or the minimum limit order and tries every single 
        'e-series combination, continuously recording if the tried combination yields
        'a more accurate result than the top contender. 

        'Initialize Variables
        Dim tmpValues(3, 1) As Double       'Holds calculated E-Series Values
        Dim rtnValues(3) As Double          'Return values
        Dim iOrder As Integer               'Order index counter
        Dim iOrderStart As Integer          'Order end point
        Dim iIndex As Integer               'Index counter
        Dim CheckCounter As Integer

        Dim dValueOrder As Double           'Order of magnitude calculated value


        'Initialize return values
        rtnValues(0) = 0                    'First found resistor
        rtnValues(1) = 0                    'Second found resistor
        rtnValues(2) = 10 ^ 10              'Error between target and combo of two resistors (Initialize to 10 GOhm)
        rtnValues(3) = 10 ^ 10              'Error between MatchType value and closest resistor (Initialize to 10G)

        'Calculate dMinLimit order of magnitude
        dValueOrder = Math.Floor(Math.Log10(dMinLimit))     'Calculate order of magnitude on dMinLimit
        iOrderStart = dValueOrder - 2                       'Figure out starting order for searching on E-Series table

        'MATCH ROUTINE
        '-------------
        'Loop through E-Series values and find closest pairing value using calculations or brute-force
        iOrder = iOrderStart          'Initialize iOrder
        Do
            For iIndex = 0 To (eValues.Length - 1)
                'Find the next component #1 e-Series value.   Not, for the first value, the routine will 

                'loop through all E-series possibilities from dMinLimit to dTarget ensuring total coverage
                tmpValues(0, 0) = eValues(iIndex) * (10 ^ iOrder)   'First E-Series Value
                If tmpValues(0, 0) >= dTarget Then Return rtnValues 'Exit once all values are tried
                tmpValues(0, 1) = tmpValues(0, 0)                   'First E-Series Value

                'Calculate Secondary values
                tmpValues(1, 0) = dTarget - tmpValues(0, 0)                     'Calculated ideal Secondary Value
                tmpValues(1, 1) = simpleESERIES(tmpValues(1, 0), eValues, +1)   'Closest Secondary E-Series Value (higher)
                tmpValues(1, 0) = simpleESERIES(tmpValues(1, 0), eValues, -1)   'Closest Secondary E-Series Value (lower)
                'Calculate both offset values of combo from target
                tmpValues(2, 0) = OffsetValue((tmpValues(0, 0) + tmpValues(1, 0)), dTarget, iCalcType)   'Lower Resistor Offset
                tmpValues(2, 1) = OffsetValue((tmpValues(0, 1) + tmpValues(1, 1)), dTarget, iCalcType)   'Upper Resistor Offset

                'See if either one is a better match
                If tmpValues(2, 0) > rtnValues(2) And tmpValues(2, 1) > rtnValues(2) Then Continue For   'Both are bad.   Try next set

                'One of the values is equal to or better than our existing return values
                For CheckCounter = 0 To 1
                    'Eliminate Bad Values
                    If tmpValues(2, CheckCounter) > rtnValues(2) Then Continue For      'No match, check next set of values
                    If (tmpValues(0, CheckCounter) < dMinLimit) Or (tmpValues(0, CheckCounter) > dMaxLimit) Or
                       (tmpValues(1, CheckCounter) < dMinLimit) Or (tmpValues(1, CheckCounter) > dMaxLimit) Then Continue For   'Out of Range
                    Select Case iRound
                        Case -1
                            If (tmpValues(0, CheckCounter) + tmpValues(1, CheckCounter) > dTarget) Then Continue For            'Doesn't meet Rounding
                        Case +1
                            If (tmpValues(0, CheckCounter) + tmpValues(1, CheckCounter) < dTarget) Then Continue For            'Doesn't meet Rounding
                    End Select

                    'Values are equal/better, within limits, and meet rounding criteria.   Need to Calculate Secondary Parameters
                    Select Case dMatchType
                        Case -1, 0  'Prefer values that are as close to equal as possible
                            tmpValues(3, CheckCounter) = Math.Abs(tmpValues(0, CheckCounter) - tmpValues(1, CheckCounter))
                        Case > 1    'Prefer one value be as close as possible to dMatchType
                            tmpValues(3, CheckCounter) = Math.Min(
                                                  OffsetValue(tmpValues(0, CheckCounter), dMatchType, iCalcType),
                                                  OffsetValue(tmpValues(1, CheckCounter), dMatchType, iCalcType))
                    End Select

                    'Equal/Better Match Found
                    If tmpValues(2, CheckCounter) < rtnValues(2) Then       'Better match found.   Use this always
                        rtnValues(0) = tmpValues(0, CheckCounter)
                        rtnValues(1) = tmpValues(1, CheckCounter)
                        rtnValues(2) = tmpValues(2, CheckCounter)
                        rtnValues(3) = tmpValues(3, CheckCounter)
                    Else                                                        'Equal match found
                        'Check Secondary Parameters
                        If tmpValues(3, CheckCounter) < rtnValues(3) Then   'Better match based on secondary parameters found.
                            rtnValues(0) = tmpValues(0, CheckCounter)
                            rtnValues(1) = tmpValues(1, CheckCounter)
                            rtnValues(2) = tmpValues(2, CheckCounter)
                            rtnValues(3) = tmpValues(3, CheckCounter)
                        End If
                    End If
                Next CheckCounter

                'Check for Perfect Match and quick matching
                If (rtnValues(2) = 0) And (dMatchType = 0) Then Return rtnValues

            Next iIndex
            iOrder += 1 'Increment
        Loop

        'Return values
        Return rtnValues

    End Function

    'Returns value of offset based on iCalcType parameter
    Private Function OffsetValue(dVA As Double, dVE As Double, Optional iCalcType As Integer = 0)
        'Returns offset of value from target
        'dVA = Actual value     (test value)
        'dVE = Expected value   (ideal calculated value)
        'iCalcType = option value to determine what kind of calculation we should use and what the function returns
        '   0 = Algebraic 
        '   1 = Percent 
        '   2 = Percent Difference  (Geometric)
        '   3 = Logarithmic

        Select Case iCalcType
            Case 0      'Calculate Algebraic Error
                OffsetValue = Math.Abs(dVA - dVE)
            Case 1      'Calculate Percent Error
                OffsetValue = Math.Abs((dVA - dVE) / dVE)
            Case 2      'Calculate Percent Difference Error
                OffsetValue = (Math.Abs(dVA - dVE)) / ((dVA + dVE) / 2)
            Case 3      'Calculate Logarithmic Error
                OffsetValue = Math.Abs((0.5) * Math.Log10(dVA / dVE))
            Case Else   'Not valid
                Return 0
        End Select
    End Function


End Module
