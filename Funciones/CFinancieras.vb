Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Public Class CFinancieras
    Private Const Categoria = "Creasys"
#Region "Curve"
    <ExcelFunction(category:=Categoria, Description:="Calcula tasa dado un factor de descuento. (utiliza la función 'YearFraction')")> _
    Shared Function GetRateFromDf(<ExcelArgument(Description:="fecha inicio")> _
                                              ByVal startDate As Date, _
                                              <ExcelArgument(Description:="fecha fin")> _
                                              ByVal endDate As Date, _
                                              <ExcelArgument(Description:="factor de descuento")> _
                                              ByVal Df As Double, _
                                              <ExcelArgument(Description:="Lin Act/365, Comp act/360 ...0")> _
                                              ByVal typeOfRate As String) As Double
        Dim Result, Plazo As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            typeOfRate = typeOfRate.ToLower.Trim
            Select Case typeOfRate
                Case "lin act/360"
                    Plazo = DateDiff(DateInterval.Day, startDate, endDate)
                    Result = (1 / Df - 1) * 360 / Plazo
                Case "lin 30/360"
                    Plazo = CountDays(startDate, endDate, "30/360")
                    Result = (1 / Df - 1) * 360 / Plazo
                Case "lin act/365"
                    Plazo = DateDiff(DateInterval.Day, startDate, endDate)
                    Result = (1 / Df - 1) * 365 / Plazo
                Case "comp 30/360"
                    Plazo = CountDays(startDate, endDate, "30/360")
                    Result = (1 / Df) ^ (360 / Plazo) - 1
                Case "comp act/360"
                    Plazo = DateDiff(DateInterval.Day, startDate, endDate)
                    Result = (1 / Df) ^ (360 / Plazo) - 1
                Case "comp act/365"
                    Plazo = DateDiff(DateInterval.Day, startDate, endDate)
                    Result = (1 / Df) ^ (365 / Plazo) - 1
                Case Else
                    Plazo = DateDiff(DateInterval.Day, startDate, endDate)
                    Result = (1 / Df - 1) * 360 / Plazo
            End Select
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula tasa lineal dado un factor de descuento. (utiliza la función 'YearFraction')")> _
    Shared Function GetRateFromDiscountFactor(<ExcelArgument(Description:="fecha inicio")> _
                                              ByVal FechaIni As Date, _
                                              <ExcelArgument(Description:="fecha fin")> _
                                              ByVal FechaFin As Date, _
                                              <ExcelArgument(Description:="factor de descuento")> _
                                              ByVal Df As Double, _
                                              <ExcelArgument(Description:="Act/365, act/360 ó 30/360")> _
                                              ByVal Forma As String) As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Forma = Forma.ToLower.Trim
        Select Case Forma
            Case "30/360"
                Result = (1 / Df - 1) * 1 / YearFraction(FechaIni, FechaFin, Forma)
            Case "act/360"
                Result = (1 / Df - 1) * 1 / YearFraction(FechaIni, FechaFin, Forma)
        End Select
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function GetRateFromDiscountFactor_v2(ByVal Plazo As Double, ByVal Df As Double, ByVal Forma As String, ByVal compound As Integer) As Double
        'Forma = Forma.ToLower.Trim
        'Select Case compound
        '    Case 1
        '        GetRateFromDiscountFactor_v2 = (1 / Df - 1) * 1 / YearFraction(0, Plazo, Forma)
        '    Case 2
        '        GetRateFromDiscountFactor_v2 = (1 / Df) ^ (1 / YearFraction(0, Plazo, Forma)) - 1
        '    Case 3
        '        GetRateFromDiscountFactor_v2 = -Math.Log(Df) / YearFraction(0, Plazo, Forma)
        '    Case Else
        '        GetRateFromDiscountFactor_v2 = (1 / Df) ^ (1 / YearFraction(0, Plazo, Forma)) - 1
        'End Select
        Return 0
    End Function
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function GetRateFromCurve(ByVal days As Double, _
                                     ByVal tenors As Object, _
                                     ByVal rates As Object, _
                                     ByVal Model As Object, _
                                     ByVal parameters() As Double, _
                                     Optional ByVal basis As Double = 365, _
                                     Optional ByVal compound As Integer = 2) As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If IsNothing(Model) Then
                Result = LinInterpol(tenors, rates, days)
            ElseIf IsNothing(parameters) Then
                Result = 0
            Else
                Result = RateFromSvensson(days, parameters, basis, compound)
            End If
        Catch ex As Exception
            Result = 0
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el factor descuento asociado a un plazo y una curva (interpola linealmente en tasa)")> _
    Shared Function GetDiscountFactorFromCurve(<ExcelArgument(Description:="Tiempo en días al que se quiere obtener el factor de descuento")> _
                                                      ByVal days As Double, _
                                                      <ExcelArgument(Description:="Tenors de la curva")> _
                                                      ByVal Tenors As Object, _
                                                      <ExcelArgument(Description:="Tasas correpondientes a cada tenor")> _
                                                      ByVal Rates As Object, _
                                                      <ExcelArgument(Description:="360, 365")> _
                                                      ByVal basis As Double, _
                                                      <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (valor defecto 2)")> _
                                                      ByVal compound As Integer) As Double
        Dim parameters() As Double = Nothing
        Dim rate As Double
        Dim Result As Double


        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            rate = LinInterpol(Tenors, Rates, days)
            Select Case compound
                Case 1
                    Result = 1 / (1 + rate * days / basis)
                Case 2
                    Result = 1 / (1 + rate) ^ (days / basis)
                Case 3
                    Result = Math.Exp(-rate * days / basis)
                Case Else
                    Result = 1 / (1 + rate) ^ (days / basis)
            End Select
        Catch ex As Exception
            Result = Nothing
        End Try
        Return Result

    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el df fwd entre dos plazos (interpola linealmente en tasa)")> _
    Shared Function GetDfFwdFromCurve(<ExcelArgument(Description:="Días iniciales")> _
                                                      ByVal days1 As Double, _
                                                      <ExcelArgument(Description:="Días finales")> _
                                                      ByVal days2 As Double, _
                                                      <ExcelArgument(Description:="Tenors de la curva")> _
                                                      ByVal Tenors As Object, _
                                                      <ExcelArgument(Description:="Tasas correpondientes a cada tenor")> _
                                                      ByVal Rates As Object, _
                                                      <ExcelArgument(Description:="360, 365")> _
                                                      ByVal basis As Double, _
                                                      <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (valor defecto 2)")> _
                                                      ByVal compound As Integer) As Double
        Dim Rate1, Rate2, aux1, aux2 As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Rate1 = LinInterpol(Tenors, Rates, days1)
            Rate2 = LinInterpol(Tenors, Rates, days2)
            Select Case compound
                Case 1
                    aux1 = 1 / (1 + Rate1 * days1 / basis)
                    aux2 = 1 / (1 + Rate2 * days2 / basis)
                    Result = aux2 / aux1
                Case 2
                    aux1 = 1 / (1 + Rate1) ^ (days1 / basis)
                    aux2 = 1 / (1 + Rate2) ^ (days2 / basis)
                    Result = aux2 / aux1
                Case 3
                    aux1 = Math.Exp(-Rate1 * days1 / basis)
                    aux2 = Math.Exp(-Rate2 * days2 / basis)
                    Result = aux2 / aux1
                Case Else
                    aux1 = 1 / (1 + Rate1) ^ (days1 / basis)
                    aux2 = 1 / (1 + Rate2) ^ (days2 / basis)
                    Result = aux2 / aux1
            End Select
        Catch ex As Exception
            Result = Nothing
        End Try

        Return Result

    End Function
    <ExcelFunction(category:=Categoria, Description:="Interpola lineal en tasa")> _
    Shared Function LinInterpol(<ExcelArgument(Description:="Plazos")> _
                                        ByVal plazos As Object, _
                                        <ExcelArgument(Description:="Tasas correspondiente a cada plazo definido anteriormente")> _
                                        ByVal tasas As Object, _
                                        <ExcelArgument(Description:="Plazo al que se interpolará la tasa")> _
                                        ByVal Plazo As Double) As Double
        Dim i As Integer
        Dim j As Integer
        Dim tasa1 As Double
        Dim tasa2 As Double
        Dim plazo1 As Double
        Dim plazo2 As Double
        Dim Result As Double
        Dim arrayTasas() As Double
        Dim arrayPlazos() As Double
        Dim IndexBase As Integer
        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try

            If tasas Is Nothing Or plazos Is Nothing Then
                Return 0
            End If

            arrayTasas = Range2Array(Of Double)(tasas)
            arrayPlazos = Range2Array(Of Double)(plazos)



            IndexBase = 0
            j = arrayTasas.GetUpperBound(0) - 1
            i = arrayPlazos.GetUpperBound(0) - 1

            If i <> j Then
                Result = 0
            ElseIf Plazo >= arrayPlazos(j) Then
                Result = arrayTasas(j)
            ElseIf Plazo <= arrayPlazos(IndexBase) Then
                Result = arrayTasas(IndexBase)
            Else
                For i = IndexBase To j
                    If Plazo > arrayPlazos(i) Then
                        plazo1 = arrayPlazos(i)
                        tasa1 = arrayTasas(i)
                        plazo2 = arrayPlazos(i + 1)
                        tasa2 = arrayTasas(i + 1)
                        'MsgBox(plazo1 & "," & tasa1 & "," & plazo2 & "," & tasa2)
                    Else
                        Exit For

                    End If
                Next i
                Result = tasa1 + (tasa2 - tasa1) / (plazo2 - plazo1) * (Plazo - plazo1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try


        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Interpola exponencial en factor de descuento")> _
    Shared Function ExpInterpol(<ExcelArgument(Description:="Plazos")> _
                                ByVal plazos() As Integer, _
                                <ExcelArgument(Description:="Factores de descuento a cada plazo definido anteriormente")> _
                                ByVal factores() As Double, _
                                <ExcelArgument(Description:="Plazo al que se interpolará el factor de descuento")> _
                                ByVal Plazo As Integer) As Double
        Dim i As Integer
        Dim j As Integer
        Dim plazo1 As Integer
        Dim plazo2 As Integer
        Dim tasa1 As Double = 0
        Dim tasa2 As Double = 0
        Dim factor1 As Double
        Dim factor2 As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            i = plazos.Length
            j = factores.Length
            If i <> j Then
                Result = 0
            ElseIf Plazo >= plazos(j) Then
                Result = factores(j)
            ElseIf Plazo <= plazos(1) Then
                Result = factores(1)
            Else
                For i = 1 To j
                    If Plazo > plazos(i) Then
                        plazo1 = plazos(i)
                        factor1 = factores(i)
                        plazo2 = plazos(i + 1)
                        factor2 = factores(i + 1)
                    End If
                Next i
                Result = factor1 * (factor2 / factor1) ^ ((Plazo - plazo1) / (plazo2 - plazo1))
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function

#End Region
#Region "Svensson"
    <ExcelFunction(category:=Categoria, Description:="Calcula factor descuento en función del tiempo y los parametros de N-S-S")> _
    Shared Function DiscountFactorFromSvensson(<ExcelArgument(Description:="Tiempo al cual se calcula factor descuento")> _
                                                      ByVal t As Double, _
                                                      <ExcelArgument(Description:="Los parametros son; beta0, beta1, beta2, t1 y t2")> _
                                                      ByVal parameters() As Double) As Double
        Dim beta0 As Double
        Dim beta1 As Double
        Dim beta2 As Double
        Dim beta3 As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        beta0 = parameters(1)
        beta1 = parameters(2)
        beta2 = parameters(3)
        beta3 = parameters(4)
        t1 = parameters(5)
        t2 = parameters(6)
        Try
            If t = 0 Then
                Result = beta0 + beta1
            Else
                Result = Math.Exp(-t * ((beta0 + beta1 * (1 - Math.Exp(-t / t1)) * t1 / t + beta2 * ((1 - Math.Exp(-t / t1)) * t1 / t - Math.Exp(-t / t1)) + beta3 * ((1 - Math.Exp(-t / t2)) * t2 / t - Math.Exp(-t / t2)))))
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula tasa cero asociada al factor descuento N-S-S dada una base y composición")> _
    Shared Function RateFromSvensson(<ExcelArgument(Description:="Tiempo al cual se calcula factor descuento")> _
                                     ByVal days As Double, _
                                     <ExcelArgument(Description:="Los parametros son; beta0, beta1, beta2, t1 y t2")> _
                                     ByVal parameters() As Double, _
                                     <ExcelArgument(Description:="360, 365")> _
                                     ByVal basis As Double, _
                                     <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                     ByVal compound As Integer) As Double
        Dim Df As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Df = DiscountFactorFromSvensson(days / 365, parameters)
            Select Case compound
                Case 1
                    Result = (1 / Df - 1) * basis / days
                Case 2
                    Result = (1 / Df) ^ (basis / days) - 1
                Case 3
                    Result = -basis / days * Math.Log(Df)
                Case Else
                    Result = (1 / Df) ^ (basis / days) - 1
            End Select
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
#End Region
#Region "Swaps"
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de flujos de caja puestos en distintas fechas (utiliza función GetDiscoundFactorFroma Curve)")> _
    Shared Function pvCashFlowsFromCurve(<ExcelArgument(Description:="Fechas de flujos cajas")> _
                                         ByVal tenorsCashFlows() As Double, _
                                         <ExcelArgument(Description:="flujos de cajas")> _
                                         ByVal cashFlows() As Double, _
                                         <ExcelArgument(Description:="tenors de curva de descuento")> _
                                         ByVal tenors() As Double, _
                                         <ExcelArgument(Description:="tasas a los correspondientes tenors")> _
                                         ByVal rates() As Double, _
                                         <ExcelArgument(Description:="Base de tasa de referencia (360, 365)")> _
                                         ByVal basis As Integer, _
                                         <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (valor defecto 2)")> _
                                         ByVal compound As Integer) As Double
        Dim Result As Double = 0
        Dim i As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            For i = 0 To tenorsCashFlows.GetUpperBound(0)
                Result += cashFlows(i) * GetDiscountFactorFromCurve(tenorsCashFlows(i), tenors, rates, basis, compound)
            Next i
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el interes de un cupon pata camara proyectando la parte no fijada")> _
    Shared Function couponCamara(<ExcelArgument(Description:="Fecha de valoracion")> _
                                         ByVal valueDate As Date, _
                                         <ExcelArgument(Description:="Fecha inicio cupon")> _
                                         ByVal startDate As Date, _
                                         <ExcelArgument(Description:="Fecha fin de cupon")> _
                                         ByVal endDate As Date, _
                                         <ExcelArgument(Description:="Fijacion hasta value date")> _
                                         ByVal fixing As Double, _
                                         <ExcelArgument(Description:="Spread aditivo")> _
                                         ByVal spread As Double, _
                                         <ExcelArgument(Description:="Tipo cupon")> _
                                         ByVal typeCoupon As String, _
                                         <ExcelArgument(Description:="Nocional")> _
                                         ByVal notional As Double, _
                                         <ExcelArgument(Description:="Plazos de la curva")> _
                                         ByVal tenors As Object, _
                                         <ExcelArgument(Description:="Tasas de la curva")> _
                                         ByVal rates As Object, _
                                         <ExcelArgument(Description:="Base de las tasas")> _
                                         ByVal basis As Double, _
                                         <ExcelArgument(Description:="Tipo de composición 1:lineal 2:compuesto")> _
                                         ByVal compound As Double) As Double

        Dim Result As Double = Nothing
        Dim aux, df1, df2 As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If valueDate > endDate Then
                Result = 0
            ElseIf valueDate <= startDate Then
                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, startDate), tenors, rates, basis, compound)
                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, endDate), tenors, rates, basis, compound)
                Result = interest(startDate, endDate, spread, typeCoupon, notional) + notional * (df1 / df2 - 1)
            Else
                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, endDate), tenors, rates, basis, compound)
                aux = interest(startDate, valueDate, fixing, typeCoupon, notional)
                Result = interest(startDate, endDate, spread, typeCoupon, notional) + (notional + aux) / df2 - notional
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el interes de un cupon pata flotante proyectando si no ha fijado")> _
        Shared Function couponFloat(<ExcelArgument(Description:="Fecha de valoracion")> _
                                             ByVal valueDate As Date, _
                                             <ExcelArgument(Description:="Fecha inicio cupon")> _
                                             ByVal startDate As Date, _
                                             <ExcelArgument(Description:="Fecha fin de cupon")> _
                                             ByVal endDate As Date, _
                                             <ExcelArgument(Description:="Tenor indice flotante")> _
                                             ByVal floatIndexTenor As String, _
                                             <ExcelArgument(Description:="Lag de fijacion del índice flotante")> _
                                             ByVal fixingLag As Integer, _
                                             <ExcelArgument(Description:="Shift de partida del índice flotante")> _
                                             ByVal floatIndexShift As Integer, _
                                             <ExcelArgument(Description:="Fijacion")> _
                                             ByVal fixing As Double, _
                                             <ExcelArgument(Description:="Spread aditivo")> _
                                             ByVal spread As Double, _
                                             <ExcelArgument(Description:="Tipo cupon")> _
                                             ByVal typeCoupon As String, _
                                             <ExcelArgument(Description:="Nocional")> _
                                             ByVal notional As Double, _
                                             <ExcelArgument(Description:="Plazos de la curva")> _
                                             ByVal tenors As Object, _
                                             <ExcelArgument(Description:="Tasas de la curva")> _
                                             ByVal rates As Object, _
                                             <ExcelArgument(Description:="Base de las tasas")> _
                                             ByVal basis As Double, _
                                             <ExcelArgument(Description:="Tipo de composición 1:lineal 2:compuesto")> _
                                             ByVal compound As Double) As Double

        Dim Result As Double = Nothing
        Dim aux, df1, df2 As Double
        Dim fixingDate, floatStartDate, tenorDate As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            fixingDate = lag(startDate, fixingLag)
            If valueDate > endDate Then
                Result = 0
            ElseIf valueDate < fixingDate Then
                floatStartDate = shift(fixingDate, floatIndexShift)
                tenorDate = BussDay(AddMonths(floatStartDate, Ten(floatIndexTenor)))
                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, floatStartDate), tenors, rates, basis, compound)
                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, tenorDate), tenors, rates, basis, compound)
                aux = GetRateFromDf(floatStartDate, tenorDate, df2 / df1, typeCoupon)
                Result = interest(startDate, endDate, aux + spread, typeCoupon, notional)
            Else
                Result = interest(startDate, endDate, fixing + spread, typeCoupon, notional)
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el valor presente de una pata cualquiera")> _
 Shared Function pvAnyLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                         ByVal valueDate As Date, _
                                         <ExcelArgument(Description:="Fecha ini")> _
                                         ByVal startDate As Date, _
                                         <ExcelArgument(Description:="Fecha fin")> _
                                         ByVal endDate As Date, _
                                         <ExcelArgument(Description:="Nocional")> _
                                         ByVal notional As Double, _
                                         <ExcelArgument(Description:="Fijo o flotante")> _
                                         ByVal fixFloat As String, _
                                         <ExcelArgument(Description:="Tenors curva proy.")> _
                                         ByVal projTenors As Object, _
                                         <ExcelArgument(Description:="Tasas curva proy.")> _
                                         ByVal projRates As Object, _
                                         <ExcelArgument(Description:="Tenors curva desc.")> _
                                         ByVal discTenors As Object, _
                                         <ExcelArgument(Description:="Tasas curva desc.")> _
                                         ByVal discRates As Object, _
                                         <ExcelArgument(Description:="Periodicidad")> _
                                         ByVal periodicity As Integer, _
                                         <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                         ByVal accAdjustment As Integer, _
                                         <ExcelArgument(Description:="Ajuste de fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> _
                                         ByVal pmtAdjustment As Integer, _
                                         <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                         ByVal typeStubPeriod As Integer, _
                                         <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1=un día hábil; 2=dos días hábiles ")> _
                                         ByVal fixingLag As Integer, _
                                         <ExcelArgument(Description:="Periocidad de fijación: n, se fija cada n periodos")> _
                                         ByVal fixingRatio As Integer, _
                                         <ExcelArgument(Name:="fixingStubPer", Description:="Periodo fijación única: 0:al final; 1:al principio")> _
                                         ByVal fixingStubPeriod As Integer, _
                                         <ExcelArgument(Description:="1:bullet, 2:constante, 3:frances, else:bullet")> _
                                         ByVal typeAmortization As Integer, _
                                         <ExcelArgument(Description:="Tasa para frances Com 30/360")> _
                                         ByVal amortizeRate As Double, _
                                         <ExcelArgument(Description:="No implementado")> _
                                         ByVal typeAmortizeRate As String, _
                                         <ExcelArgument(Description:="Amortizaciones y nocional vigente")> _
                                         ByVal currentNotionalAndAmort As Object, _
                                         <ExcelArgument(Description:="Tipo de capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360) (por defecto asume comp)")> _
                                         ByVal typeOfSpreadOrRate As String, _
                                         <ExcelArgument(Description:="Tenor tasa ref")> _
                                         ByVal floatIndexTenor As String, _
                                         <ExcelArgument(Description:="Ajuste tasa ref")> _
                                         ByVal floatIndexShift As Integer, _
                                         <ExcelArgument(Description:="1: es cámara, Else: no")> _
                                         ByVal camaraFlag As Integer, _
                                         <ExcelArgument(Description:="1: paga amortizaciones, Else: no paga")> _
                                         ByVal exchange As Integer, _
                                         <ExcelArgument(Description:="Ultimo fixing")> _
                                         ByVal lastFixing As Double, _
                                         <ExcelArgument(Description:="Tasa o spread")> _
                                         ByVal spreadOrRate As Double) As Double

        'typeOfRate can be Lin Act/360, Lin 30/360, Com Act/365
        Dim Result As Double
        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If


        If fixFloat = "Fixed" Then
            If typeAmortization = 4 Then
                Result = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, notional, discTenors, discRates, 365, 2, exchange)
            Else
                Result = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, discTenors, discRates, 365, 2, exchange)
            End If
        Else
            If typeAmortization = 4 Then
                Result = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, currentNotionalAndAmort, lastFixing, spreadOrRate, typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, notional, projTenors, projRates, 365, 2, camaraFlag, exchange, discTenors, discRates, 365, 2)
            Else
                Result = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, notional, projTenors, projRates, 365, 2, camaraFlag, exchange, discTenors, discRates, 365, 2)
            End If
        End If
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el valor presente de una pata fija amortizable")> _
    Shared Function pvFixedAmortSwapLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                            ByVal valueDate As Date, _
                                            <ExcelArgument(Description:="Fecha ini")> _
                                            ByVal startDate As Date, _
                                            <ExcelArgument(Description:="Fecha fin")> _
                                            ByVal endDate As Date, _
                                            <ExcelArgument(Description:="Periocidad")> _
                                            ByVal periodicity As Integer, _
                                            <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                            ByVal accAdjustment As Integer, _
                                            <ExcelArgument(Description:="Ajuste de fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> _
                                            ByVal pmtAdjustment As Integer, _
                                            <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                            ByVal typeStubPeriod As Integer, _
                                            <ExcelArgument(Description:="1:bullet, 2:constante, 3:frances, else:bullet")> _
                                            ByVal typeAmortization As Integer, _
                                            <ExcelArgument(Description:="Tasa para frances Com 30/360")> _
                                            ByVal amortizeRate As Double, _
                                            <ExcelArgument(Description:="No implementado")> _
                                            ByVal typeAmortizeRate As String, _
                                            <ExcelArgument(Description:="Tasa fija del Swap")> _
                                            ByVal rate As Double, _
                                            <ExcelArgument(Description:="Tipo de capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360) (por defecto asume comp)")> _
                                            ByVal typeOfRate As String, _
                                            <ExcelArgument(Description:="Nocional")> _
                                            ByVal notional As Double, _
                                            <ExcelArgument(Description:="Tenors de la curva")> _
                                            ByVal tenors As Object, _
                                            <ExcelArgument(Description:="Tasas correpondientes a cada tenor")> _
                                            ByVal rates As Object, _
                                            <ExcelArgument(Description:="360, 365")> _
                                            ByVal basis As Double, _
                                            <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3")> _
                                            ByVal compound As Integer, _
                                            <ExcelArgument(Description:="Al ingresar un valor se asume pago de intereses + capital en último periódo.De lo contrario, sólo pago de interes")> _
                                            ByVal exchange As Object) As Double
        'typeOfRate can be Lin Act/360, Lin 30/360, Com Act/365
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim aux, currentNotional As Double
        Dim auxDate As Date
        Dim i As Integer
        Dim Result As Double
        Dim indexBase As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                Result = 0
            Else
                schedule = calendarAmortize(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, typeAmortizeRate)
                numberOfPeriods = UBound(schedule, 1)

                currentNotional = 0
                i = numberOfPeriods
                auxDate = schedule(i, 2)
                aux = 0
                While valueDate < auxDate
                    aux = aux + (interest(schedule(i, 0), schedule(i, 1), rate, typeOfRate, notional * schedule(i, 4)) + exchange * notional * schedule(i, 3)) _
                    * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors, rates, basis, compound)
                    i = i - 1
                    If i < 0 Then Exit While
                    auxDate = schedule(i, 2)
                End While
            End If
            Result = aux
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el valor presente de una pata fija amortizable customizada")> _
    Shared Function pvFixedCustomAmortSwapLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                            ByVal valueDate As Date, _
                                            <ExcelArgument(Description:="Fecha ini")> _
                                            ByVal startDate As Date, _
                                            <ExcelArgument(Description:="Fecha fin")> _
                                            ByVal endDate As Date, _
                                            <ExcelArgument(Description:="Periocidad")> _
                                            ByVal periodicity As Integer, _
                                            <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                            ByVal accAdjustment As Integer, _
                                            <ExcelArgument(Description:="Ajuste de fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> _
                                            ByVal pmtAdjustment As Integer, _
                                            <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                            ByVal typeStubPeriod As Integer, _
                                            <ExcelArgument(Description:="Matriz de % de amort y noc vigente")> _
                                            ByVal currentNotionalAndAmort As Object, _
                                            <ExcelArgument(Description:="Tasa fija del Swap")> _
                                            ByVal rate As Double, _
                                            <ExcelArgument(Description:="Tipo de capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360) (por defecto asume comp)")> _
                                            ByVal typeOfRate As String, _
                                            <ExcelArgument(Description:="Nocional")> _
                                            ByVal notional As Double, _
                                            <ExcelArgument(Description:="Tenors de la curva")> _
                                            ByVal tenors As Object, _
                                            <ExcelArgument(Description:="Tasas correpondientes a cada tenor")> _
                                            ByVal rates As Object, _
                                            <ExcelArgument(Description:="360, 365")> _
                                            ByVal basis As Double, _
                                            <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3")> _
                                            ByVal compound As Integer, _
                                            <ExcelArgument(Description:="Al ingresar un valor se asume pago de intereses + capital en último periódo.De lo contrario, sólo pago de interes")> _
                                            ByVal exchange As Object) As Double
        'typeOfRate can be Lin Act/360, Lin 30/360, Com Act/365
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim aux As Double
        Dim auxDate As Date
        Dim i As Integer
        Dim Result As Double
        Dim indexBase As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                Result = 0
            Else
                schedule = calendar(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod)
                numberOfPeriods = UBound(schedule, 1)

                If numberOfPeriods <> UBound(currentNotionalAndAmort, 1) Then Return "Bad Amortizations"
                i = numberOfPeriods
                auxDate = schedule(i, 2)
                aux = 0
                While valueDate < auxDate
                    aux = aux + (interest(schedule(i, 0), schedule(i, 1), rate, typeOfRate, notional * currentNotionalAndAmort(i, 1)) + exchange * notional * currentNotionalAndAmort(i, 0)) _
                    * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors, rates, basis, compound)
                    i = i - 1
                    If i < 0 Then Exit While
                    auxDate = schedule(i, 2)
                End While
            End If
            Result = aux
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de pata fija de un swap (utiliza la función 'calendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
    Shared Function pvFixedSwapLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                            ByVal valueDate As Date, _
                                            <ExcelArgument(Description:="Fecha ini")> _
                                            ByVal startDate As Date, _
                                            <ExcelArgument(Description:="Fecha fin")> _
                                            ByVal endDate As Date, _
                                            <ExcelArgument(Description:="Periocidad")> _
                                         ByVal periodicity As Integer, _
                                         <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                         ByVal adjustment As Integer, _
                                         <ExcelArgument(Description:="Ajuste de fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> _
                                         ByVal pmtAdjustment As Integer, _
                                         <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                         ByVal typeStubPeriod As Integer, _
                                         <ExcelArgument(Description:="Tasa fija del Swap")> _
                                         ByVal rate As Double, _
                                         <ExcelArgument(Description:="Tipo de capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360) (por defecto asume comp)")> _
                                         ByVal typeOfRate As String, _
                                         <ExcelArgument(Description:="Nocional")> _
                                         ByVal notional As Double, _
                                         <ExcelArgument(Description:="Tenors de la curva")> _
                                         ByVal tenors As Object, _
                                         <ExcelArgument(Description:="Tasas correpondientes a cada tenor")> _
                                         ByVal rates As Object, _
                                         <ExcelArgument(Description:="360, 365")> _
                                         ByVal basis As Double, _
                                         <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3")> _
                                         ByVal compound As Integer, _
                                         <ExcelArgument(Description:="Al ingresar un valor se asume pago de intereses + capital en último periódo.De lo contrario, sólo pago de interes")> _
                                         ByVal exchange As Object) As Double
        'typeOfRate can be Lin Act/360, Lin 30/360, Com Act/365
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim aux As Double
        Dim auxDate As Date
        Dim i As Integer
        Dim flag_exchange As Integer
        Dim Result As Double
        Dim indexBase As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                Result = 0
            Else
                schedule = calendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod)
                indexBase = LBound(schedule, 1)
                numberOfPeriods = UBound(schedule, 1)
                'auxDate = endDate
                auxDate = schedule(numberOfPeriods, indexBase + 2)
                i = numberOfPeriods
                flag_exchange = exchange
                While valueDate < auxDate
                    aux = aux + (interest(schedule(i, indexBase), schedule(i, indexBase + 1), rate, typeOfRate, notional) + (flag_exchange) * Indicador(i, numberOfPeriods) * notional) _
                    * GetDiscountFactorFromCurve(Math.Abs(DateDiff(DateInterval.Day, schedule(i, indexBase + 2), valueDate)), tenors, rates, basis, compound)
                    i = i - 1
                    If i < indexBase Then Exit While
                    auxDate = schedule(i, indexBase + 2)
                End While
                Result = aux
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Esta función calcula los intereses de pata fija de un swap, a partir del periodo al que pertenece valueDate")> _
    Shared Function accruedInterestFixedSwapLeg(<ExcelArgument(Description:="Fecha de cálculo")> _
                                                ByVal valueDate As Date, _
                                                <ExcelArgument(Description:="Fecha inicio")> _
                                                ByVal startDate As Date, _
                                                <ExcelArgument(Description:="Fecha fin")> _
                                                ByVal endDate As Date, _
                                                <ExcelArgument(Description:="Periocidad")> _
                                                ByVal periodicity As Integer, _
                                                <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                                ByVal adjustment As Integer, _
                                                <ExcelArgument(Description:="Ajuste de fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> _
                                                ByVal pmtAdjustment As Integer, _
                                                <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                                ByVal typeStubPeriod As Integer, _
                                                <ExcelArgument(Description:="tasa de la pata fija del swap")> _
                                                ByVal rate As Double, _
                                                <ExcelArgument(Description:="Tipo de capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360) (por defecto asume comp)")> _
                                                ByVal typeOfRate As String, _
                                                <ExcelArgument(Description:="nocional")> _
                                                ByVal notional As Double) As Double
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim initialDate As Date
        Dim i As Integer
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            schedule = calendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod)
            numberOfPeriods = schedule.GetUpperBound(1)
            If valueDate >= endDate Then
                REM VALIDACION: accruedInterestFixedSwapLeg = "valueDate>=endDate" : Exit Function
            ElseIf valueDate < startDate Then
                Result = 0
            Else
                For i = 1 To numberOfPeriods
                    If valueDate >= schedule(i, 1) And valueDate <= schedule(i, 2) Then
                        initialDate = schedule(i, 1)
                    End If
                Next i
                Result = interest(initialDate, valueDate, rate, typeOfRate, notional)
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el valor presente de la pata variable de un swap (utiliza la función 'floatCalendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
     Public Shared Function pvFloatAmortSwapLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                             ByVal valueDate As Date, _
                                             <ExcelArgument(Description:="Fecha ini")> _
                                             ByVal startDate As Date, _
                                             <ExcelArgument(Description:="Fecha fin")> _
                                             ByVal endDate As Date, _
                                             <ExcelArgument(Name:="Per", Description:="Periocidad")> _
                                             ByVal periodicity As Integer, _
                                             <ExcelArgument(Name:="Adj", Description:="ver función floatCalendar")> _
                                             ByVal accAdjustment As Integer, _
                                             <ExcelArgument(Name:="PerAdj", Description:="Ver función floatCalendar")> _
                                             ByVal pmtAdjustment As Integer, _
                                             <ExcelArgument(Name:="typeStubPer", Description:="Periodo corto: 0=al final; 1=al principio")> _
                                             ByVal typeStubPeriod As Integer, _
                                             <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1=un día hábil; 2=dos días hábiles ")> _
                                             ByVal fixingLag As Integer, _
                                             <ExcelArgument(Description:="Cada cuantos periodos fija")> _
                                             ByVal fixingRatio As Integer, _
                                             <ExcelArgument(Name:="fixingStubPer", Description:="Periodo fijación única: 0:al final; 1:al principio")> _
                                             ByVal fixingStubPeriod As Integer, _
                                             <ExcelArgument(Name:="typeAmort", Description:="1:bullet, 2:constante, 3:frances, else:bullet")> _
                                             ByVal typeAmortization As Integer, _
                                             <ExcelArgument(Name:="amortRate", Description:="Tasa para frances Com 30/360")> _
                                             ByVal amortizeRate As Double, _
                                             <ExcelArgument(Name:="typeAmortRate", Description:="No implementado")> _
                                             ByVal typeAmortizeRate As String, _
                                             <ExcelArgument(Description:="Última tasa fijada antes de valueDate")> _
                                             ByVal lastFixing As Double, _
                                             <ExcelArgument(Description:="Spread")> _
                                             ByVal spread As Double, _
                                             <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360")> _
                                             ByVal typeOfSpAndR As String, _
                                             <ExcelArgument(Description:="Tenor tasa ref")> _
                                             ByVal floatIndexTenor As String, _
                                             <ExcelArgument(Description:="Ajuste tasa ref")> _
                                             ByVal floatIndexShift As Integer, _
                                             <ExcelArgument(Description:="Nocional")> _
                                             ByVal notional As Double, _
                                             <ExcelArgument(Description:="Tenors tasa ref")> _
                                             ByVal tenors As Object, _
                                             <ExcelArgument(Description:="Tasas a los tenors")> _
                                             ByVal rates As Object, _
                                             <ExcelArgument(Description:="Base tasa referencia (360,365)")> _
                                             ByVal basis As Double, _
                                             <ExcelArgument(Name:="comp", Description:=".")> _
                                             ByVal compound As Integer, _
                                             <ExcelArgument(Description:="")> _
                                             ByVal cmsFlag As Integer, _
                                             ByVal exchange As Object) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        'floatIndexTenor se ingresa en terminología tenor
        Dim flag_exchange As Integer
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim i As Integer
        Dim interestOfPeriod() As Double
        Dim aux As Double
        Dim auxDate, firstDate, secondDate As Date
        Dim df1 As Double
        Dim df2 As Double
        Dim currentNotional, fwdRate As Double
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                Result = 0
            Else
                schedule = floatCalendarAmortize(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, amortizeRate, typeAmortizeRate)
                numberOfPeriods = UBound(schedule, 1)
                ReDim interestOfPeriod(numberOfPeriods)
                Result = 0.0

                For i = 0 To numberOfPeriods
                    currentNotional = notional * schedule(i, 5)
                    If valueDate >= schedule(i, 2) Then
                        interestOfPeriod(i) = 0
                        currentNotional = 0
                        schedule(i, 4) = 0
                    Else
                        If valueDate >= schedule(i, 3) Then
                            If cmsFlag = 1 Then
                                aux = -currentNotional + (currentNotional + interest(schedule(i, 0), valueDate, lastFixing, typeOfSpAndR, currentNotional)) / _
                                GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 1)), tenors, rates, basis, compound)
                                interestOfPeriod(i) = aux + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                            Else
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), lastFixing + spread, typeOfSpAndR, currentNotional)
                            End If
                        Else
                            If cmsFlag = 1 Then
                                auxDate = schedule(i, 0)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                auxDate = schedule(i, 1)
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = currentNotional * (aux - 1) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                            Else
                                firstDate = shift(schedule(i, 3), floatIndexShift)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, firstDate), tenors, rates, basis, compound)
                                secondDate = BussDay(AddMonths(firstDate, Ten(floatIndexTenor)))
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, secondDate), tenors, rates, basis, compound)
                                aux = df2 / df1
                                fwdRate = GetRateFromDf(firstDate, secondDate, aux, typeOfSpAndR)
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), fwdRate + spread, typeOfSpAndR, currentNotional)
                            End If
                        End If
                    End If
                    Result += (interestOfPeriod(i) + exchange * notional * schedule(i, 4)) * _
                    GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors, rates, basis, compound)
                Next i
            End If

        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el valor presente de la pata variable de un swap (utiliza la función 'floatCalendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
    Public Shared Function pvFloatCustomAmortSwapLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                            ByVal valueDate As Date, _
                                            <ExcelArgument(Description:="Fecha ini")> _
                                            ByVal startDate As Date, _
                                            <ExcelArgument(Description:="Fecha fin")> _
                                            ByVal endDate As Date, _
                                            <ExcelArgument(Name:="Per", Description:="Periocidad")> _
                                            ByVal periodicity As Integer, _
                                            <ExcelArgument(Name:="Adj", Description:="ver función floatCalendar")> _
                                            ByVal accAdjustment As Integer, _
                                            <ExcelArgument(Name:="PerAdj", Description:="Ver función floatCalendar")> _
                                            ByVal pmtAdjustment As Integer, _
                                            <ExcelArgument(Name:="typeStubPer", Description:="Periodo corto: 0=al final; 1=al principio")> _
                                            ByVal typeStubPeriod As Integer, _
                                            <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1=un día hábil; 2=dos días hábiles ")> _
                                            ByVal fixingLag As Integer, _
                                            <ExcelArgument(Description:="Cada cuantos periodos fija")> _
                                            ByVal fixingRatio As Integer, _
                                            <ExcelArgument(Name:="fixingStubPer", Description:="Periodo fijación única: 0:al final; 1:al principio")> _
                                            ByVal fixingStubPeriod As Integer, _
                                            <ExcelArgument(Name:="AmoNot", Description:="Amorta y noc. vig.")> _
                                            ByVal currentNotionalAndAmort As Object, _
                                            <ExcelArgument(Description:="Última tasa fijada antes de valueDate")> _
                                            ByVal lastFixing As Double, _
                                            <ExcelArgument(Description:="Spread")> _
                                            ByVal spread As Double, _
                                            <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360")> _
                                            ByVal typeOfSpAndR As String, _
                                            <ExcelArgument(Description:="Tenor tasa ref")> _
                                            ByVal floatIndexTenor As String, _
                                            <ExcelArgument(Description:="Ajuste tasa ref")> _
                                            ByVal floatIndexShift As Integer, _
                                            <ExcelArgument(Description:="Nocional")> _
                                            ByVal notional As Double, _
                                            <ExcelArgument(Description:="Tenors tasa ref")> _
                                            ByVal tenors As Object, _
                                            <ExcelArgument(Description:="Tasas a los tenors")> _
                                            ByVal rates As Object, _
                                            <ExcelArgument(Description:="Base tasa referencia (360,365)")> _
                                            ByVal basis As Double, _
                                            <ExcelArgument(Name:="comp", Description:=".")> _
                                            ByVal compound As Integer, _
                                            <ExcelArgument(Description:="")> _
                                            ByVal cmsFlag As Integer, _
                                            ByVal exchange As Object) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        'floatIndexTenor se ingresa en terminología tenor

        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim i As Integer
        Dim interestOfPeriod() As Double
        Dim aux As Double
        Dim auxDate, firstDate, secondDate As Date
        Dim df1 As Double
        Dim df2 As Double
        Dim currentNotional, fwdRate As Double
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                Result = 0
            Else
                schedule = floatCalendar(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                If numberOfPeriods <> UBound(currentNotionalAndAmort, 1) Then Return "Bad Amortizations"
                ReDim interestOfPeriod(numberOfPeriods)
                Result = 0

                For i = 0 To numberOfPeriods
                    currentNotional = currentNotionalAndAmort(i, 1) * notional
                    If valueDate >= schedule(i, 2) Then
                        interestOfPeriod(i) = 0
                        currentNotional = 0
                        currentNotionalAndAmort(i, 0) = 0
                    Else
                        If valueDate >= schedule(i, 3) Then
                            If cmsFlag = 1 Then
                                aux = -currentNotional + (currentNotional + interest(schedule(i, 0), valueDate, lastFixing, typeOfSpAndR, currentNotional)) _
                                / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 1)), tenors, rates, basis, compound)
                                interestOfPeriod(i) = aux + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                            Else
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), lastFixing + spread, typeOfSpAndR, currentNotional)
                            End If
                        Else
                            If cmsFlag = 1 Then
                                auxDate = schedule(i, 0)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                auxDate = schedule(i, 1)
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = currentNotional * (aux - 1) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                            Else
                                firstDate = shift(schedule(i, 3), floatIndexShift)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, firstDate), tenors, rates, basis, compound)
                                secondDate = BussDay(AddMonths(firstDate, Ten(floatIndexTenor)))
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, secondDate), tenors, rates, basis, compound)
                                aux = df2 / df1
                                fwdRate = GetRateFromDf(firstDate, secondDate, aux, typeOfSpAndR)
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), fwdRate + spread, typeOfSpAndR, currentNotional)
                            End If
                        End If
                    End If
                    Result += (interestOfPeriod(i) + exchange * notional * currentNotionalAndAmort(i, 0)) * _
                    GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors, rates, basis, compound)
                Next i
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try

        Return Result

    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el valor presente de la pata variable de un swap (utiliza la función 'floatCalendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
           Public Shared Function pvFloatSwapLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                                   ByVal valueDate As Date, _
                                                   <ExcelArgument(Description:="Fecha ini")> _
                                                   ByVal startDate As Date, _
                                                   <ExcelArgument(Description:="Fecha fin")> _
                                                   ByVal endDate As Date, _
                                                   <ExcelArgument(Name:="Per", Description:="Periocidad")> _
                                                   ByVal periodicity As Integer, _
                                                   <ExcelArgument(Name:="Adj", Description:="ver función floatCalendar")> _
                                                   ByVal adjustment As Integer, _
                                                   <ExcelArgument(Name:="PerAdj", Description:="Ver función floatCalendar")> _
                                                   ByVal pmtAdjustment As Integer, _
                                                   <ExcelArgument(Name:="typeStubPer", Description:="Periodo corto: 0=al final; 1=al principio")> _
                                                   ByVal typeStubPeriod As Integer, _
                                                   <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1=un día hábil; 2=dos días hábiles ")> _
                                                   ByVal fixingLag As Integer, _
                                                   <ExcelArgument(Description:="Periocidad de fijación: n, se fija cada n periodos")> _
                                                   ByVal fixingRatio As Integer, _
                                                   <ExcelArgument(Name:="fixingStubPer", Description:="Periodo fijación única: 0:al final; 1:al principio")> _
                                                   ByVal fixingStubPeriod As Integer, _
                                                   <ExcelArgument(Description:="Última tasa fijada antes de valueDate")> _
                                                   ByVal lastFixing As Double, _
                                                   <ExcelArgument(Description:="Spread")> _
                                                   ByVal spread As Double, _
                                                   <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360")> _
                                                   ByVal typeOfSpreadAndRate As String, _
                                                   <ExcelArgument(Description:="Tenor tasa ref")> _
                                                   ByVal floatIndexTenor As String, _
                                                   <ExcelArgument(Description:="Ajuste tasa ref")> _
                                                   ByVal floatIndexShift As Integer, _
                                                   <ExcelArgument(Description:="Nocional")> _
                                                   ByVal notional As Double, _
                                                   <ExcelArgument(Description:="Tenors tasa ref")> _
                                                   ByVal tenors As Object, _
                                                   <ExcelArgument(Description:="Tasas a los tenors")> _
                                                   ByVal rates As Object, _
                                                   <ExcelArgument(Description:="Base tasa referencia (360,365)")> _
                                                   ByVal basis As Double, _
                                                   <ExcelArgument(Name:="comp", Description:=".")> _
                                                   ByVal compound As Integer, _
                                                   <ExcelArgument(Description:="")> _
                                                   ByVal cmsFlag As Integer, _
                                                   ByVal exchange As Object) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        'floatIndexTenor se ingresa en terminología tenor
        Dim flag_exchange As Integer
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim i As Integer
        Dim interestOfPeriod() As Double
        Dim aux As Double
        Dim auxDate As Date
        Dim K1 As Date
        Dim K2 As Date
        Dim FechaIni As Date
        Dim FechaFin As Date
        Dim df1 As Double
        Dim df2 As Double
        Dim indexBase As Integer
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                REM VALIDACION: pvFloatSwapLegFromCurve = "endDate<=startDate" : Exit Function
                Result = 0
            Else
                schedule = floatCalendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
                indexBase = LBound(schedule, 1)
                numberOfPeriods = UBound(schedule, 1)
                ReDim interestOfPeriod(0 To numberOfPeriods)
                flag_exchange = exchange
                For i = indexBase To numberOfPeriods
                    If valueDate >= schedule(i, indexBase + 1) Then
                        interestOfPeriod(i) = 0
                    Else
                        If valueDate >= schedule(i, indexBase + 3) Then
                            If cmsFlag = 1 Then
                                aux = -notional + (notional + interest(schedule(i, indexBase), valueDate, lastFixing, typeOfSpreadAndRate, notional)) / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, indexBase + 1)), tenors, rates, basis, compound)
                                interestOfPeriod(i) = aux + interest(schedule(i, indexBase), schedule(i, indexBase + 1), spread, typeOfSpreadAndRate, notional)
                            Else
                                interestOfPeriod(i) = interest(schedule(i, indexBase), schedule(i, indexBase + 1), lastFixing + spread, typeOfSpreadAndRate, notional)
                            End If
                        Else
                            If cmsFlag = 1 Then
                                auxDate = schedule(i, indexBase)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                auxDate = schedule(i, indexBase + 1)
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = notional * (aux - 1) + interest(schedule(i, indexBase), schedule(i, indexBase + 1), spread, typeOfSpreadAndRate, notional)
                            Else
                                FechaIni = schedule(i, 0)
                                FechaFin = schedule(i, 1)
                                K1 = shift(schedule(i, indexBase + 3), floatIndexShift)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, K1), tenors, rates, basis, compound)
                                K2 = BussDay(AddMonths(K1, Ten(floatIndexTenor)))
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, K2), tenors, rates, basis, compound)
                                aux = ((df1 / df2) - 1) * (DateDiff(DateInterval.Day, FechaIni, FechaFin) / DateDiff(DateInterval.Day, K1, K2))
                                interestOfPeriod(i) = notional * (aux) + interest(schedule(i, indexBase), schedule(i, indexBase + 1), spread, typeOfSpreadAndRate, notional)
                            End If
                        End If
                    End If
                    Result = Result + (interestOfPeriod(i) + (flag_exchange) * Indicador(i, numberOfPeriods) * notional) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, indexBase + 2)), tenors, rates, basis, compound)
                Next i
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Entrega la fecha del último fixing en una pata variable tipo Libor. Si no ha fijado es 0")> _
       Public Shared Function currentFixingDate(<ExcelArgument(Description:="Fecha val")> _
                                               ByVal valueDate As Date, _
                                               <ExcelArgument(Description:="Fecha ini")> _
                                               ByVal startDate As Date, _
                                               <ExcelArgument(Description:="Fecha fin")> _
                                               ByVal endDate As Date, _
                                               <ExcelArgument(Name:="Per", Description:="Periocidad")> _
                                               ByVal periodicity As Integer, _
                                               <ExcelArgument(Name:="Adj", Description:="ver función floatCalendar")> _
                                               ByVal adjustment As Integer, _
                                               <ExcelArgument(Name:="PerAdj", Description:="Ver función floatCalendar")> _
                                               ByVal pmtAdjustment As Integer, _
                                               <ExcelArgument(Name:="typeStubPer", Description:="Periodo corto: 0=al final; 1=al principio")> _
                                               ByVal typeStubPeriod As Integer, _
                                               <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1=un día hábil; 2=dos días hábiles ")> _
                                               ByVal fixingLag As Integer, _
                                               <ExcelArgument(Description:="Periocidad de fijación: n, se fija cada n periodos")> _
                                               ByVal fixingRatio As Integer, _
                                               <ExcelArgument(Name:="fixingStubPer", Description:="Periodo fijación única: 0:al final; 1:al principio")> _
                                               ByVal fixingStubPeriod As Integer) As Date
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim i As Integer
        Dim Result As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                REM VALIDACION: pvFloatSwapLegFromCurve = "endDate<=startDate" : Exit Function
                Result = "1/1/1900"
            Else
                schedule = floatCalendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                For i = 0 To numberOfPeriods
                    If valueDate >= schedule(i, 3) Then
                        Result = schedule(i, 3)
                    End If
                Next i
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Entrega la fecha de inicio del cupón corriente de cualquier pata. Si no ha partido es 0")> _
           Public Shared Function currentStartDate(<ExcelArgument(Description:="Fecha val")> _
                                                   ByVal valueDate As Date, _
                                                   <ExcelArgument(Description:="Fecha ini")> _
                                                   ByVal startDate As Date, _
                                                   <ExcelArgument(Description:="Fecha fin")> _
                                                   ByVal endDate As Date, _
                                                   <ExcelArgument(Name:="Per", Description:="Periocidad")> _
                                                   ByVal periodicity As Integer, _
                                                   <ExcelArgument(Name:="Adj", Description:="ver función floatCalendar")> _
                                                   ByVal adjustment As Integer, _
                                                   <ExcelArgument(Name:="PerAdj", Description:="Ver función floatCalendar")> _
                                                   ByVal pmtAdjustment As Integer, _
                                                   <ExcelArgument(Name:="typeStubPer", Description:="Periodo corto: 0=al final; 1=al principio")> _
                                                   ByVal typeStubPeriod As Integer) As Date

        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim i As Integer
        Dim Result As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                REM VALIDACION: pvFloatSwapLegFromCurve = "endDate<=startDate" : Exit Function
                Result = "1/1/1900"
            Else
                schedule = calendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                For i = 0 To numberOfPeriods
                    If valueDate >= schedule(i, 1) Then
                        Result = schedule(i, 1)
                    End If
                Next i
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Entrega tasa USD onshore en función de tasa en CLP, spot y puntos forwards")> _
        Shared Function cldRateFromFwdPoints(<ExcelArgument(Description:="Tiempo en días al que se calcula la tasa")> _
                                             ByVal days As Integer, _
                                             <ExcelArgument(Description:="Spot USD/CLP")> _
                                             ByVal spot As Double, _
                                             <ExcelArgument(Description:="puntos forwdars ")> _
                                             ByVal points As Double, _
                                             <ExcelArgument(Description:="tenors curva en CLP")> _
                                             ByVal tenors() As Double, _
                                             <ExcelArgument(Description:="tasas a los correspondientes tenors")> _
                                             ByVal rates() As Double, _
                                             <ExcelArgument(Description:="Base tasa en CLP (360, 365)")> _
                                             ByVal basis As Double, _
                                             <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                             ByVal compound As Integer, _
                                             <ExcelArgument(Description:="Forma capitilación tasa calculada")> _
                                             ByVal linOrComp As Integer, _
                                             <ExcelArgument(Description:="Base tasa calculada")> _
                                             ByVal basisReturn As Double) As Double
        'returns the rate in Act/360
        Dim clpDf As Double
        Dim fwd As Double
        Dim cldDf As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            clpDf = GetDiscountFactorFromCurve(days, tenors, rates, basis, compound)
            fwd = spot + points
            cldDf = fwd / spot * clpDf 'fwd=spot*cldDf/clpDf
            cldDf = 1 / cldDf
            If linOrComp = 1 Then
                Result = (cldDf - 1) * basisReturn / days
            Else
                Result = (cldDf) ^ (basisReturn / days) - 1
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de pata variable de un cross que proyecta con curva y descuenta con otra (utiliza 'Calendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
    Public Shared Function pvFloatAmortCrossLegFromCurve(<ExcelArgument(Description:="Fecha de valoración")> _
                                             ByVal valueDate As Date, _
                                             <ExcelArgument(Description:="Fecha inicio")> _
                                          ByVal startDate As Date, _
                                          <ExcelArgument(Description:="Fecha fin")> _
                                          ByVal endDate As Date, _
                                          <ExcelArgument(Description:="Periocidad")> _
                                          ByVal periodicity As Integer, _
                                          <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                          ByVal accAdjustment As Integer, _
                                          <ExcelArgument(Description:="Ajuste de fechas de pago")> _
                                          ByVal pmtAdjustment As Integer, _
                                          <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                          ByVal typeStubPeriod As Integer, _
                                          <ExcelArgument(Description:="Lag del indice flotante")> _
                                          ByVal fixingLag As Integer, _
                                          <ExcelArgument(Description:="Cada cuantos periodos se fija")> _
                                          ByVal fixingRatio As Integer, _
                                          <ExcelArgument(Description:="Periodo corto de fixing")> _
                                          ByVal fixingStubPeriod As Integer, _
                                          <ExcelArgument(Description:="1:bullet, 2:constante, 3:frances, else:bullet")> _
                                          ByVal typeAmortization As Integer, _
                                          <ExcelArgument(Description:="Tasa frances Com 30/360")> _
                                          ByVal amortizeRate As Double, _
                                          <ExcelArgument(Description:="No implementado")> _
                                          ByVal typeAmortizeRate As String, _
                                          <ExcelArgument(Description:="Última tasa fijada")> _
                                          ByVal lastFixing As Double, _
                                          <ExcelArgument(Description:="spread")> _
                                          ByVal spread As Double, _
                                          <ExcelArgument(Description:="Tipo capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360) (por defecto comp)")> _
                                          ByVal typeOfSpAndR As String, _
                                          <ExcelArgument(Description:="Tenor indice flotante")> _
                                          ByVal floatIndexTenor As String, _
                                          <ExcelArgument(Description:="Start Lag indice flotante")> _
                                          ByVal floatIndexShift As Integer, _
                                          <ExcelArgument(Description:="Nocional")> _
                                          ByVal notional As Double, _
                                          <ExcelArgument(Description:="Tenors en los que se observa la tasa de refencia")> _
                                          ByVal tenors As Object, _
                                          <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                          ByVal rates As Object, _
                                          <ExcelArgument(Description:="Base de la tasa de referencia (360, 365)")> _
                                          ByVal basis As Double, _
                                          <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                          ByVal compound As Integer, _
                                          <ExcelArgument(Description:="cmsFlag=1 es cámara promedio, si no es swap normal")> _
                                          ByVal cmsFlag As Integer, _
                                          ByVal exchange As Integer, _
                                          ByVal tenors1 As Object, _
                                          ByVal rates1 As Object, _
                                          ByVal basis1 As Double, _
                                          ByVal compound1 As Integer) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        Dim schedule(,) As Object
        Dim i, numberOfPeriods As Integer
        Dim aux, df1, df2, fwdRate, currentNotional As Double
        Dim interestOfPeriod() As Double
        Dim Result As Double = 0
        Dim auxDate, firstDate, secondDate As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then Return "endDate<=startDate"

            schedule = floatCalendarAmortize(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, amortizeRate, typeAmortizeRate)
            numberOfPeriods = UBound(schedule, 1)
            ReDim interestOfPeriod(numberOfPeriods)
            Result = 0

            For i = 0 To numberOfPeriods
                currentNotional = notional * schedule(i, 5)
                If valueDate >= schedule(i, 1) Then
                    interestOfPeriod(i) = 0
                    currentNotional = 0
                    schedule(i, 4) = 0
                Else
                    If valueDate >= schedule(i, 3) Then
                        If cmsFlag = 1 Then
                            aux = -currentNotional + (currentNotional + interest(schedule(i, 0), valueDate, lastFixing, typeOfSpAndR, currentNotional)) _
                            / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 1)), tenors, rates, basis, compound)
                            interestOfPeriod(i) = aux + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                        Else
                            interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), lastFixing + spread, typeOfSpAndR, currentNotional)
                        End If
                    Else
                        If cmsFlag = 1 Then
                            auxDate = schedule(i, 0)
                            df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                            auxDate = schedule(i, 1)
                            df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                            aux = df1 / df2
                            interestOfPeriod(i) = currentNotional * (aux - 1) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                        Else
                            firstDate = shift(schedule(i, 3), floatIndexShift)
                            df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, firstDate), tenors, rates, basis, compound)
                            secondDate = BussDay(AddMonths(firstDate, Ten(floatIndexTenor)))
                            df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, secondDate), tenors, rates, basis, compound)
                            aux = df2 / df1
                            fwdRate = GetRateFromDf(firstDate, secondDate, aux, typeOfSpAndR)
                            interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), fwdRate + spread, typeOfSpAndR, currentNotional)
                        End If
                    End If
                End If

                Result += (interestOfPeriod(i) + schedule(i, 4) * exchange * notional) _
                * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors1, rates1, basis1, compound1)

            Next i
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de pata variable de un cross que proyecta con curva y descuenta con otra (utiliza 'Calendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
    Public Shared Function pvFloatCustomAmortCrossLegFromCurve(<ExcelArgument(Description:="Fecha de valoración")> _
                                             ByVal valueDate As Date, _
                                             <ExcelArgument(Description:="Fecha inicio")> _
                                          ByVal startDate As Date, _
                                          <ExcelArgument(Description:="Fecha fin")> _
                                          ByVal endDate As Date, _
                                          <ExcelArgument(Description:="Periocidad")> _
                                          ByVal periodicity As Integer, _
                                          <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                          ByVal accAdjustment As Integer, _
                                          <ExcelArgument(Description:="Ajuste de fechas de pago")> _
                                          ByVal pmtAdjustment As Integer, _
                                          <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                          ByVal typeStubPeriod As Integer, _
                                          <ExcelArgument(Description:="Lag del indice flotante")> _
                                          ByVal fixingLag As Integer, _
                                          <ExcelArgument(Description:="Cada cuantos periodos se fija")> _
                                          ByVal fixingRatio As Integer, _
                                          <ExcelArgument(Description:="Periodo corto de fixing")> _
                                          ByVal fixingStubPeriod As Integer, _
                                          <ExcelArgument(Description:="Amortizaciones y nocional vigente")> _
                                          ByVal currentNotionalAndAmort As Object, _
                                          <ExcelArgument(Description:="Última tasa fijada")> _
                                          ByVal lastFixing As Double, _
                                          <ExcelArgument(Description:="spread")> _
                                          ByVal spread As Double, _
                                          <ExcelArgument(Description:="Tipo capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360) (por defecto comp)")> _
                                          ByVal typeOfSpAndR As String, _
                                          <ExcelArgument(Description:="Tenor indice flotante")> _
                                          ByVal floatIndexTenor As String, _
                                          <ExcelArgument(Description:="Start Lag indice flotante")> _
                                          ByVal floatIndexShift As Integer, _
                                          <ExcelArgument(Description:="Nocional")> _
                                          ByVal notional As Double, _
                                          <ExcelArgument(Description:="Tenors en los que se observa la tasa de refencia")> _
                                          ByVal tenors As Object, _
                                          <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                          ByVal rates As Object, _
                                          <ExcelArgument(Description:="Base de la tasa de referencia (360, 365)")> _
                                          ByVal basis As Double, _
                                          <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                          ByVal compound As Integer, _
                                          <ExcelArgument(Description:="cmsFlag=1 es cámara promedio, si no es swap normal")> _
                                          ByVal cmsFlag As Integer, _
                                          ByVal exchange As Integer, _
                                          ByVal tenors1 As Object, _
                                          ByVal rates1 As Object, _
                                          ByVal basis1 As Double, _
                                          ByVal compound1 As Integer) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        Dim schedule(,) As Object
        Dim i, numberOfPeriods As Integer
        Dim aux, df1, df2, fwdRate, currentNotional As Double
        Dim interestOfPeriod() As Double
        Dim Result As Double = 0
        Dim auxDate, firstDate, secondDate As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then Return "endDate<=startDate"

            schedule = floatCalendar(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
            numberOfPeriods = UBound(schedule, 1)
            If numberOfPeriods <> UBound(currentNotionalAndAmort, 1) Then Return "Bad Amortizations"
            If 1 <> UBound(currentNotionalAndAmort, 2) Then Return "Bad Amortizations"
            ReDim interestOfPeriod(numberOfPeriods)
            Result = 0

            For i = 0 To numberOfPeriods
                currentNotional = notional * currentNotionalAndAmort(i, 1)
                If valueDate >= schedule(i, 2) Then
                    interestOfPeriod(i) = 0
                    currentNotional = 0
                    currentNotionalAndAmort(i, 0) = 0
                Else
                    If valueDate >= schedule(i, 3) Then
                        If cmsFlag = 1 Then
                            aux = -currentNotional + (currentNotional + interest(schedule(i, 0), valueDate, lastFixing, typeOfSpAndR, currentNotional)) _
                            / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 1)), tenors, rates, basis, compound)
                            interestOfPeriod(i) = aux + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                        Else
                            interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), lastFixing + spread, typeOfSpAndR, currentNotional)
                        End If
                    Else
                        If cmsFlag = 1 Then
                            auxDate = schedule(i, 0)
                            df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                            auxDate = schedule(i, 1)
                            df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                            aux = df1 / df2
                            interestOfPeriod(i) = currentNotional * (aux - 1) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, currentNotional)
                        Else
                            firstDate = shift(schedule(i, 3), floatIndexShift)
                            df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, firstDate), tenors, rates, basis, compound)
                            secondDate = BussDay(AddMonths(firstDate, Ten(floatIndexTenor)))
                            df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, secondDate), tenors, rates, basis, compound)
                            aux = df2 / df1
                            fwdRate = GetRateFromDf(firstDate, secondDate, aux, typeOfSpAndR)
                            interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), fwdRate + spread, typeOfSpAndR, currentNotional)
                        End If
                    End If
                End If

                Result += (interestOfPeriod(i) + currentNotionalAndAmort(i, 0) * exchange * notional) _
                * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors1, rates1, basis1, compound1)

            Next i
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de pata variable de un cross que proyecta con curva y descuenta con otra (utiliza 'Calendar', 'interest' y 'GetDiscoundFactorFromCurve')")> _
Public Shared Function pvFloatCrossLegFromCurve(<ExcelArgument(Description:="Fecha de valoración")> _
                                         ByVal valueDate As Date, _
                                         <ExcelArgument(Description:="Fecha inicio")> _
                                      ByVal startDate As Date, _
                                      <ExcelArgument(Description:="Fecha fin")> _
                                      ByVal endDate As Date, _
                                      <ExcelArgument(Description:="Periocidad")> _
                                      ByVal periodicity As Integer, _
                                      <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                      ByVal accAdjustment As Integer, _
                                      <ExcelArgument(Description:="Ajuste de fechas de pago")> _
                                      ByVal pmtAdjustment As Integer, _
                                      <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                      ByVal typeStubPeriod As Integer, _
                                      <ExcelArgument(Description:="Lag del indice flotante")> _
                                      ByVal fixingLag As Integer, _
                                      <ExcelArgument(Description:="Cada cuantos periodos se fija")> _
                                      ByVal fixingRatio As Integer, _
                                      <ExcelArgument(Description:="Periodo corto de fixing")> _
                                      ByVal fixingStubPeriod As Integer, _
                                      <ExcelArgument(Description:="Última tasa fijada")> _
                                      ByVal lastFixing As Double, _
                                      <ExcelArgument(Description:="spread")> _
                                      ByVal spread As Double, _
                                      <ExcelArgument(Description:="Tipo capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360) (por defecto comp)")> _
                                      ByVal typeOfSpAndR As String, _
                                      <ExcelArgument(Description:="Tenor indice flotante")> _
                                      ByVal floatIndexTenor As String, _
                                      <ExcelArgument(Description:="Start Lag indice flotante")> _
                                      ByVal floatIndexShift As Integer, _
                                      <ExcelArgument(Description:="Nocional")> _
                                      ByVal notional As Double, _
                                      <ExcelArgument(Description:="Tenors en los que se observa la tasa de refencia")> _
                                      ByVal tenors As Object, _
                                      <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                      ByVal rates As Object, _
                                      <ExcelArgument(Description:="Base de la tasa de referencia (360, 365)")> _
                                      ByVal basis As Double, _
                                      <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                      ByVal compound As Integer, _
                                      <ExcelArgument(Description:="cmsFlag=1 es cámara promedio, si no es swap normal")> _
                                      ByVal cmsFlag As Integer, _
                                      ByVal exchange As Integer, _
                                      ByVal tenors1 As Object, _
                                      ByVal rates1 As Object, _
                                      ByVal basis1 As Double, _
                                      ByVal compound1 As Integer) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        Dim schedule(,) As Object
        Dim i, numberOfPeriods As Integer
        Dim aux, df1, df2, fwdRate As Double
        Dim interestOfPeriod() As Double
        Dim Result As Double = 0
        Dim auxDate, firstDate, secondDate As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                Result = Nothing
            Else
                schedule = floatCalendar(startDate, endDate, periodicity, accAdjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                ReDim interestOfPeriod(numberOfPeriods)
                For i = 0 To numberOfPeriods
                    If valueDate >= schedule(i, 1) Then
                        interestOfPeriod(i) = 0
                    Else
                        If valueDate >= schedule(i, 3) Then
                            If cmsFlag = 1 Then
                                aux = -notional + (notional + interest(schedule(i, 0), valueDate, lastFixing, typeOfSpAndR, notional)) _
                                / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 1)), tenors, rates, basis, compound)
                                interestOfPeriod(i) = aux + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, notional)
                            Else
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), lastFixing + spread, typeOfSpAndR, notional)
                            End If
                        Else
                            If cmsFlag = 1 Then
                                auxDate = schedule(i, 0)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                auxDate = schedule(i, 1)
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = notional * (aux - 1) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpAndR, notional)
                            Else
                                firstDate = shift(schedule(i, 3), floatIndexShift)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, firstDate), tenors, rates, basis, compound)
                                secondDate = BussDay(AddMonths(firstDate, Ten(floatIndexTenor)))
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, secondDate), tenors, rates, basis, compound)
                                aux = df2 / df1
                                fwdRate = GetRateFromDf(firstDate, secondDate, aux, typeOfSpAndR)
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), fwdRate + spread, typeOfSpAndR, notional)
                            End If
                        End If
                    End If

                    Result += (interestOfPeriod(i) + notional * exchange * Indicador(i, numberOfPeriods)) _
                    * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors1, rates1, basis1, compound1)

                Next i
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de la pata variable de un swap que poyecta con una curva y descuenta con otra (utiliza floatCalendar, interest y GetDiscoundFactorFromCurve)")> _
    Shared Function pvFloatSwapLegFromDoubleCurve(<ExcelArgument(Name:="valueD", Description:="Fecha val")> _
                                                 ByVal valueDate As Date, _
                                                 <ExcelArgument(Name:="startD", Description:="Fecha ini")> _
                                                  ByVal startDate As Date, _
                                                 <ExcelArgument(Name:="endD", Description:="Fecha fin")> _
                                                  ByVal endDate As Date, _
                                                  <ExcelArgument(Name:="per", Description:="Periocidad")> _
                                                  ByVal periodicity As Integer, _
                                                  <ExcelArgument(Name:="adj", Description:="Ver función floatCalendar")> _
                                                  ByVal adjustment As Integer, _
                                                  <ExcelArgument(Name:="pmtAdj", Description:="Ver función floatCalendar")> _
                                                  ByVal pmtAdjustment As Integer, _
                                                  <ExcelArgument(Name:="typeStubPer", Description:="Periodo corto: 0=al final; 1=al principio")> _
                                                  ByVal typeStubPeriod As Integer, _
                                                  <ExcelArgument(Description:="Ver función floatCalendar")> _
                                                  ByVal fixingLag As Integer, _
                                                  <ExcelArgument(Description:="Periocidad de fijación: n, se fija cada n per")> _
                                                  ByVal fixingRatio As Integer, _
                                                  <ExcelArgument(Name:="fixingStubPer", Description:="Periodo de fijación única: 0=al final; 1=al principio")> _
                                                  ByVal fixingStubPeriod As Integer, _
                                                  <ExcelArgument(Description:="Última tasa fijada")> _
                                                  ByVal lastFixing As Double, _
                                                  <ExcelArgument(name:="Sprd", Description:="Spread")> _
                                                  ByVal spread As Double, _
                                                  <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360")> _
                                                  ByVal typeOfSpreadAndRate As String, _
                                                  <ExcelArgument(Description:="Tenor de la tasa de ref")> _
                                                  ByVal floatIndexTenor As String, _
                                                  <ExcelArgument(Description:="Ajuste de la tasa de ref ")> _
                                                  ByVal floatIndexShift As Integer, _
                                                  <ExcelArgument(Description:="Nocional")> _
                                                  ByVal notional As Double, _
                                                  <ExcelArgument(Description:="Tenors tasa ref")> _
                                                  ByVal tenors As Object, _
                                                  <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                                  ByVal rates As Object, _
                                                  <ExcelArgument(Description:="Base tasa referencia (360, 365)")> _
                                                  ByVal basis As Double, _
                                                  <ExcelArgument(Name:="comp", Description:=".")> _
                                                  ByVal compound As Integer, _
                                                  <ExcelArgument(Description:="")> _
                                                  ByVal tenors2 As Object, _
                                                  <ExcelArgument(Description:="")> _
                                                  ByVal rates2 As Object, _
                                                  <ExcelArgument(Description:="")> _
                                                  ByVal basis2 As Double, _
                                                  <ExcelArgument(Name:="comp2", Description:="")> _
                                                  ByVal compound2 As Integer, _
                                                  ByVal cmsFlag As Integer, _
                                                  ByVal exchange As Object) As Double
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        'floatIndexTenor se ingresa en terminología tenor
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim flag_exchange As Integer
        Dim df1 As Double
        Dim df2 As Double
        Dim auxDate As Date
        Dim K1 As Date
        Dim K2 As Date
        Dim FechaIni As Date
        Dim FechaFin As Date
        Dim aux As Double
        Dim interestOfPeriod() As Double
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                REM VALIDACION:  pvFloatSwapLegFromDoubleCurve = "endDate<=startDate" : Exit Function
                Result = Nothing
            Else
                schedule = floatCalendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                ReDim interestOfPeriod(0 To numberOfPeriods + 1)
                flag_exchange = exchange
                For i = 0 To numberOfPeriods
                    If valueDate >= schedule(i, 1) Then
                        interestOfPeriod(i) = 0
                    Else
                        If valueDate >= schedule(i, 3) Then
                            If cmsFlag = 1 Then
                                aux = -notional + (notional + interest(schedule(i, 0), valueDate, lastFixing, typeOfSpreadAndRate, notional)) / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 1)), tenors, rates, basis, compound)
                                interestOfPeriod(i) = aux + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpreadAndRate, notional)
                            Else
                                interestOfPeriod(i) = interest(schedule(i, 0), schedule(i, 1), lastFixing + spread, typeOfSpreadAndRate, notional)
                            End If
                        Else
                            If cmsFlag = 1 Then
                                auxDate = schedule(i, 0)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                auxDate = schedule(i, 1)
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, auxDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = notional * (aux - 1) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpreadAndRate, notional)
                            Else
                                FechaIni = schedule(i, 0)
                                FechaFin = schedule(i, 1)
                                K1 = shift(schedule(i, 3), floatIndexShift)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, K1), tenors, rates, basis, compound)
                                K2 = BussDay(AddMonths(K1, Ten(floatIndexTenor)))
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, K2), tenors, rates, basis, compound)
                                aux = ((df1 / df2) - 1) * (DateDiff(DateInterval.Day, FechaIni, FechaFin) / DateDiff(DateInterval.Day, K1, K2))
                                interestOfPeriod(i) = notional * (aux) + interest(schedule(i, 0), schedule(i, 1), spread, typeOfSpreadAndRate, notional)
                            End If
                        End If
                    End If
                    Result += (interestOfPeriod(i) + (flag_exchange) * Indicador(i, numberOfPeriods) * notional) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, valueDate, schedule(i, 2)), tenors2, rates2, basis2, compound2)
                Next i
            End If
        Catch ex As Exception
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function interestFloatSwapLegFromCurve(ByVal valueDate As Date, _
                                           ByVal startDate As Date, _
                                           ByVal endDate As Date, _
                                           ByVal periodicity As Integer, _
                                           ByVal adjustment As Integer, _
                                           ByVal pmtAdjustment As Integer, _
                                           ByVal typeStubPeriod As Integer, _
                                           ByVal fixingLag As Integer, _
                                           ByVal fixingRatio As Integer, _
                                           ByVal fixingStubPeriod As Integer, _
                                           ByVal lastFixing As Double, _
                                           ByVal spread As Double, _
                                           ByVal typeOfSpreadAndRate As String, _
                                           ByVal floatIndexTenor As String, _
                                           ByVal floatIndexShift As Integer, _
                                           ByVal notional As Double, _
                                           ByVal tenors() As Double, _
                                           ByVal rates() As Double, _
                                           ByVal basis As Double, _
                                           ByVal compound As Integer, _
                                           ByVal cmsFlag As Integer, _
                                           Optional ByVal exchange As Object = Nothing) As Double()
        'cmsFlag=1 significa que es un cámara promedio, si no es un swap  normal
        'floatIndexTenor se ingresa en terminología tenor
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim interestOfPeriod() As Double
        Dim flag_exchange As Integer
        Dim aux As Double
        Dim auxDate As Date
        Dim df1 As Double
        Dim df2 As Double
        Dim Result As Double()

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                REM VALIDACION: interestFloatSwapLegFromCurve = "endDate<=startDate" : Exit Function
                Result = Nothing
            Else
                schedule = floatCalendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                ReDim interestOfPeriod(0 To numberOfPeriods + 1)
                If IsNothing(exchange) Then flag_exchange = 1 Else flag_exchange = 0
                For i = 1 To numberOfPeriods
                    If valueDate >= schedule(i, 2) Then
                        interestOfPeriod(i) = 0
                    Else
                        If valueDate >= schedule(i, 4) Then
                            If cmsFlag = 1 Then
                                aux = -notional + (notional + interest(schedule(i, 1), valueDate, lastFixing, typeOfSpreadAndRate, notional)) / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, schedule(i, 2), valueDate), tenors, rates, basis, compound)
                                interestOfPeriod(i) = aux + interest(schedule(i, 1), schedule(i, 2), spread, typeOfSpreadAndRate, notional)
                            Else
                                interestOfPeriod(i) = interest(schedule(i, 1), schedule(i, 2), lastFixing + spread, typeOfSpreadAndRate, notional)
                            End If
                        Else
                            If cmsFlag = 1 Then
                                auxDate = schedule(i, 1)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, auxDate, valueDate), tenors, rates, basis, compound)
                                auxDate = schedule(i, 2)
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, auxDate, valueDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = notional * (aux - 1) + interest(schedule(i, 1), schedule(i, 2), spread, typeOfSpreadAndRate, notional)
                            Else
                                auxDate = shift(schedule(i, 4), floatIndexShift)
                                df1 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, auxDate, valueDate), tenors, rates, basis, compound)
                                auxDate = BussDay(AddMonths(auxDate, Ten(floatIndexTenor)))
                                df2 = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, auxDate, valueDate), tenors, rates, basis, compound)
                                aux = df1 / df2
                                interestOfPeriod(i) = notional * (aux - 1) + interest(schedule(i, 1), schedule(i, 2), spread, typeOfSpreadAndRate, notional)
                            End If
                        End If
                    End If
                Next i
                Result = interestOfPeriod
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula la sensibilidad de una pata cualquiera")> _
    Shared Function sensAnyLegFromCurve(<ExcelArgument(Description:="Fecha val")> _
                                        ByVal valueDate As Date, _
                                        <ExcelArgument(Description:="Fecha ini")> _
                                        ByVal startDate As Date, _
                                        <ExcelArgument(Description:="Fecha fin")> _
                                        ByVal endDate As Date, _
                                        <ExcelArgument(Description:="Nocional")> _
                                        ByVal notional As Double, _
                                        <ExcelArgument(Description:="Fijo o flotante")> _
                                        ByVal fixFloat As String, _
                                        <ExcelArgument(Description:="Tenors curva proy.")> _
                                        ByVal projTenors As Object, _
                                        <ExcelArgument(Description:="Tasas curva proy.")> _
                                        ByVal projRates As Object, _
                                        <ExcelArgument(Description:="Tenors curva desc.")> _
                                        ByVal discTenors As Object, _
                                        <ExcelArgument(Description:="Tasas curva desc.")> _
                                        ByVal discRates As Object, _
                                        <ExcelArgument(Description:="Periodicidad")> _
                                        ByVal periodicity As Integer, _
                                        <ExcelArgument(Description:="Ajuste de fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                        ByVal accAdjustment As Integer, _
                                        <ExcelArgument(Description:="Ajuste de fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> _
                                        ByVal pmtAdjustment As Integer, _
                                        <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                        ByVal typeStubPeriod As Integer, _
                                        <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1=un día hábil; 2=dos días hábiles ")> _
                                        ByVal fixingLag As Integer, _
                                        <ExcelArgument(Description:="Periocidad de fijación: n, se fija cada n periodos")> _
                                        ByVal fixingRatio As Integer, _
                                        <ExcelArgument(Name:="fixingStubPer", Description:="Periodo fijación única: 0:al final; 1:al principio")> _
                                        ByVal fixingStubPeriod As Integer, _
                                        <ExcelArgument(Description:="1:bullet, 2:constante, 3:frances, else:bullet")> _
                                        ByVal typeAmortization As Integer, _
                                        <ExcelArgument(Description:="Tasa para frances Com 30/360")> _
                                        ByVal amortizeRate As Double, _
                                        <ExcelArgument(Description:="No implementado")> _
                                        ByVal typeAmortizeRate As String, _
                                        <ExcelArgument(Description:="Amortizaciones y nocional vigente")> _
                                        ByVal currentNotionalAndAmort As Object, _
                                        <ExcelArgument(Description:="Tipo de capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Com 30/360) (por defecto asume comp)")> _
                                        ByVal typeOfSpreadOrRate As String, _
                                        <ExcelArgument(Description:="Tenor tasa ref")> _
                                        ByVal floatIndexTenor As String, _
                                        <ExcelArgument(Description:="Ajuste tasa ref")> _
                                        ByVal floatIndexShift As Integer, _
                                        <ExcelArgument(Description:="1: es cámara, Else: no")> _
                                        ByVal camaraFlag As Integer, _
                                        <ExcelArgument(Description:="1: paga amortizaciones, Else: no paga")> _
                                        ByVal exchange As Integer, _
                                        <ExcelArgument(Description:="Ultimo fixing")> _
                                        ByVal lastFixing As Double, _
                                        <ExcelArgument(Description:="Tasa o spread")> _
                                        ByVal spreadOrRate As Double, _
                                        <ExcelArgument(Description:="Que curva: 1 es proj, 2 es disc, 3 es ambas")> _
                                        ByVal whichCurve As Integer, _
                                        <ExcelArgument(Description:="Cantidad de bp's")> _
                                        ByVal bps As Double) As Object



        'typeOfRate can be Lin Act/360, Lin 30/360, Com Act/365

        Dim Result() As Double

        Dim arrayProjTenors(), arrayProjRates(), arrayDiscTenors(), arrayDiscRates() As Double

        Dim arrayProjTenors2(), arrayProjRates2(), arrayDiscTenors2(), arrayDiscRates2() As Double

        Dim aux As Double

        Dim i, numberOfRates, numProj, numDisc As Integer


        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If



        arrayProjRates = Range2Array(Of Double)(projRates)

        arrayDiscRates = Range2Array(Of Double)(discRates)

        arrayProjTenors = Range2Array(Of Double)(projTenors)

        arrayDiscTenors = Range2Array(Of Double)(discTenors)

        numProj = arrayProjTenors.GetUpperBound(0) - 1

        numDisc = arrayDiscTenors.GetUpperBound(0) - 1



        ReDim arrayProjTenors2(0 To numProj), arrayProjRates2(0 To numProj)

        ReDim arrayDiscTenors2(0 To numDisc), arrayDiscRates2(0 To numDisc)



        For i = 0 To numProj

            arrayProjTenors2(i) = arrayProjTenors(i)

            arrayProjRates2(i) = arrayProjRates(i)

        Next



        For i = 0 To numDisc

            arrayDiscTenors2(i) = arrayDiscTenors(i)

            arrayDiscRates2(i) = arrayDiscRates(i)

        Next







        If whichCurve = 1 Then

            If fixFloat = "Fixed" Then

                numberOfRates = UBound(projTenors, 1)

                ReDim Result(0 To numberOfRates)

                If typeAmortization = 4 Then

                    aux = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, _
                                                            accAdjustment, pmtAdjustment, typeStubPeriod, _
                                                            currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, _
                                                            notional, arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        Result(i) = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, _
                                                                        accAdjustment, pmtAdjustment, typeStubPeriod, _
                                                                        currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, _
                                                                        notional, arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                    Next i

                Else

                    aux = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                       pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, _
                                                       typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, _
                                                       arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        Result(i) = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                                   pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, _
                                                                 typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, _
                                                                   arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                    Next i

                End If

            Else

                numberOfRates = UBound(projTenors, 1)

                ReDim Result(0 To numberOfRates)

                If typeAmortization = 4 Then

                    aux = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                              pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, _
                                                              fixingStubPeriod, currentNotionalAndAmort, lastFixing, _
                                                              spreadOrRate, typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, _
                                                              notional, arrayProjTenors2, arrayProjRates2, 365, 2, camaraFlag, exchange, _
                                                              arrayDiscTenors2, arrayDiscRates2, 365, 2)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        Result(i) = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                                        pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, _
                                                                        fixingStubPeriod, currentNotionalAndAmort, lastFixing, _
                                                                        spreadOrRate, typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, _
                                                                        notional, arrayProjTenors2, arrayProjRates2, 365, 2, camaraFlag, exchange, _
                                                                        arrayDiscTenors2, arrayDiscRates2, 365, 2) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                    Next i

                Else

                    aux = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, _
                                                        typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, _
                                                        amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, _
                                                        floatIndexTenor, floatIndexShift, notional, arrayProjTenors2, arrayProjRates2, 365, 2, _
                                                        camaraFlag, exchange, arrayDiscTenors2, arrayDiscRates2, 365, 2)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        Result(i) = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, _
                                                            typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, _
                                                            amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, _
                                                            floatIndexTenor, floatIndexShift, notional, arrayProjTenors2, arrayProjRates2, 365, 2, _
                                                            camaraFlag, exchange, arrayDiscTenors2, arrayDiscRates2, 365, 2) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                    Next i

                End If

            End If

        ElseIf whichCurve = 2 Then

            If fixFloat = "Fixed" Then

                numberOfRates = UBound(discTenors, 1)

                ReDim Result(0 To numberOfRates)

                If typeAmortization = 4 Then

                    aux = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, _
                                                             accAdjustment, pmtAdjustment, typeStubPeriod, _
                                                             currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, _
                                                             notional, arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange)

                    For i = 0 To numberOfRates

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, _
                                                                       accAdjustment, pmtAdjustment, typeStubPeriod, _
                                                                       currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, _
                                                                       notional, arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange) - aux

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                Else

                    aux = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                       pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, _
                                                       typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, _
                                                       arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange)

                    For i = 0 To numberOfRates

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                           pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, _
                                                           typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, _
                                                           arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange) - aux

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                End If

            Else

                numberOfRates = UBound(discTenors, 1)

                ReDim Result(0 To numberOfRates)

                If typeAmortization = 4 Then

                    aux = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                              pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, _
                                                              fixingStubPeriod, currentNotionalAndAmort, lastFixing, spreadOrRate, _
                                                              typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, notional, _
                                                              arrayProjTenors2, arrayProjRates2, 365, 2, camaraFlag, exchange, _
                                                              arrayDiscTenors2, arrayDiscRates2, 365, 2)

                    For i = 0 To numberOfRates

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                                        pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, _
                                                                        fixingStubPeriod, currentNotionalAndAmort, lastFixing, _
                                                                        spreadOrRate, typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, _
                                                                        notional, arrayProjTenors2, arrayProjRates2, 365, 2, camaraFlag, exchange, _
                                                                        arrayDiscTenors2, arrayDiscRates2, 365, 2) - aux

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                Else

                    aux = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, _
                                                        typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, _
                                                        amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, _
                                                        floatIndexTenor, floatIndexShift, notional, arrayProjTenors2, arrayProjRates2, 365, 2, _
                                                        camaraFlag, exchange, arrayDiscTenors2, arrayDiscRates2, 365, 2)

                    For i = 0 To numberOfRates

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, _
                                                            typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, _
                                                            amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, _
                                                            floatIndexTenor, floatIndexShift, notional, arrayProjTenors2, arrayProjRates2, 365, 2, _
                                                            camaraFlag, exchange, arrayDiscTenors2, arrayDiscRates2, 365, 2) - aux

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                End If

            End If

        Else

            If fixFloat = "Fixed" Then

                numberOfRates = UBound(discTenors, 1)

                ReDim Result(0 To numberOfRates)

                If typeAmortization = 4 Then

                    aux = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, _
                                                             accAdjustment, pmtAdjustment, typeStubPeriod, _
                                                             currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, _
                                                             notional, arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFixedCustomAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, _
                                                                       accAdjustment, pmtAdjustment, typeStubPeriod, _
                                                                       currentNotionalAndAmort, spreadOrRate, typeOfSpreadOrRate, _
                                                                       notional, arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                Else

                    aux = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                       pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, _
                                                       typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, _
                                                       arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFixedAmortSwapLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                           pmtAdjustment, typeStubPeriod, typeAmortization, amortizeRate, _
                                                           typeAmortizeRate, spreadOrRate, typeOfSpreadOrRate, notional, _
                                                           arrayDiscTenors2, arrayDiscRates2, 365, 2, exchange) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                End If

            Else

                numberOfRates = UBound(discTenors, 1)

                ReDim Result(0 To numberOfRates)

                If typeAmortization = 4 Then

                    aux = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                              pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, _
                                                              fixingStubPeriod, currentNotionalAndAmort, lastFixing, spreadOrRate, _
                                                              typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, notional, _
                                                              arrayProjTenors2, arrayProjRates2, 365, 2, camaraFlag, exchange, _
                                                              arrayDiscTenors2, arrayDiscRates2, 365, 2)

                    For i = 0 To numberOfRates



                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps



                        Result(i) = pvFloatCustomAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, _
                                                                        pmtAdjustment, typeStubPeriod, fixingLag, fixingRatio, _
                                                                        fixingStubPeriod, currentNotionalAndAmort, lastFixing, _
                                                                        spreadOrRate, typeOfSpreadOrRate, floatIndexTenor, floatIndexShift, _
                                                                        notional, arrayProjTenors2, arrayProjRates2, 365, 2, camaraFlag, exchange, _
                                                                        arrayDiscTenors2, arrayDiscRates2, 365, 2) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                Else

                    aux = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, _
                                                        typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, _
                                                        amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, _
                                                        floatIndexTenor, floatIndexShift, notional, arrayProjTenors2, arrayProjRates2, 365, 2, _
                                                        camaraFlag, exchange, arrayDiscTenors2, arrayDiscRates2, 365, 2)

                    For i = 0 To numberOfRates

                        arrayProjRates2(i) = arrayProjRates2(i) + bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) + bps

                        Result(i) = pvFloatAmortCrossLegFromCurve(valueDate, startDate, endDate, periodicity, accAdjustment, pmtAdjustment, _
                                                            typeStubPeriod, fixingLag, fixingRatio, fixingStubPeriod, typeAmortization, _
                                                            amortizeRate, typeAmortizeRate, lastFixing, spreadOrRate, typeOfSpreadOrRate, _
                                                            floatIndexTenor, floatIndexShift, notional, arrayProjTenors2, arrayProjRates2, 365, 2, _
                                                            camaraFlag, exchange, arrayDiscTenors2, arrayDiscRates2, 365, 2) - aux

                        arrayProjRates2(i) = arrayProjRates2(i) - bps

                        arrayDiscRates2(i) = arrayDiscRates2(i) - bps

                    Next i

                End If

            End If

        End If

        Return Result

    End Function

#End Region
#Region "BootStrapping"
    Shared Function ITB(ByVal FechaVal As Date, _
                     ByVal Dtenors As String, _
                     ByVal Drates() As Double, _
                     ByVal Dbasis As Double, _
                     ByVal DComp As Integer, _
                     ByVal Stenors() As String, _
                     ByVal Srates() As Double, _
                     ByVal Sdaycount As String, _
                     ByVal Sfreq As Double) As Double()
        Dim N As Integer
        Dim m As Integer
        Dim Dtenors1() As Date
        Dim Stenors1() As Date
        Dim tenor As Double
        Dim i As Integer
        Dim Tenors() As Date
        Dim Rates() As Double
        Dim Tenors1() As Double
        Dim Rates1() As Double
        Dim aux As Double
        Dim diff As Double
        Dim dp As Double
        Dim p As Double
        Dim Result() As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            N = Drates.GetUpperBound(1)
            m = Srates.GetUpperBound(1)
            ReDim Dtenors1(N)
            ReDim Stenors1(m)
            For i = 1 To N
                tenor = Ten(Dtenors(i))
                If UCase(Right(Dtenors(i), 1)) <> "M" And UCase(Right(Dtenors(i), 1)) <> "Y" Then
                    Dtenors1(i) = DateAdd(DateInterval.Day, tenor, FechaVal)

                Else
                    Dtenors1(i) = BussDay(DateAdd(DateInterval.Month, tenor, FechaVal))
                End If
            Next i
            For i = 1 To m
                tenor = Ten(Stenors(i))
                Stenors1(i) = BussDay(DateAdd(DateInterval.Month, tenor, FechaVal))

            Next i
            ReDim Tenors(0 To N + m + 1)
            ReDim Rates(0 To N + m + 1)
            For i = 1 To N
                Tenors(i) = Dtenors1(i)
                Rates(i) = Drates(i)
            Next i
            For i = 1 To m
                Tenors(i + N) = Stenors1(i)
                Rates(i + N) = 0
            Next i
            For i = 1 To m
                ReDim Tenors1(0 To N + i + 1)
                ReDim Rates1(0 To N + i + 1)
                For j = 1 To N + i - 1
                    Tenors1(j) = YearFraction(FechaVal, Tenors(j), "Act/365") * 365
                    Rates1(j) = Rates(j)
                Next j
                Tenors1(N + i) = YearFraction(FechaVal, Stenors1(i), "Act/365") * 365
                aux = 0.02
                diff = 1
                dp = 1.0E+20
                While Math.Abs(diff) > 0.000000000000001
                    Rates1(N + i) = aux - diff / dp
                    p = PVB_PvFromCurve(FechaVal, 1, Srates(i), FechaVal, Stenors1(i), Sdaycount, Sfreq, Tenors1, Rates1, Dbasis, DComp)
                    aux = Rates1(N + i)
                    Rates1(N + i) = Rates1(N + i) + 0.000001
                    dp = (PVB_PvFromCurve(FechaVal, 1, Srates(i), FechaVal, Stenors1(i), Sdaycount, Sfreq, Tenors1, Rates1, Dbasis, DComp) - p) / 0.000001
                    diff = p - 1
                End While
                Rates(N + i) = aux

            Next i

            Result = Rates
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Name:="ITB_2", Description:="Realiza Bootstrapping imponiendo la condición: VP pata fija=Par")> _
    Shared Function ITB2(<ExcelArgument(Description:="Trade date")> ByVal valueDate As Date, _
              <ExcelArgument(Description:="Rezago inicio")> ByVal startLag As Integer, _
              <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360")> ByVal DepoType As String, _
              <ExcelArgument(Description:="Tenores a los que se observan las tasas cero")> ByVal DepoTenors As Object, _
              <ExcelArgument(Description:="Tasas a correspondientes tenors")> ByVal DepoRates As Object, _
              <ExcelArgument(Description:="Periocidad pata fija")> ByVal periodicity As Integer, _
              <ExcelArgument(Description:="Ajuste fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> ByVal accrualAdjustment As Integer, _
              <ExcelArgument(Description:="Ajuste fechas pago (respecto fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> ByVal paymentAdjustment As Integer, _
              <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> ByVal stubPeriod As Integer, _
              <ExcelArgument(Description:="Tipo capitalización y convención conteo de días de swaps(Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360) (por defecto comp)")> ByVal swapType As String, _
              <ExcelArgument(Description:="Tenores a los que se observan swaps")> ByVal swapTenors As Object, _
              <ExcelArgument(Description:="Tasas a los correspondientes tenors")> ByVal swapRates As Object) As Object(,)
        Dim startDate As Date
        Dim cuantosDepos As Integer
        Dim cuantosSwaps As Integer
        Dim auxTenors() As Double
        Dim auxRates() As Double
        Dim aux As Date
        Dim lastEndDate As Date
        Dim longCalendar(,) As Object
        Dim auxPeriods As Integer
        Dim auxPeriodsLong As Double
        Dim tenorsCashFlows() As Double
        Dim cashFlows() As Double
        Dim y0 As Double
        Dim diff As Double
        Dim dp As Double
        Dim nAux As Double
        Dim p As Double
        Dim auxVal As Double
        Dim arrayDepoTenors, arraySwapTenors As String()
        Dim arrayDepoRates, arraySwapRates As Double()
        Dim Result(,) As Object
        Dim fecha As Date = #1/1/1900#
        Dim dias As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            arrayDepoTenors = Range2Array(Of String)(DepoTenors)
            arraySwapTenors = Range2Array(Of String)(swapTenors)
            arrayDepoRates = Range2Array(Of Double)(DepoRates)
            arraySwapRates = Range2Array(Of Double)(swapRates)

            startDate = shift(valueDate, startLag)
            cuantosDepos = arrayDepoTenors.GetUpperBound(0)
            cuantosSwaps = arraySwapTenors.GetUpperBound(0)
            ReDim auxTenors(0 To cuantosDepos + cuantosSwaps - 1)
            ReDim auxRates(0 To cuantosDepos + cuantosSwaps - 1)
            For i = 1 To cuantosDepos
                If Right(arrayDepoTenors(i - 1).ToUpper, 1) = "C" Then
                    aux = Prev2(AddMonthsC(startDate, Ten(arrayDepoTenors(i))))
                ElseIf Right(arrayDepoTenors(i - 1).ToUpper, 1) <> "C" And Right(arrayDepoTenors(i - 1).ToUpper, 1) <> "M" And Right(arrayDepoTenors(i - 1).ToUpper, 1) <> "Y" Then
                    dias = arrayDepoTenors(i - 1) - 2
                    aux = DateAdd(DateInterval.Day, dias, fecha)
                    'aux = CType(arrayDepoTenors(i - 1), Date)
                Else
                    aux = BussDay(AddMonths(startDate, Ten(arrayDepoTenors(i - 1))))
                End If
                auxTenors(i - 1) = DateDiff(DateInterval.Day, startDate, aux)
                auxRates(i - 1) = (1 + interest(startDate, aux, arrayDepoRates(i - 1), DepoType, 1)) ^ (365 / auxTenors(i - 1)) - 1
            Next i
            For i = 1 To cuantosSwaps
                If UCase(Right(arraySwapTenors(i - 1), 1)) = "C" Then
                    aux = AddMonthsC(startDate, Ten(arraySwapTenors(i - 1)))
                Else
                    aux = AddMonths(startDate, Ten(arraySwapTenors(i - 1)))
                End If
                If accrualAdjustment <> 0 Then aux = shift(aux, 0)
                If paymentAdjustment > 2 Then paymentAdjustment = 1
                aux = shift(aux, paymentAdjustment)
                auxTenors(i + cuantosDepos + 1 - 2) = DateDiff(DateInterval.Day, startDate, aux)
                auxRates(i + cuantosDepos + 1 - 2) = 0
            Next i
            If UCase(Right(arraySwapTenors(cuantosSwaps - 1), 1)) = "C" Then
                lastEndDate = AddMonthsC(startDate, Ten(arraySwapTenors(cuantosSwaps - 1)))
            Else
                lastEndDate = AddMonths(startDate, Ten(arraySwapTenors(cuantosSwaps - 1)))
            End If
            longCalendar = calendar(startDate, lastEndDate, periodicity, accrualAdjustment, paymentAdjustment, stubPeriod)
            auxPeriods = UBound(longCalendar, 1)
            auxPeriodsLong = 1 + auxPeriods - Ten(arraySwapTenors(cuantosSwaps - 1)) / periodicity 'OJO
            ReDim tenorsCashFlows(0 To auxPeriods)
            ReDim cashFlows(0 To auxPeriods)
            For i = 1 To auxPeriods
                tenorsCashFlows(i - 1) = 0
                cashFlows(i - 1) = 0
            Next i
            y0 = 10000000
            For i = 1 To cuantosSwaps
                auxRates(i + cuantosDepos + 1 - 2) = 0.0#
                diff = 1
                dp = 1.0E+20
                nAux = 0.0
                auxPeriods = CType(Ten(arraySwapTenors(i - 1)) / periodicity + auxPeriodsLong, Integer)
                For j = 0 To auxPeriods - 1
                    tenorsCashFlows(j) = DateDiff(DateInterval.Day, startDate, longCalendar(j, 2))
                    cashFlows(j) = interest(longCalendar(j, 0), longCalendar(j, 1), arraySwapRates(i - 1), swapType, y0) + y0 * Indicador(j, auxPeriods - 1)
                Next j
                While Math.Abs(diff) > 0.000001
                    auxRates(i + cuantosDepos + 1 - 2) = nAux - diff / dp
                    p = pvCashFlowsFromCurve(tenorsCashFlows, cashFlows, auxTenors, auxRates, 365, 2)
                    nAux = auxRates(i + cuantosDepos + 1 - 2)
                    auxRates(i + cuantosDepos + 1 - 2) += 0.0001
                    auxVal = pvCashFlowsFromCurve(tenorsCashFlows, cashFlows, auxTenors, auxRates, 365, 2)
                    dp = (auxVal - p) / 0.0001
                    diff = p - y0
                    auxRates(i + cuantosDepos + 1 - 2) = nAux
                End While


            Next i
            ReDim Result(0 To cuantosDepos + cuantosSwaps + 1 - 1, 1)
            For i = 1 To cuantosDepos + cuantosSwaps '- 1
                Result(i - 1, 0) = auxTenors(i - 1)
                Result(i - 1, 1) = auxRates(i - 1)
            Next i
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, name:="ITBCross", Description:="Realiza bootstrapping patas flotantes con spread. Asume conocida curva proyección, que patas están a la par y determina curva descuento")> _
    Shared Function ITBCross(<ExcelArgument(Description:="Trade date")> _
                         ByVal valueDate As Date, _
                         <ExcelArgument(Description:="Rezago de inicio")> _
                         ByVal startLag As Integer, _
                         <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360")> _
                         ByVal depoType As String, _
                         <ExcelArgument(Description:="Tenores tasas cero")> _
                         ByVal depoTenors As Object, _
                         <ExcelArgument(Description:="Tasas a los tenors")> _
                         ByVal depoRates As Object, _
                         <ExcelArgument(Name:="Per", Description:="Periocidad pata fija(")> _
                         ByVal periodicity As Integer, _
                         <ExcelArgument(name:="accrualAdj", Description:="Ver floatCalendar(")> _
                         ByVal accrualAdjustment As Integer, _
                         <ExcelArgument(name:="paymentAdj", Description:="Ver floatCalendar(")> _
                         ByVal paymentAdjustment As Integer, _
                         <ExcelArgument(name:="stubPer", Description:="Ver floatCalendar(")> _
                         ByVal stubPeriod As Integer, _
                         <ExcelArgument(Description:="Ver floatCalendar")> _
                         ByVal fixingLag As Integer, _
                         <ExcelArgument(Description:="Ver floatCalendar")> _
                         ByVal fixingRatio As Integer, _
                         <ExcelArgument(Name:="fixingStubPer", Description:="Ver floatCalendar(")> _
                         ByVal fixingStubPeriod As Integer, _
                         <ExcelArgument(Description:="Última tasa fijada")> _
                         ByVal lastFixing As Double, _
                         <ExcelArgument(Description:="Tenor tasa ref")> _
                         ByVal floatIndexTenor As String, _
                         <ExcelArgument(Description:="Ajuste tasa ref")> _
                         ByVal floatIndexShift As Integer, _
                         <ExcelArgument(Description:="Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360")> _
                         ByVal swapType As String, _
                         <ExcelArgument(Description:="Tenores a las tasas swaps")> _
                         ByVal swapTenors As Object, _
                         <ExcelArgument(Description:="Tasas a los tenors")> _
                         ByVal swapRates As Object, _
                         <ExcelArgument(Description:="Tenors tasa ref")> _
                         ByVal tenors As Object, _
                         <ExcelArgument(Description:=".")> _
                         ByVal rates As Object, _
                         ByVal basis As Double, _
                         ByVal compound As Integer, _
                         <ExcelArgument(Name:="flag")> _
                         ByVal cmsflag As Integer, _
                         <ExcelArgument(Name:="exchange")> _
                         ByVal exchange As Integer) As Object
        'Esta versión realiza bootstrapping de patas flotantes con spread. Asume conocida la curva de proyección
        'y determina la curva de descuento, suponiendo que las patas están a la par.
        Dim startDate As Date
        Dim cuantosDepos As Integer
        Dim cuantosSwaps As Integer
        Dim auxTenors() As Double
        Dim auxRates() As Double
        Dim aux As Date
        Dim aux1 As Double
        Dim schedule(,) As Object
        Dim auxPeriods As Integer
        Dim tenorsCashFlows() As Double
        Dim cashFlows() As Double
        Dim y0 As Double
        Dim diff As Double
        Dim dp As Double
        Dim p As Double
        Dim auxVal As Double
        Dim final(,) As Double
        Dim arrayDepoTenors, arraySwapTenors As String()
        Dim arrayDepoRates, arraySwapRates As Double()
        Dim fecha As Date = #1/1/1900#
        Dim dias As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            startDate = shift(valueDate, startLag)
            arrayDepoTenors = Range2Array(Of String)(depoTenors)
            arraySwapTenors = Range2Array(Of String)(swapTenors)
            arrayDepoRates = Range2Array(Of Double)(depoRates)
            arraySwapRates = Range2Array(Of Double)(swapRates)

            cuantosDepos = arrayDepoTenors.GetUpperBound(0)
            cuantosSwaps = arraySwapTenors.GetUpperBound(0)

            ReDim auxTenors(cuantosDepos + cuantosSwaps - 1)
            ReDim auxRates(cuantosDepos + cuantosSwaps - 1)

            For i = 0 To cuantosDepos + cuantosSwaps - 1
                auxTenors(i) = 36500.0 * (i + 1)
                auxRates(i) = 0.0
            Next


            For i = 0 To cuantosDepos - 1
                If Right(arrayDepoTenors(i).ToUpper, 1) <> "C" And _
                   (Right(arrayDepoTenors(i).ToUpper, 1) <> "M" And _
                    (Right(arrayDepoTenors(i).ToUpper, 1) <> "Y")) Then
                    dias = arrayDepoTenors(i) - 2
                    aux = DateAdd(DateInterval.Day, dias, fecha)
                ElseIf Right(arrayDepoTenors(i).ToUpper, 1) = "C" Then
                    aux = Prev2(AddMonthsC(startDate, Ten(arrayDepoTenors(i))))
                Else
                    aux = BussDay(AddMonths(startDate, Ten(arrayDepoTenors(i))))
                End If
                auxTenors(i) = DateDiff(DateInterval.Day, startDate, aux)
                auxRates(i) = (1 + interest(startDate, aux, arrayDepoRates(i), depoType, 1)) ^ (365 / auxTenors(i)) - 1
            Next i

            For i = 1 To cuantosSwaps
                If Right(arraySwapTenors(i - 1).ToString.ToUpper, 1) = "C" Then
                    aux = AddMonthsC(startDate, Ten(arraySwapTenors(i - 1)))
                Else
                    aux = AddMonths(startDate, Ten(arraySwapTenors(i - 1)))
                End If

                schedule = floatCalendar(startDate, aux, periodicity, accrualAdjustment, paymentAdjustment, stubPeriod, _
                                         fixingLag, fixingRatio, fixingStubPeriod)
                auxPeriods = UBound(schedule, 1)
                auxTenors(i + cuantosDepos - 1) = DateDiff(DateInterval.Day, startDate, schedule(auxPeriods, 2))

                y0 = 100000000.0
                auxRates(i + cuantosDepos - 1) = 10.0# : diff = 1 : dp = 1.0E+20 : aux1 = 0.0#

                ReDim tenorsCashFlows(auxPeriods), cashFlows(auxPeriods)

                valueDate = startDate

                For j = 0 To auxPeriods
                    tenorsCashFlows(j) = DateDiff(DateInterval.Day, startDate, schedule(j, 2))
                    If cmsflag = 1 Then
                        cashFlows(j) = couponCamara(startDate, schedule(j, 0), schedule(j, 1), lastFixing, arraySwapRates(i - 1), _
                                                  swapType, y0, tenors, rates, 365, 2) + Indicador(auxPeriods, j) * y0
                    Else
                        cashFlows(j) = couponFloat(valueDate, schedule(j, 0), schedule(j, 1), floatIndexTenor, fixingLag, _
                                                   floatIndexShift, lastFixing, arraySwapRates(i - 1), swapType, y0, tenors, _
                                                   rates, basis, compound) + Indicador(auxPeriods, j) * y0
                    End If
                Next j

                While Math.Abs(diff) > 0.000001
                    auxRates(i + cuantosDepos - 1) = aux1 - diff / dp
                    p = pvCashFlowsFromCurve(tenorsCashFlows, cashFlows, auxTenors, auxRates, 365, 2)
                    aux1 = auxRates(i + cuantosDepos - 1)
                    auxRates(i + cuantosDepos - 1) = auxRates(i + cuantosDepos - 1) + 0.0001
                    auxVal = pvCashFlowsFromCurve(tenorsCashFlows, cashFlows, auxTenors, auxRates, 365, 2)
                    dp = (auxVal - p) / 0.0001
                    diff = p - y0
                    auxRates(i + cuantosDepos - 1) = aux1
                End While
            Next i

            ReDim final(cuantosDepos + cuantosSwaps - 1, 1)
            For i = 0 To cuantosDepos + cuantosSwaps - 1
                final(i, 0) = auxTenors(i)
                final(i, 1) = auxRates(i)
            Next i
        Catch ex As Exception
            final = Nothing
        Finally
        End Try
        Return final
    End Function

#End Region
#Region "Spread TAB"

    <ExcelFunction(category:=Categoria, Description:="Estima parÃ¡metro sigma del modelo CIR")> _
    Shared Function Sigma(<ExcelArgument(Description:="Serie histÃ³rica de Spread")> _
                          ByVal Spread() As Double, _
                          <ExcelArgument(Description:="Basis")> _
                          ByVal d As Integer) As Double

        Dim Result As Double = 0
        Dim i As Integer
        Dim delta As Double
        Dim SumN As Double = 0
        Dim SumD As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        delta = 1 / d

        For i = 0 To Spread.GetUpperBound(0) - 1
            SumN = SumN + (Spread(i + 1) - Spread(i)) ^ 2
            SumD = SumD + Spread(i + 1)
        Next i

        Result = (SumN / (SumD * (delta))) ^ (0.5)

        Return Result

    End Function
    <ExcelFunction(category:=Categoria, Description:="Estima parÃ¡metros alpha y beta del modelo CIR")> _
        Shared Function CIREstimationFuction(<ExcelArgument(Description:="Serie histÃ³rica de spreads")> _
                                             ByVal Spread() As Double, _
                                             <ExcelArgument(Description:="alpha")> _
                                             ByVal a As Double, _
                                             <ExcelArgument(Description:="beta")> _
                                             ByVal b As Double, _
                                             <ExcelArgument(Description:="sigma")> _
                                             ByVal S As Double, _
                                             <ExcelArgument(Description:="basis")> _
                                             ByVal d As Integer) As Double

        Dim Result As Double = 0
        Dim i As Integer
        Dim j As Integer
        Dim delta As Double
        Dim p1 As Double
        Dim p2 As Double

        Dim F() As Double
        Dim F1() As Double
        Dim F2() As Double
        Dim De() As Double
        Dim V1() As Double
        Dim V2() As Double

        j = Spread.GetUpperBound(0)

        ReDim F(j)
        ReDim F1(j)
        ReDim F2(j)
        ReDim De(j)
        ReDim V1(j)
        ReDim V2(j)

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        delta = 1 / d

        For i = 0 To j - 1
            F(i) = -a / b * (1 - Math.Exp(b * delta)) + Spread(i) * Math.Exp(b * delta)
            F1(i) = -1 / b * (1 - Math.Exp(b * delta))
            F2(i) = a / b ^ 2 * (1 - Math.Exp(b * delta)) + a * delta / b * Math.Exp(b * delta) + Spread(i) * delta * Math.Exp(b * delta)

            De(i) = (a * S ^ 2) / (2 * b ^ 2) - S ^ 2 / b * (a / b + Spread(i)) * Math.Exp(b * delta) + S ^ 2 / b * (a / (2 * b) + Spread(i)) * Math.Exp(2 * b * delta)

            V1(i) = F1(i) / De(i) * (Spread(i + 1) - F(i))
            p1 = p1 + V1(i)
            V2(i) = F2(i) / De(i) * (Spread(i + 1) - F(i))
            p2 = p2 + V2(i)

        Next i

        Result = p1 ^ 2 + p2 ^ 2

        Return Result

    End Function
    <ExcelFunction(category:=Categoria, Description:="Simula trayectorias del modelo CIR utilizando esquema de Milstein")> _
                Shared Function MilsteinCIR(<ExcelArgument(Description:="Spread inicial")> _
                                             ByVal SpreadIni As Double, _
                                             <ExcelArgument(Description:="alpha")> _
                                             ByVal a As Double, _
                                             <ExcelArgument(Description:="beta")> _
                                             ByVal b As Double, _
                                             <ExcelArgument(Description:="sigma")> _
                                             ByVal s As Double, _
                                             <ExcelArgument(Description:="basis")> _
                                             ByVal d As Double, _
                                             <ExcelArgument(Description:="delta")> _
                                             ByVal n As Double, _
                                             <ExcelArgument(Description:="NÃºmero de dÃ­as a simular")> _
                                             ByVal m As Double, _
                                             <ExcelArgument(Description:="NÃºmero de perÃ­odos")> _
                                             ByVal p As Integer) As Double(,)

        Dim Y(,) As Double
        Dim K(,) As Double

        Dim i As Integer
        Dim j As Integer

        Dim delta As Double
        Dim adelta As Double
        Dim bdelta As Double
        Dim Sqrdelta As Double
        Dim sSqrdelta As Double
        Dim s2delta As Double

        Dim x As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        If m Mod 2 <> 0 Then
            m = m + 1
        End If

        ReDim Y(0 To m - 1, 0 To n)
        ReDim K(0 To m - 1, 0 To p)

        delta = 1 / d
        adelta = a * delta
        bdelta = b * delta
        Sqrdelta = Math.Sqrt(delta)
        sSqrdelta = s * Sqrdelta
        s2delta = s ^ 2 / 4 * delta

        For i = 0 To m - 2 Step 2
            Y(i, 0) = SpreadIni
            Y(i + 1, 0) = SpreadIni
            For j = 0 To n - 1
                'x = xNORMAL(0, 1)
                'x = 0.001
                Randomize()
                x = ICND(Rnd)
                Y(i, j + 1) = Y(i, j) + adelta + bdelta * Y(i, j) + sSqrdelta * Math.Sqrt(Y(i, j)) * x + s2delta * (x ^ 2 - 1)
                Y(i + 1, j + 1) = Y(i + 1, j) + adelta + bdelta * Y(i + 1, j) + sSqrdelta * Math.Sqrt(Y(i + 1, j)) * (-x) + s2delta * ((-x) ^ 2 - 1)
            Next j
        Next i

        For i = 0 To m - 1
            K(i, 0) = Y(i, 0)
            K(i, p) = Y(i, n)
        Next i

        For j = 1 To p - 1
            For i = 0 To m - 1
                K(i, j) = Y(i, (n / p) * j)
            Next i
        Next j

        Return K

    End Function
    Private Shared Function ICND(ByVal u As Double) As Double

        'Attribute(ICND.VB_ProcData.VB_Invoke_Func = " \n14")
        'Inverse of the Cumulative Normal Distribution
        'Beasley&Springer

        Dim a(0 To 4) As Double
        Dim b(0 To 4) As Double
        Dim c(0 To 9) As Double
        Dim x As Double
        Dim r As Double

        a(1) = 2.50662823884
        a(2) = -18.61500062529
        a(3) = 41.39119773534
        a(4) = -25.44106049637
        b(1) = -8.4735109309
        b(2) = 23.08336743743
        b(3) = -21.06224101826
        b(4) = 3.13082909833
        c(1) = 0.337475482272615
        c(2) = 0.976169019091719
        c(3) = 0.160797971491821
        c(4) = 0.0276438810333863
        c(5) = 0.0038405729373609
        c(6) = 0.0003951896511919
        c(7) = 0.0000321767881768
        c(8) = 0.0000002888167364
        c(9) = 0.0000003960315187


        x = u - 0.5

        If Math.Abs(x) <= 0.42 Then

            r = x ^ 2

            r = x * (a(1) + r * (a(2) + r * (a(3) + a(4) * r))) / (1 + r * (b(1) + r * (b(2) + r * (b(3) + r * b(4)))))

        Else

            r = u

            If x > 0 Then r = 1 - u

            r = Math.Log(-Math.Log(r))

            r = c(1) + r * (c(2) + r * (c(3) + r * (c(4) + r * (c(5) + r * (c(6) + r * (c(7) + r * (c(8) + r * c(9))))))))

            If x < 0 Then r = -r

        End If

        Return r

    End Function
    Private Shared Function xNORMAL(ByVal mu As Double, ByVal s As Double) As Double

        Dim NORMAL01 As Double
        Const Pi As Double = 3.14159265358979
        Randomize()
        NORMAL01 = Math.Sqrt((-2 * LN(Rnd))) * Math.Sin(2 * Pi * Rnd())
        xNORMAL = mu + s * NORMAL01

        Return xNORMAL

    End Function
    Private Shared Function LN(ByVal x As Double) As Double

        LN = Math.Log(x) / Math.Log(Math.Exp(1))
        Return LN

    End Function

#End Region
#Region "PVB"
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de plain vanilla bond dada una curva (utiliza la función 'calendar', 'GetDiscountFactorFromCurve', 'Indicador', 'interest')")> _
    Shared Function PVB_PresentValueFromCurve1(<ExcelArgument(Description:="Fecha de valoracización")> _
                                            ByVal valueDate As Date, _
                                            <ExcelArgument(Description:="Fecha inicio")> _
                                            ByVal startDate As Date, _
                                            <ExcelArgument(Description:="Fecha fin")> _
                                            ByVal endDate As Date, _
                                            <ExcelArgument(Description:="Periodicidad")> _
                                            ByVal periodicity As Integer, _
                                            <ExcelArgument(Description:="Ajuste fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> _
                                            ByVal adjustment As Integer, _
                                            <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> _
                                            ByVal typeStubPeriod As Integer, _
                                            <ExcelArgument(Description:="Cupón de bono")> _
                                            ByVal rate As Double, _
                                            <ExcelArgument(Description:="Tipo capitalización y convención conteo días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360) (por defecto comp)")> _
                                            ByVal typeOfRate As String, _
                                            <ExcelArgument(Description:="nocional")> _
                                            ByVal notional As Double, _
                                            <ExcelArgument(Description:="Tenors en los que se observa tasa refencia")> _
                                            ByVal tenors() As Double, _
                                            <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                            ByVal rates() As Double, _
                                            <ExcelArgument(Description:="Base tasa referencia (360, 365)")> _
                                            ByVal basis As Double, _
                                            <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                            ByVal compound As Integer) As Double
        Dim schedule(,) As Object
        Dim numberOfPeriods As Integer
        Dim aux As Double
        Dim auxDate As Date
        Dim i As Integer
        Dim Result As Double = 0
        Dim pmtAdjustment As Integer = 0 'VALIDACION:

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If endDate <= startDate Then
                REM VALIDACION: PVB_PresentValueFromCurve1 = "endDate<=startDate" : Exit Function
                Result = Nothing
            Else
                schedule = calendar(startDate, endDate, periodicity, adjustment, pmtAdjustment, typeStubPeriod)
                numberOfPeriods = UBound(schedule, 1)
                aux = 0
                auxDate = endDate
                i = numberOfPeriods
                While valueDate < auxDate
                    aux = aux + (interest(schedule(i, 1), schedule(i, 2), rate, typeOfRate, notional) + notional * Indicador(schedule(i, 2), endDate)) _
                    * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, schedule(i, 2), valueDate), tenors, rates, basis, compound)
                    i = i - 1
                    If i < 1 Then Exit While
                    auxDate = schedule(i, 2)
                End While
                PVB_PresentValueFromCurve1 = aux
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de plain vanilla bond dada una curva (utiliza 'GetDiscountFactorFromCurve', 'Indicador', 'YearFraction') ")> _
        Shared Function PVB_PvFromCurve(<ExcelArgument(Description:="Fecha de valoracización")> _
                                        ByVal Fecha_Val As Date, _
                                        <ExcelArgument(Description:="Nocional")> _
                                        ByVal Nocional As Double, _
                                        <ExcelArgument(Description:="Cupón del bono")> _
                                        ByVal Tasa As Double, _
                                        <ExcelArgument(Description:="Fecha inicio")> _
                                        ByVal FechaIni As Date, _
                                        <ExcelArgument(Description:="Fecha fin")> _
                                        ByVal FechaMat As Date, _
                                        <ExcelArgument(Description:="Act/365, act/360 ó 30/360")> _
                                        ByVal Forma As String, _
                                        <ExcelArgument(Description:="periocidad")> _
                                        ByVal Per As Double, _
                                        <ExcelArgument(Description:="Tenors en los que se observa tasa refencia")> _
                                        ByVal tenors() As Double, _
                                        <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                        ByVal rates() As Double, _
                                        <ExcelArgument(Description:="Base de la tasa de referencia (360, 365)")> _
                                        ByVal basis As Double, _
                                        <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                        ByVal compound As Integer) As Double
        Dim nPagos As Integer
        Dim nPagos1 As Integer
        Dim fechapago() As Date
        Dim flujopago() As Double
        Dim Nominal() As Double
        Dim proxpago As Integer
        Dim i As Integer
        Dim fechaaux As Date
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If Fecha_Val > FechaMat Then
                Result = 0
            Else
                nPagos = CInt(((Year(FechaMat) - Year(FechaIni)) + 1) * 12 / Per)
                ReDim fechapago(nPagos)
                ReDim flujopago(nPagos)
                ReDim Nominal(nPagos)
                proxpago = 0
                'Definir fechas de pago y determinar la próxima vigente
                fechapago(0) = FechaIni
                i = 1
                While fechaaux < FechaMat
                    fechapago(i) = BussDay(DateAdd(DateInterval.Month, i * Per, FechaIni))
                    If Fecha_Val > fechapago(i) Then proxpago = i
                    fechaaux = fechapago(i)
                    i = i + 1
                End While
                proxpago = proxpago + 1
                nPagos1 = i - 1
                'Determinar los pagos del bono
                For i = proxpago To nPagos1
                    flujopago(i) = Nocional * Tasa * YearFraction(fechapago(i - 1), fechapago(i), Forma) + Nocional * Indicador(fechapago(nPagos1), fechapago(i))
                Next i
                'Calcular valor presente de los flujos
                Result = 0
                For i = proxpago To nPagos1
                    Result += +flujopago(i) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, fechapago(i), Fecha_Val), tenors, rates, basis, compound)
                Next i
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula valor presente de plain vanilla bond dada una TIR (utiliza 'YearFraction')")> _
    Shared Function PVB_PvFromTir(<ExcelArgument(Description:="Fecha de valoracización")> _
                                  ByVal Fecha_Val As Date, _
                                  <ExcelArgument(Description:="nocional")> _
                                  ByVal Nocional As Double, _
                                  <ExcelArgument(Description:="cupón del bono")> _
                                  ByVal Tasa As Double, _
                                  <ExcelArgument(Description:="fecha inicio")> _
                                  ByVal FechaIni As Date, _
                                  <ExcelArgument(Description:="fecha fin")> _
                                  ByVal FechaMat As Date, _
                                  <ExcelArgument(Description:="Act/365, act/360 ó 30/360")> _
                                  ByVal Forma As String, _
                                  <ExcelArgument(Description:="periociadad")> _
                                  ByVal Per As Integer, _
                                  <ExcelArgument(Description:="TIR a la que se descuenta bono")> _
                                  ByVal Tir As Double, _
                                  <ExcelArgument(Description:="Act/365, act/360 ó 30/360")> _
                                  ByVal FormaTir As String) As Double
        Dim npagos As Integer
        Dim npagos1 As Integer
        Dim fechapago() As Date
        Dim flujoPago() As Double
        Dim Nominal() As Double
        Dim proxpago As Integer
        Dim i As Integer
        Dim fechaaux As Date
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If Fecha_Val > FechaMat Then PVB_PvFromTir = 0 : Exit Function
            npagos = CInt(((Year(FechaMat) - Year(FechaIni)) + 1) * 12 / Per)
            ReDim fechapago(npagos)
            ReDim flujoPago(npagos)
            ReDim Nominal(npagos)
            proxpago = 0
            'Definir fechas de pago y determinar la próxima vigente
            fechapago(0) = FechaIni
            i = 1
            While fechaaux < FechaMat
                fechapago(i) = DateAdd(DateInterval.Month, i * Per, FechaIni)
                If Fecha_Val > fechapago(i) Then proxpago = i
                fechaaux = fechapago(i)
                i = i + 1
            End While
            proxpago = proxpago + 1
            npagos1 = i - 1

            'Determinar los pagos del bono
            For i = proxpago To npagos1
                flujoPago(i) = Nocional * Tasa * YearFraction(fechapago(i - 1), fechapago(i), Forma) + Nocional * Indicador(fechapago(npagos1), fechapago(i))
            Next i

            'Calcular valor presente de los flujos
            Result = 0
            For i = proxpago To npagos1
                Result += flujoPago(i) / (1 + Tir) ^ (YearFraction(Fecha_Val, fechapago(i), FormaTir))
            Next i
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Encuentra cupón que hace que bono valga par(utiliza PVB_pvFromCurve)")> _
        Shared Function PVB_FindCouponFromCurve(<ExcelArgument(Description:="fecha valorización")> _
                                                ByVal Fecha_Val As Date, _
                                                <ExcelArgument(Description:="fecha inicio")> _
                                                ByVal FechaIni As Date, _
                                                <ExcelArgument(Description:="fecha fin")> _
                                                ByVal FechaMat As Date, _
                                                <ExcelArgument(Description:="Act/365, act/360 ó 30/360")> _
                                                ByVal Forma As String, _
                                                <ExcelArgument(Description:="periocidad")> _
                                                ByVal Per As Integer, _
                                                <ExcelArgument(Description:="Tenors en los que se observa tasa refencia")> _
                                                ByVal tenors() As Double, _
                                                <ExcelArgument(Description:="Tasas a los correspondientes tenors")> _
                                                ByVal rates() As Double, _
                                                <ExcelArgument(Description:="Base tasa referencia (360, 365)")> _
                                                ByVal basis As Double, _
                                                <ExcelArgument(Description:="Lineal:1, Compuesta:2, Compuesta continua:3 (por defecto 2)")> _
                                                ByVal compound As Integer) As Double
        Dim cupon As Double
        Dim diff As Double
        Dim dp As Double
        Dim p As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If Fecha_Val > FechaMat Then
                Result = 0
            Else
                cupon = 0.02
                diff = 1
                dp = 1.0E+20
                While Math.Abs(diff) > 0.000001
                    cupon = cupon - diff / dp
                    p = PVB_PvFromCurve(Fecha_Val, 100, cupon, FechaIni, FechaMat, Forma, Per, tenors, rates, basis, compound)
                    dp = (PVB_PvFromCurve(Fecha_Val, 100, cupon + 0.01, FechaIni, FechaMat, Forma, Per, tenors, rates, basis, compound) - p) / 0.01
                    diff = p - 100
                End While
                PVB_FindCouponFromCurve = cupon
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
#End Region
#Region "FRB"
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function FRB_PvFromCurve(ByVal Fecha_Val As Date, _
                                 ByVal Nocional As Double, _
                                 ByVal tasavig As Double, _
                                 ByVal FechaIni As Date, _
                                 ByVal FechaMat As Date, _
                                 ByVal spread As Double, _
                                 ByVal Forma As String, _
                                 ByVal Per As Integer, _
                                 ByVal tenors1() As Double, _
                                 ByVal rates1() As Double, _
                                 ByVal basis1 As Double, _
                                 ByVal compound1 As Integer, _
                                 Optional ByVal tenors2() As Double = Nothing, _
                                 Optional ByVal rates2() As Double = Nothing, _
                                 Optional ByVal basis2 As Double = Nothing, _
                                 Optional ByVal compound2 As Integer = Nothing) As Double
        Dim nPagos As Integer
        Dim nPagos1 As Integer
        Dim fechapago() As Date
        Dim flujoPago() As Double
        Dim Nominal() As Double
        Dim proxpago As Integer
        Dim i As Integer
        Dim fechaAux As Date
        Dim df As Double
        Dim tasa As Double
        Dim QueCurva As Boolean
        Dim FRB As Double
        Dim Result As Double = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If Fecha_Val > FechaMat Then
                Result = 0
            Else
                nPagos = CInt((Year(FechaMat) - Year(FechaIni)) * 12 / Per + 2)
                ReDim fechapago(nPagos)
                ReDim flujoPago(nPagos)
                ReDim Nominal(nPagos)
                proxpago = 0
                'Definir fechas de pago y determinar la próxima vigente
                fechapago(0) = FechaIni
                i = 1
                While fechaAux < FechaMat
                    fechapago(i) = BussDay(DateAdd(DateInterval.Month, i * Per, FechaIni))
                    If Fecha_Val >= fechapago(i) Then proxpago = i
                    fechaAux = fechapago(i)
                    i = i + 1
                End While
                proxpago = proxpago + 1
                nPagos1 = i - 1

                'Determinar los pagos del bono
                flujoPago(proxpago) = Nocional * (tasavig + spread) * YearFraction(fechapago(proxpago - 1), fechapago(proxpago), Forma)
                For i = proxpago + 1 To nPagos1
                    df = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, fechapago(i), Fecha_Val), tenors1, rates1, basis1, compound1) / GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, fechapago(i - 1), Fecha_Val), tenors1, rates1, basis1, compound1)
                    tasa = GetRateFromDiscountFactor_v2(DateDiff(DateInterval.Day, fechapago(i), fechapago(i - 1)), df, Forma, 1) + spread
                    flujoPago(i) = Nocional * tasa * YearFraction(fechapago(i - 1), fechapago(i), Forma) + Nocional * Indicador(fechapago(nPagos1), fechapago(i))
                Next i

                'Calcular valor presente de los flujos
                QueCurva = IsNothing(tenors1)
                FRB = 0
                Select Case QueCurva
                    Case True
                        For i = proxpago To nPagos1
                            FRB = FRB + flujoPago(i) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, fechapago(i), Fecha_Val), tenors1, rates1, basis1, compound1)
                        Next i
                    Case False
                        For i = proxpago To nPagos1
                            FRB = FRB + flujoPago(i) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, fechapago(i), Fecha_Val), tenors2, rates2, basis2, compound2)
                        Next i
                End Select
                Result = FRB
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function FRB_FindSpreadFromCurveAndBook(ByVal Fecha_Val As Date, _
                                    ByVal Book As Double, _
                                    ByVal Nocional As Double, _
                                    ByVal tasavig As Double, _
                                    ByVal FechaIni As Date, _
                                    ByVal FechaMat As Date, _
                                    ByVal Forma As String, _
                                    ByVal Per As Integer, _
                                    ByVal tenors1() As Double, _
                                    ByVal rates1() As Double, _
                                    ByVal basis1 As Double, _
                                    ByVal compound1 As Integer, _
                                    Optional ByVal tenors2() As Double = Nothing, _
                                    Optional ByVal rates2() As Double = Nothing, _
                                    Optional ByVal basis2 As Double = Nothing, _
                                    Optional ByVal Compund2 As Integer = Nothing) As Double
        Dim spread As Double
        Dim diff As Double
        Dim dp As Double
        Dim p As Double
        Dim Result As Double
        'Return the Dirty price of a bond paying Coupon (expresed as a percentage)
        'Freq times per year. BasisCount is the number of days of the year

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If Fecha_Val > FechaMat Then
                Result = 0
            Else
                spread = 0.02
                diff = 1
                dp = 1.0E+20
                While Math.Abs(diff) > 0.000001
                    spread = spread - diff / dp
                    p = FRB_PvFromCurve(Fecha_Val, Nocional, tasavig, FechaIni, FechaMat, spread, Forma, Per, tenors1, rates1, basis1, compound1, tenors2, rates2, basis2, Compund2)
                    dp = (FRB_PvFromCurve(Fecha_Val, Nocional, tasavig, FechaIni, FechaMat, spread + 0.01, Forma, Per, tenors1, rates1, basis1, compound1, tenors2, rates2, basis2, Compund2) - p) / 0.01
                    diff = p - Book
                End While
                Result = spread
            End If
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function AmortFRB_PvFromCurve(ByVal Fecha_Val As Date, _
                                  ByVal IniDates() As Date, _
                                  ByVal EndDates() As Date, _
                                  ByVal Nocionales() As Double, _
                                  ByVal tasavig As Double, _
                                  ByVal spread As Double, _
                                  ByVal Forma As String, _
                                  ByVal tenors1() As Double, _
                                  ByVal rates1() As Double, _
                                  ByVal basis1 As Double, _
                                  ByVal compound1 As Integer, _
                                  Optional ByVal tenors2() As Double = Nothing, _
                                  Optional ByVal rates2() As Double = Nothing, _
                                  Optional ByVal basis2 As Double = Nothing, _
                                  Optional ByVal compound2 As Integer = Nothing) As Double
        Dim QueCurva As Boolean
        Dim n As Integer
        Dim m As Integer
        Dim npagos As Integer
        Dim npagos1 As Integer
        Dim proxpago As Integer
        Dim flujopago() As Double
        Dim i As Integer
        Dim fechaaux As Date
        Dim df As Double
        Dim tasa As Double
        Dim FRB As Double
        Dim Result As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            n = IniDates.GetUpperBound(1)
            m = Nocionales.GetUpperBound(1)
            If n <> m Then AmortFRB_PvFromCurve = 0 : Exit Function
            If Fecha_Val > IniDates(n) Then AmortFRB_PvFromCurve = 0 : Exit Function
            npagos = n
            ReDim flujopago(npagos)
            proxpago = 0

            'Definir fechas de pago y determinar la próxima vigente
            i = 1
            fechaaux = EndDates(1)
            While fechaaux < EndDates(n)
                If Fecha_Val > EndDates(i) Then proxpago = i
                fechaaux = EndDates(i)
                i = i + 1
            End While
            proxpago = proxpago + 1
            npagos1 = i - 1

            'Determinar los pagos del bono
            flujopago(proxpago) = Nocionales(proxpago) * (tasavig + spread) * YearFraction(IniDates(proxpago), EndDates(proxpago), Forma) + Nocionales(proxpago - 1) - Nocionales(proxpago)
            For i = proxpago + 1 To npagos1
                df = GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, EndDates(i), Fecha_Val), tenors1, rates1, basis1, compound1) / _
                     GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, IniDates(i), Fecha_Val), tenors1, rates1, basis1, compound1)
                tasa = GetRateFromDiscountFactor(IniDates(i), EndDates(i), df, Forma) + spread
                flujopago(i) = Nocionales(i) * tasa * YearFraction(IniDates(i), EndDates(i), Forma) + Nocionales(i) - Nocionales(i + 1)
            Next i

            'Calcular valor presente de los flujos
            QueCurva = IsNothing(tenors2)
            FRB = 0
            Select Case QueCurva

                Case True
                    For i = proxpago To npagos1
                        FRB = FRB + flujopago(i) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, EndDates(i), Fecha_Val), tenors1, rates1, basis1, compound1)
                    Next i
                    AmortFRB_PvFromCurve = FRB

                Case False
                    For i = proxpago To npagos1
                        FRB = FRB + flujopago(i) * GetDiscountFactorFromCurve(DateDiff(DateInterval.Day, EndDates(i), Fecha_Val), tenors2, rates2, basis2, compound2)
                    Next i
                    AmortFRB_PvFromCurve = FRB

            End Select
        Catch ex As Exception
            Result = Nothing
        Finally
        End Try

    End Function
#End Region
#Region "AUX"
    <ExcelFunction(category:=Categoria, Description:="Devuelve interes devengado entre dos fecha para tasa y nocional dados (utiliza 'YearFraction')")> _
    Shared Function interest(<ExcelArgument(Description:="Fecha inicio")> ByVal startDate As Date, _
                                    <ExcelArgument(Description:="Fecha fin")> ByVal endDate As Date, _
                                    <ExcelArgument(Description:="Tasa")> ByVal rate As Double, _
                                    <ExcelArgument(Description:="Tipo capitalización y convención conteo de días (Lin act/365, Lin act/360, Lin 30/360, Comp act/365, Comp act/360, Comp 30/360) (por defecto comp)")> ByVal typeOfRate As String, _
                                    <ExcelArgument(Description:="Nocional")> ByVal notional As Double) As Double
        Dim aux As Integer
        Dim nInterest As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            aux = Len(typeOfRate) - 4
            If typeOfRate.Substring(0, 3).ToLower = "lin" Then
                nInterest = notional * rate * YearFraction(startDate, endDate, typeOfRate.Substring(4, aux))
            Else
                nInterest = notional * ((1 + rate) ^ YearFraction(startDate, endDate, typeOfRate.Substring(4, aux)) - 1)
            End If
        Catch ex As Exception
            nInterest = 0
        Finally
        End Try
        Return nInterest
    End Function
    <ExcelFunction(category:=Categoria, Description:="Devuele 1 si n=m, si no 0")> _
    Shared Function Indicador(<ExcelArgument(Description:="n")> ByVal n As Object, _
                                     <ExcelArgument(Description:="m")> ByVal m As Object) As Integer
        Dim cell1 As Microsoft.Office.Interop.Excel.Range
        Dim cell2 As Microsoft.Office.Interop.Excel.Range
        Dim result As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            If IsNumeric(n) Then
                result = CInt(IIf(CDbl(n) = CDbl(m), 1, 0))
            ElseIf IsDate(n) Then
                result = CInt(IIf(CDate(n) = CDate(m), 1, 0))
            Else
                cell1 = CType(n, Microsoft.Office.Interop.Excel.Range)
                cell2 = CType(m, Microsoft.Office.Interop.Excel.Range)
                If Object.Equals(cell1.Value2, cell2.Value2) Then
                    result = 1
                Else
                    result = 0
                End If
            End If
        Catch ex As Exception
            result = Nothing
        Finally
        End Try
        Return result
    End Function
    Function valad(ByVal x As String) As Integer
        Return CInt(Val(x))
    End Function
    <ExcelFunction(category:=Categoria, Description:="...")> _
    Shared Function Ten(ByVal Tn As String) As Integer
        Dim cTen As String
        Dim nTen As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            cTen = Tn
            If UCase(Tn.Substring(Tn.Length - 1)) = "C" Then
                cTen = Tn.Substring(0, Tn.Length - 1)
            End If
            If cTen = "Spot" Then
                nTen = 0
            ElseIf UCase(cTen.Substring(cTen.Length - 1)) = "D" Then
                nTen = 1
            ElseIf UCase(cTen.Substring(cTen.Length - 1)) = "W" Then
                nTen = CInt(Val(cTen) * 7)
            Else
                If Right(cTen, 1).ToUpper = "M" Then nTen = 1 Else nTen = 12
                nTen = CInt(Val(cTen) * nTen)
            End If
        Catch ex As Exception
        Finally
        End Try
        Return nTen
    End Function
    <ExcelFunction(category:=Categoria, Description:="Dada una fecha y tenor, entrega fecha del tenor local, es decir, finalizando el día 9 del mes correspondiente(utiliza 'AddMonths' y 'Ten')")> _
        Shared Function localTenor(<ExcelArgument(Description:="fecha inicial")> _
                                   ByVal startDate As Date, _
                                   <ExcelArgument(Description:="TENOR")> _
                                   ByVal Tn As String) As Date
        Dim dLocalTenor As Date = Nothing
        Dim aux As Integer
        Dim aux1 As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            aux = Ten(Tn)
            aux1 = AddMonths(startDate, aux)
            dLocalTenor = DateSerial(Year(aux1), Month(aux1), 9)
        Catch ex As Exception
        Finally
        End Try
        Return dLocalTenor
    End Function
    <ExcelFunction(category:=Categoria, Description:="Ajusta al hábil siguiente (si fecha ingresada es hábil entonces devuelve tal fecha)")> _
    Shared Function BussDay(<ExcelArgument(description:="Fecha")> ByVal a As Date) As Date
        Dim dBussDay As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Select Case Weekday(a)
                Case 1
                    dBussDay = a.AddDays(1)
                Case 7
                    dBussDay = a.AddDays(2)
                Case Else
                    dBussDay = a
            End Select
        Catch ex As Exception
            dBussDay = Nothing
        Finally
        End Try
        Return dBussDay
    End Function
    <ExcelFunction(category:=Categoria, Description:="Quita n días hábiles a una fecha dada. La opción 0 quita un día habíl si fecha ingresada es inhábil")> _
    Shared Function lag(<ExcelArgument(Description:="Fecha")> _
                        ByVal a As Date, _
                        <ExcelArgument(Description:="Número de días hábiles")> _
                        ByVal nStep As Integer) As Date
        Dim Ret As Date
        Dim i As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Ret = a
            Select Case nStep
                Case 0
                    Ret = Prev2(a)
                Case Else
                    For i = 1 To nStep
                        Ret = Prev2(Ret.AddDays(-1))
                    Next
            End Select
        Catch ex As Exception
            Ret = Nothing
        Finally
        End Try
        Return Ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Agrega n días hábiles a una fecha dada. La opción 0 agrega un día habíl si fecha ingresada es inhábil")> _
    Shared Function shift(<ExcelArgument(Description:="Fecha")> _
                                 ByVal a As Date, _
                                 <ExcelArgument(Description:="Número de días hábiles")> _
                                 ByVal nStep As Integer) As Date
        Dim ret As Date
        Dim i As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ret = a
            Select Case nStep
                Case 0
                    ret = BussDay(a)
                Case Else
                    For i = 1 To nStep
                        ret = BussDay(ret.AddDays(1))
                    Next
            End Select
        Catch ex As Exception
            ret = Nothing
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Ajusta a hábil siguiente. Si se cambia mes entonces ajusta al hábil anterior. (si la fecha ingresada es hábil entonces devuelve tal fecha)")> _
    Shared Function ModBussDay(<ExcelArgument(Description:="Fecha")> _
                        ByVal a As Date) As Date
        Dim m1 As Integer = 0
        Dim m2 As Integer = 0
        Dim b As Date
        Dim w As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            m1 = Month(a)
            b = BussDay(a)
            m2 = Month(b)
            If m2 <> m1 Then
                b = b.AddDays(-1)
                w = Weekday(b)
                m2 = Month(b)
                While (w = 1 Or w = 7) Or m2 <> m1
                    b = b.AddDays(-1)
                    w = Weekday(b)
                    m2 = Month(b)
                End While
            End If
        Catch ex As Exception
        Finally
        End Try
        Return b
    End Function
    <ExcelFunction(category:=Categoria, Description:="Si la fecha ingresada no es hábil, la ajusta a hábil anterior")> _
    Shared Function Prev2(<ExcelArgument(Description:="Fecha ingresada")> ByVal a As Date) As Date
        Dim dPrev2 As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Select Case Weekday(a)
                Case 1
                    dPrev2 = a.AddDays(-2)
                Case 7
                    dPrev2 = a.AddDays(-1)
                Case Else
                    dPrev2 = a
            End Select
        Catch ex As Exception
        Finally
        End Try
        Return dPrev2
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el máximo entre 'a'y 'b'")> _
    Shared Function Max(<ExcelArgument(Description:="a")> ByVal a As Integer, _
                        <ExcelArgument(Description:="b")> ByVal b As Integer) As Integer
        Dim ret As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ret = a
            If b > a Then ret = b
        Catch ex As Exception
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula el mínimo entre 'a'y 'b'")> _
    Shared Function Min(<ExcelArgument(Description:="a")> ByVal a As Integer, _
                        <ExcelArgument(Description:="b")> ByVal b As Integer) As Integer
        Dim Ret As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Ret = a
            If b < a Then Ret = b
        Catch ex As Exception
        Finally
        End Try
        Return Ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula tiempo en días entre dos fechas utilizando base especificada")> _
    Shared Function CountDays(<ExcelArgument(Description:="Fecha inicio")> _
                              ByVal t1 As Date, _
                              <ExcelArgument(Description:="Fecha fin")> _
                              ByVal t2 As Date, _
                              <ExcelArgument(Description:="base: Act/365, act/360 ó 30/360. Notar que act/365=act/360 ya que cuenta en días (por defecto act/365=act/360)")> _
                              ByVal basis As String) As Double
        Dim ret As Long
        Dim d1 As Integer
        Dim d2 As Integer
        Dim m1 As Integer
        Dim m2 As Integer
        Dim y1 As Integer
        Dim y2 As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Select Case basis.Trim.ToLower
                Case "act/360", "act/365"
                    ret = DateDiff(DateInterval.Day, t2, t1)
                Case "30/360", "30/365"
                    d1 = t1.Day
                    d2 = t2.Day
                    m1 = t1.Month
                    m2 = t2.Month
                    y1 = t1.Year
                    y2 = t2.Year
                    If d1 = 31 Then d1 = 30
                    If d2 = 31 And d1 = 30 Then d2 = 30
                    ret = (d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1)
                    'CountDays = (Max(0, 30 - Day(t1)) + Min(30, Day(t2)) + 30 * (Month(t2) - Month(t1) - 1) + 360 * (Year(t2) - Year(t1)))
                Case Else   'Default
                    ret = DateDiff(DateInterval.Day, t2, t1)
            End Select
        Catch ex As Exception
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula tiempo en años entre dos fechas utilizando base especificada")> _
    Shared Function YearFraction(<ExcelArgument(Description:="Fecha inicio")> _
                                        ByVal t1 As Date, _
                                        <ExcelArgument(Description:="Fecha fin")> _
                                        ByVal t2 As Date, _
                                        <ExcelArgument(Description:="base: Act/365, act/360 ó 30/360 (por defecto Act/365)")> _
                                        ByVal basis As String) As Double
        Dim ret As Double = 0
        Dim d1 As Integer = 0
        Dim d2 As Integer = 0
        Dim m1 As Integer = 0
        Dim m2 As Integer = 0
        Dim y1 As Integer = 0
        Dim y2 As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Select Case basis.ToLower.Trim
                Case "act/360"
                    ret = DateDiff(DateInterval.Day, t1, t2) / (3.6 * 100)
                Case "act/365"
                    ret = DateDiff(DateInterval.Day, t1, t2) / (3.65 * 100)
                Case "30/360"
                    d1 = t1.Day
                    d2 = t2.Day
                    m1 = t1.Month
                    m2 = t2.Month
                    y1 = t1.Year
                    y2 = t2.Year
                    If d1 = 31 Then d1 = 30
                    If d2 = 31 And d1 = 30 Then d2 = 30
                    ret = ((d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1)) / (3.6 * 100)
                Case Else   'Default
                    ret = DateDiff(DateInterval.Day, t1, t2) / (3.65 * 100)
            End Select
        Catch ex As Exception
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Agrega o quita meses a una fecha dada")> _
    Shared Function AddMonths(<ExcelArgument(Description:="Fecha inicial")> _
                              ByVal Dat As Date, _
                              <ExcelArgument(Description:="Número de meses a agregar o quitar")> _
                              ByVal Mon As Integer, _
                              <ExcelArgument(Description:="Fecha inicial")> _
                              Optional ByVal Metodo As String = "Following") As Date
        Dim dm(12) As Integer
        Dim D As Integer = 0
        Dim m As Integer = 0
        Dim y As Integer = 0
        Dim eom As Boolean
        Dim ret As Date = Nothing

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            dm(1) = 31
            dm(2) = 28
            dm(3) = 31
            dm(4) = 30
            dm(5) = 31
            dm(6) = 30
            dm(7) = 31
            dm(8) = 31
            dm(9) = 30
            dm(10) = 31
            dm(11) = 30
            dm(12) = 31
            D = Dat.Day
            m = Dat.Month
            y = Dat.Year
            If D = dm(m) And m <> 2 Then eom = True 'Detect if the date correspond to the end of a month
            Mon = Mon + m
            m = Mon - CType(12 * Int(Mon / 12), Integer)
            If m = 0 Then m = 12
            y += CType(Int((Mon - 1) / 12), Integer)
            If y < 0 Then y = 1900
            If D > dm(m) Then eom = True
            If eom = True Then   'if the date correspond to the end of a month,
                D = dm(m)        'the resulting date is forced to be the end of a month
                If (Int(y / 4) = y / 4) And m = 2 Then D = 29
            End If
            ret = DateSerial(y, m, D)
        Catch ex As Exception
            ret = Nothing
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Agrega ó quita meses a una fecha dada. Fecha final es día 9 del mes")> _
    Shared Function AddMonthsC(<ExcelArgument(Description:="Fecha inicial")> _
                               ByVal Dat As Date, _
                              <ExcelArgument(Description:="Número de meses a agregar o quitar")> _
                               ByVal Mon As Integer) As Date
        'equivalent to Excel Function EDATE()
        'add/sustracts Mon number of months to date Dat for positive/negative values of Mon
        Dim dm(12) As Integer
        Dim D As Integer
        Dim m As Integer
        Dim y As Integer
        Dim eom As Boolean
        Dim ret As Date

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            dm(1) = 31
            dm(2) = 28
            dm(3) = 31
            dm(4) = 30
            dm(5) = 31
            dm(6) = 30
            dm(7) = 31
            dm(8) = 31
            dm(9) = 30
            dm(10) = 31
            dm(11) = 30
            dm(12) = 31
            D = Dat.Day
            m = Dat.Month
            y = Dat.Year
            If D = dm(m) And m <> 2 Then eom = True 'Detect if the date correspond to the end of a month
            Mon = Mon + m
            m = CInt(Mon - 12 * Int(Mon / 12))
            If m = 0 Then m = 12
            y = CInt(y + Int((Mon - 1) / 12))
            If y < 0 Then y = 1900
            If D > dm(m) Then eom = True

            If eom = True Then   'if the date correspond to the end of a month,
                D = dm(m)        'the resulting date is forced to be the end of a month
                If (Int(y / 4) = y / 4) And m = 2 Then D = 29
            End If
            ret = DateSerial(y, m, 9)
        Catch ex As Exception
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Calcula tiempo en años entre dos fechas utilizando base act/act")> _
    Shared Function YearFracActualActual(<ExcelArgument(Description:="Fecha inicio")> _
                                         ByVal inicial As Date, _
                                         <ExcelArgument(Description:="Fecha fin")> _
                                         ByVal final As Date) As Integer
        Dim aux1 As Date = Nothing
        Dim aux2 As Date = Nothing
        Dim año1 As Integer = 0
        Dim año2 As Integer = 0
        Dim dia_inicial As Integer
        Dim mes_inicial As Integer
        Dim año_inicial As Integer
        Dim dia_final As Integer
        Dim mes_final As Integer
        Dim año_final As Integer
        Dim ret As Integer = 0

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            dia_inicial = inicial.Day
            mes_inicial = inicial.Month
            año_inicial = inicial.Year
            dia_final = final.Day
            mes_final = final.Month
            año_final = final.Year
            If mes_final > mes_inicial Then
                aux1 = DateAdd("m", (año_final - año_inicial) * 12, inicial)
            ElseIf mes_final < mes_inicial Then
                aux1 = DateAdd("m", (año_final - año_inicial - 1) * 12, inicial)
            Else
                If dia_inicial <= dia_final Then
                    aux1 = DateAdd("m", (año_final - año_inicial) * 12, inicial)
                Else
                    aux1 = DateAdd("m", (año_final - año_inicial - 1) * 12, inicial)
                End If
            End If
            año1 = Year(aux1) - año_inicial
            aux2 = aux1.AddMonths(12)
            año2 = CInt(DateDiff(DateInterval.Year, final, aux1) / DateDiff(DateInterval.Year, aux2, aux1))
            ret = año1 + año2
        Catch ex As Exception
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Si parte decimal de n es menor que 0.25 devuelve n, si es mayor o igual a 0.25 y menor que 0.75 devuelve (n+0.5) y si es mayor que 0.75 devuelve (n+1) ")> _
    Shared Function RedondeoParcial(<ExcelArgument(Description:="n")> ByVal numero As Double) As Double
        Dim ret As Double = 0
        Dim entero As Integer
        Dim mantisa As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            entero = Int(numero)
            mantisa = numero - entero
            If mantisa >= 0 And mantisa < 0.25 Then
                ret = 0
            ElseIf mantisa >= 0.25 And mantisa < 0.75 Then
                ret = 0.5
            ElseIf mantisa >= 0.75 Then
                ret = 1
            End If
            ret = ret + entero
        Catch ex As Exception
        Finally
        End Try
        Return ret
    End Function
    <ExcelFunction(category:=Categoria, Description:="Genera calendario de pata variable de un Swap")> _
    Shared Function calendarAmortize(<ExcelArgument(Description:="Fecha inicio")> ByVal startDate As Date, _
                                    <ExcelArgument(Description:="Fecha fin")> ByVal endDate As Date, _
                                    <ExcelArgument(Description:="Periocidad")> ByVal periodicity As Double, _
                                    <ExcelArgument(Description:="Ajuste fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> ByVal accAdjustment As Double, _
                                    <ExcelArgument(Description:="Ajuste fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> ByVal pmtAdjustment As Double, _
                                    <ExcelArgument(Description:="0 corto inicio, 1 corto final, 2 largo inicio, 3 largo final")> ByVal typeStubPeriod As Double, _
                                    <ExcelArgument(Description:="0 bullet, 1 constante, 2 frances, else bullet")> ByVal typeAmortization As Double, _
                                    <ExcelArgument(Description:="tasa para frances")> ByVal amortizeRate As Double, _
                                    <ExcelArgument(Description:="tipo tasa frances (no implementada siempre comp 30/360)")> ByVal typeAmortizeRate As String)
        'periodicity = expressed in months, valid is integers from 1 to 12
        'accAdjustment = 0 is same, 1 is next buss day, else is next buss day
        'pmtAdjustment = 1 is next buss day, 2 is plus 2 buss days, else is next buss day,
        '0 is same or next buss day if end date is not buss day
        '
        '   New: typeStubPeriod = 1 is SHORT at the beginning, 2 is SHORT at the end,
        '   3 is SHORT at the beginning, 4 is LONG at the end, Else=1.
        '
        '   New: typeAmortization: 1 is bullet, 2 is constant, 3 is French, Else is 1
        '   If 3 then parameterAmortization is a rate.
        '
        '   Please note: in this version the variable typeAmortizeRate is required but not used.

        Dim aux(,) As Object
        Dim Result(,) As Object
        Dim auxDate As Date
        Dim j As Integer
        Dim i As Integer
        Dim numberOfPeriods As Integer
        Dim aux1, constantInstallment, constantAmo, auxInterest As Double

        REM periodicity es necesario, sino la función entra en loop.
        If periodicity = 0 Then
            Return Nothing
        End If

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Select Case typeStubPeriod
                Case 1
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 5)

                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j

                Case 2
                    auxDate = startDate
                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 5)

                    For j = 1 To numberOfPeriods
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        If AddMonths(aux(j, 1), periodicity) < endDate Then
                            aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                        Else
                            aux(j, 2) = endDate
                        End If
                    Next j

                Case 3
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 2
                    ReDim aux(numberOfPeriods, 5)

                    For j = numberOfPeriods To 2 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                    Next j

                    aux(1, 1) = startDate
                    aux(1, 2) = aux(2, 1)

                Case 4
                    auxDate = startDate
                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 2
                    ReDim aux(numberOfPeriods, 5)

                    For j = 1 To numberOfPeriods - 1
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                    Next j

                    aux(numberOfPeriods, 1) = aux(numberOfPeriods - 1, 2)
                    aux(numberOfPeriods, 2) = endDate

                Case Else
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 5)

                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j


            End Select

            'Here we apply accrual and payment adjustments

            If accAdjustment <> 0 Then
                For i = 1 To numberOfPeriods
                    aux(i, 1) = BussDay(aux(i, 1))
                    aux(i, 2) = BussDay(aux(i, 2))
                Next i
            End If

            Select Case pmtAdjustment
                Case 1
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2).AddDays(1))
                    Next i
                Case 2
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(BussDay(aux(i, 2).AddDays(1)).AddDays(1))
                    Next i
                Case Else
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2))
                    Next i
            End Select

            'Here we generate amortization percentages

            Select Case typeAmortization
                Case 1
                    For i = 1 To numberOfPeriods - 1
                        aux(i, 4) = 0
                        aux(i, 5) = 1
                    Next i
                    aux(numberOfPeriods, 4) = 1
                    aux(numberOfPeriods, 5) = 1

                Case 2
                    constantAmo = 1 / numberOfPeriods
                    For i = 1 To numberOfPeriods
                        If i = 1 Then
                            aux(i, 5) = 1
                        Else
                            aux(i, 5) = aux(i - 1, 5) - aux(i - 1, 4)
                        End If
                        aux(i, 4) = constantAmo
                    Next i

                Case 3
                    aux1 = 0.0
                    For i = 1 To numberOfPeriods
                        aux1 = aux1 + (1.0 + amortizeRate) ^ (-YearFraction(aux(1, 1), aux(i, 2), "30/360"))
                    Next i
                    constantInstallment = 1.0 / aux1
                    aux1 = 1.0
                    For i = 1 To numberOfPeriods
                        If i = 1 Then
                            aux(i, 5) = 1.0
                        Else
                            aux(i, 5) = aux(i - 1, 5) - aux(i - 1, 4)
                        End If
                        auxInterest = aux1 * ((1.0 + amortizeRate) ^ (YearFraction(aux(i, 1), aux(i, 2), "30/360")) - 1)
                        aux(i, 4) = constantInstallment - auxInterest
                        aux1 = aux1 - aux(i, 4)
                    Next i


                Case Else
                    For i = 1 To numberOfPeriods - 1
                        aux(i, 4) = 0
                        aux(i, 5) = 1
                    Next i
                    aux(numberOfPeriods, 4) = 1
                    aux(numberOfPeriods, 5) = 1

            End Select
            Result = DContraeObj(aux)

        Catch ex As Exception
            Result = Nothing
        End Try

        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Genera calendario de pata variable de un Swap")> _
        Shared Function calendar(<ExcelArgument(Description:="Fecha inicio")> ByVal startDate As Date, _
                                        <ExcelArgument(Description:="Fecha fin")> ByVal endDate As Date, _
                                        <ExcelArgument(Description:="Periocidad")> ByVal periodicity As Double, _
                                        <ExcelArgument(Description:="Ajuste fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> ByVal accAdjustment As Double, _
                                        <ExcelArgument(Description:="Ajuste fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> ByVal pmtAdjustment As Double, _
                                        <ExcelArgument(Description:="0 corto inicio, 1 corto final, 2 largo inicio, 3 largo final")> ByVal typeStubPeriod As Double)

        Dim aux(,) As Date
        Dim Result(,) As Object
        Dim auxDate As Date
        Dim j As Integer
        Dim i As Integer
        Dim numberOfPeriods As Integer

        REM periodicity es necesario, sino la función entra en loop.
        If periodicity = 0 Then
            Return Nothing
        End If


        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            Select Case typeStubPeriod
                Case 1
                    auxDate = endDate
                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 1
                    ReDim aux(0 To numberOfPeriods, 0 To 3)
                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j

                Case 2
                    auxDate = startDate
                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 1
                    ReDim aux(0 To numberOfPeriods, 0 To 3)
                    For j = 1 To numberOfPeriods
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        If AddMonths(aux(j, 1), periodicity) < endDate Then
                            aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                        Else
                            aux(j, 2) = endDate
                        End If
                    Next j

                Case 3
                    auxDate = endDate
                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 2
                    ReDim aux(0 To numberOfPeriods, 0 To 4)
                    For j = numberOfPeriods To 2 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                    Next j
                    aux(1, 1) = startDate
                    aux(1, 2) = aux(2, 1)

                Case 4
                    auxDate = startDate
                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 2
                    ReDim aux(0 To numberOfPeriods, 0 To 4)
                    For j = 1 To numberOfPeriods - 1
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                    Next j
                    aux(numberOfPeriods, 1) = aux(numberOfPeriods - 1, 2)
                    aux(numberOfPeriods, 2) = endDate

                Case Else
                    auxDate = endDate
                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 1
                    ReDim aux(0 To numberOfPeriods, 0 To 4)
                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j
            End Select

            If accAdjustment <> 0 Then
                For i = 1 To numberOfPeriods
                    aux(i, 1) = BussDay(aux(i, 1))
                    aux(i, 2) = BussDay(aux(i, 2))
                Next i
            End If

            Select Case pmtAdjustment
                Case 1
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2).AddDays(+1))
                    Next i
                Case 2
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(BussDay(aux(i, 2).AddDays(+1)).AddDays(+1))
                    Next i
                Case Else
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2))
                    Next i
            End Select

            Result = DContrae(aux)

        Catch ex As Exception
            Result = Nothing
        Finally
        End Try
        Return Result
    End Function
    <ExcelFunction(category:=Categoria, Description:="Genera calendario de pata variable de un Swap")> _
    Shared Function floatCalendarAmortize(<ExcelArgument(Description:="Fecha inicio")> ByVal startDate As Date, _
                                    <ExcelArgument(Description:="Fecha fin")> ByVal endDate As Date, _
                                    <ExcelArgument(Description:="Periocidad")> ByVal periodicity As Integer, _
                                    <ExcelArgument(Description:="Ajuste fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> ByVal accAdjustment As Integer, _
                                    <ExcelArgument(Description:="Ajuste fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> ByVal pmtAdjustment As Integer, _
                                    <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> ByVal typeStubPeriod As Double, _
                                    <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1, un día hábil; 2, dos días hábiles ")> ByVal fixingLag As Integer, _
                                    <ExcelArgument(Description:="Periocidad fijación: n, se fija cada n periodos")> ByVal fixingRatio As Integer, _
                                    <ExcelArgument(Description:="Periodo fijación única: 0, al final; 1, al principio")> ByVal fixingStubPeriod As Integer, _
                                    <ExcelArgument(Description:="1 bullet, 2 constante, 3 frances, else bullet")> ByVal typeAmortization As Double, _
                                    <ExcelArgument(Description:="Tasa para frances Comp 30/360")> ByVal amortizeRate As Double, _
                                    <ExcelArgument(Description:="No implementada")> ByVal typeAmortizeRate As String)


        Dim aux(,) As Object
        Dim auxDate As Date
        Dim j As Integer
        Dim i As Integer
        Dim numberOfPeriods As Integer
        Dim cociente As Integer
        Dim resto As Integer
        Dim result As Object(,)
        Dim aux1, constantAmo, constantInstallment As Double

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        REM periodicity es necesario, sino la función entra en loop.
        If periodicity = 0 Then
            Return Nothing
        End If

        Try
            Select Case typeStubPeriod
                Case 1
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 6)

                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j

                Case 2
                    auxDate = startDate

                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 6)

                    For j = 1 To numberOfPeriods
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        If AddMonths(aux(j, 1), periodicity) < endDate Then
                            aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                        Else
                            aux(j, 2) = endDate
                        End If
                    Next j

                Case 3
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 2
                    ReDim aux(numberOfPeriods, 6)

                    For j = numberOfPeriods To 2 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                    Next j

                    aux(1, 1) = startDate
                    aux(1, 2) = aux(2, 1)

                Case 4
                    auxDate = startDate

                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 2
                    ReDim aux(numberOfPeriods, 6)

                    For j = 1 To numberOfPeriods - 1
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                    Next j

                    aux(numberOfPeriods, 1) = aux(numberOfPeriods - 1, 2)
                    aux(numberOfPeriods, 2) = endDate


                Case Else
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 6)

                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j


            End Select

            'Here we generate accrual adjustments and payment adjustments

            If accAdjustment <> 0 Then
                For i = 1 To numberOfPeriods
                    aux(i, 1) = BussDay(aux(i, 1))
                    aux(i, 2) = BussDay(aux(i, 2))
                Next i
            End If

            Select Case pmtAdjustment
                Case 1
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2).AddDays(1))
                    Next i
                Case 2
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(BussDay(aux(i, 2).AddDays(1)).AddDays(1))
                    Next i
                Case Else
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2))
                    Next i
            End Select

            'Here we generate fixing dates

            If fixingRatio = 1 Then
                For i = 1 To numberOfPeriods
                    aux(i, 4) = lag(aux(i, 1), fixingLag)
                Next i
            Else

                If fixingRatio > numberOfPeriods Then fixingRatio = numberOfPeriods

                Select Case fixingStubPeriod
                    Case 1
                        resto = numberOfPeriods Mod fixingRatio

                        If resto > 0 Then
                            For i = 1 To resto
                                aux(i, 4) = lag(aux(1, 1), fixingLag)
                            Next i
                        End If

                        For i = 1 To numberOfPeriods - resto
                            If i Mod fixingRatio = 1 Then
                                aux(i + resto, 4) = lag(aux(i + resto, 1), fixingLag)
                            Else
                                aux(i + resto, 4) = aux(i + resto - 1, 4)
                            End If
                        Next i

                    Case 2
                        cociente = Int(numberOfPeriods / fixingRatio)
                        resto = numberOfPeriods Mod fixingRatio

                        If cociente > 0 Then
                            For i = 1 To numberOfPeriods - resto
                                If i Mod fixingRatio = 1 Then
                                    aux(i, 4) = lag(aux(i, 1), fixingLag)
                                Else
                                    aux(i, 4) = aux(i - 1, 4)
                                End If
                            Next i
                        End If

                        If resto > 0 Then
                            For i = numberOfPeriods - resto + 1 To numberOfPeriods
                                aux(i, 4) = lag(aux(numberOfPeriods - resto + 1, 1), fixingLag)
                            Next i
                        End If

                    Case Else
                        cociente = Int(numberOfPeriods / fixingRatio)
                        resto = numberOfPeriods Mod fixingRatio

                        aux(1, 4) = lag(aux(1, 1), fixingLag)

                        If cociente > 0 Then
                            For i = 2 To numberOfPeriods - resto
                                If i Mod fixingRatio = resto Then
                                    aux(i, 4) = lag(aux(i, 1), fixingLag)
                                Else
                                    aux(i, 4) = aux(i - 1, 4)
                                End If
                            Next i
                        End If

                        If resto > 0 Then
                            For i = numberOfPeriods - resto + 1 To numberOfPeriods
                                aux(i, 4) = lag(aux(numberOfPeriods - resto + 1, 1), fixingLag)
                            Next i
                        End If
                End Select
            End If

            'Here we generate amortization percentages

            Select Case typeAmortization
                Case 1
                    For i = 1 To numberOfPeriods - 1
                        aux(i, 5) = 0
                        aux(i, 6) = 1
                    Next i
                    aux(numberOfPeriods, 5) = 1
                    aux(numberOfPeriods, 6) = 1

                Case 2
                    constantAmo = 1 / numberOfPeriods
                    For i = 1 To numberOfPeriods
                        If i = 1 Then
                            aux(i, 6) = 1
                        Else
                            aux(i, 6) = aux(i - 1, 6) - aux(i - 1, 5)
                        End If

                        aux(i, 5) = constantAmo
                    Next i

                Case 3
                    aux1 = 0
                    For i = 1 To numberOfPeriods
                        aux1 = aux1 + (1 + amortizeRate) ^ (-YearFraction(aux(1, 1), aux(i, 2), "30/360"))
                    Next i
                    constantInstallment = 1 / aux1
                    aux1 = 1
                    For i = 1 To numberOfPeriods
                        If i = 1 Then
                            aux(i, 6) = 1
                        Else
                            aux(i, 6) = aux(i - 1, 6) - aux(i - 1, 5)
                        End If
                        aux(i, 5) = constantInstallment - aux1 * ((1 + amortizeRate) ^ (YearFraction(aux(i, 1), aux(i, 2), "30/360")) - 1)
                        aux1 = aux1 - aux(i, 5)
                    Next i

                Case Else
                    For i = 1 To numberOfPeriods - 1
                        aux(i, 5) = 0
                        aux(i, 6) = 1
                    Next i
                    aux(numberOfPeriods, 5) = 1
                    aux(numberOfPeriods, 6) = 1

            End Select

            result = DContraeObj(aux)

        Catch ex As Exception
            result = Nothing
        End Try

        Return result

    End Function
    <ExcelFunction(category:=Categoria, Description:="Genera calendario de pata variable de un Swap")> _
        Shared Function floatCalendar(<ExcelArgument(Description:="Fecha inicio")> ByVal startDate As Date, _
                                        <ExcelArgument(Description:="Fecha fin")> ByVal endDate As Date, _
                                        <ExcelArgument(Description:="Periocidad")> ByVal periodicity As Integer, _
                                        <ExcelArgument(Description:="Ajuste fechas fin de periodo: 0, no ajusta; 1, hábil siguiente (sólo si la fecha no es hábil)")> ByVal accAdjustment As Integer, _
                                        <ExcelArgument(Description:="Ajuste fechas de pago (respecto a fechas fin de periodo): 0, hábil siguiente (sólo si la fecha no es hábil); 1, hábil siguiente; 2, hábil subsiguiente")> ByVal pmtAdjustment As Integer, _
                                        <ExcelArgument(Description:="Periodo corto: 0, al final; 1, al principio")> ByVal typeStubPeriod As Integer, _
                                        <ExcelArgument(Description:="Rezago de fijación (respecto a fecha inicio): 0, hábil anterior (sólo si la fecha no es hábil); 1, un día hábil; 2, dos días hábiles ")> ByVal fixingLag As Integer, _
                                        <ExcelArgument(Description:="Periocidad fijación: n, se fija cada n periodos")> ByVal fixingRatio As Integer, _
                                        <ExcelArgument(Description:="Periodo fijación única: 0, al final; 1, al principio")> ByVal fixingStubPeriod As Integer)
        Dim aux(,) As Date
        Dim auxDate As Date
        Dim j As Integer
        Dim i As Integer
        Dim numberOfPeriods As Integer
        Dim cociente As Integer
        Dim resto As Integer
        Dim result As Object(,)

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        REM periodicity es necesario, sino la función entra en loop.
        If periodicity = 0 Then
            Return Nothing
        End If

        Try
            Select Case typeStubPeriod
                Case 1
                    auxDate = endDate
                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 4)

                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j

                Case 2
                    auxDate = startDate
                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While
                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 4)

                    For j = 1 To numberOfPeriods
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        If AddMonths(aux(j, 1), periodicity) < endDate Then
                            aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                        Else
                            aux(j, 2) = endDate
                        End If
                    Next j

                Case 3
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 2
                    ReDim aux(numberOfPeriods, 4)

                    For j = numberOfPeriods To 2 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                    Next j

                    aux(1, 1) = startDate
                    aux(1, 2) = aux(2, 1)

                Case 4
                    auxDate = startDate

                    i = 1
                    While auxDate < endDate
                        auxDate = AddMonths(startDate, i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 2
                    ReDim aux(numberOfPeriods, 4)

                    For j = 1 To numberOfPeriods - 1
                        aux(j, 1) = AddMonths(startDate, (j - 1) * periodicity)
                        aux(j, 2) = AddMonths(aux(j, 1), periodicity)
                    Next j

                    aux(numberOfPeriods, 1) = aux(numberOfPeriods - 1, 2)
                    aux(numberOfPeriods, 2) = endDate


                Case Else
                    auxDate = endDate

                    i = 1
                    While auxDate > startDate
                        auxDate = AddMonths(endDate, -i * periodicity)
                        i = i + 1
                    End While

                    numberOfPeriods = i - 1
                    ReDim aux(numberOfPeriods, 4)

                    For j = numberOfPeriods To 1 Step -1
                        aux(j, 2) = DateAdd("m", -(numberOfPeriods - j) * periodicity, endDate)
                        If DateAdd("m", -periodicity, aux(j, 2)) > startDate Then
                            aux(j, 1) = DateAdd("m", -periodicity, aux(j, 2))
                        Else
                            aux(j, 1) = startDate
                        End If
                    Next j


            End Select

            If accAdjustment <> 0 Then
                For i = 1 To numberOfPeriods
                    aux(i, 1) = BussDay(aux(i, 1))
                    aux(i, 2) = BussDay(aux(i, 2))
                Next i
            End If

            Select Case pmtAdjustment
                Case 1
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2).AddDays(+1))
                    Next i
                Case 2
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(BussDay(aux(i, 2).AddDays(+1)).AddDays(+1))
                    Next i
                Case Else
                    For i = 1 To numberOfPeriods
                        aux(i, 3) = BussDay(aux(i, 2))
                    Next i
            End Select

            If fixingRatio = 1 Then
                For i = 1 To numberOfPeriods
                    aux(i, 4) = lag(aux(i, 1), fixingLag)
                Next i
            Else
                If fixingRatio > numberOfPeriods Then fixingRatio = numberOfPeriods
                Select Case fixingStubPeriod
                    Case 1
                        resto = numberOfPeriods Mod fixingRatio

                        If resto > 0 Then
                            For i = 1 To resto
                                aux(i, 4) = lag(aux(1, 1), fixingLag)
                            Next i
                        End If

                        For i = 1 To numberOfPeriods - resto
                            If i Mod fixingRatio = 1 Then
                                aux(i + resto, 4) = lag(aux(i + resto, 1), fixingLag)
                            Else
                                aux(i + resto, 4) = aux(i + resto - 1, 4)
                            End If
                        Next i

                    Case 2
                        cociente = CInt(Int(numberOfPeriods / fixingRatio))
                        resto = numberOfPeriods Mod fixingRatio

                        If cociente > 0 Then
                            For i = 1 To numberOfPeriods - resto
                                If i Mod fixingRatio = 1 Then
                                    aux(i, 4) = lag(aux(i, 1), fixingLag)
                                Else
                                    aux(i, 4) = aux(i - 1, 4)
                                End If
                            Next i
                        End If

                        If resto > 0 Then
                            For i = numberOfPeriods - resto + 1 To numberOfPeriods
                                aux(i, 4) = lag(aux(numberOfPeriods - resto + 1, 1), fixingLag)
                            Next i
                        End If

                    Case Else
                        cociente = CInt(Int(numberOfPeriods / fixingRatio))
                        resto = numberOfPeriods Mod fixingRatio

                        aux(1, 4) = lag(aux(1, 1), fixingLag)

                        If cociente > 0 Then
                            For i = 2 To numberOfPeriods - resto
                                If i Mod fixingRatio = resto Then
                                    aux(i, 4) = lag(aux(i, 1), fixingLag)
                                Else
                                    aux(i, 4) = aux(i - 1, 4)
                                End If
                            Next i
                        End If

                        If resto > 0 Then
                            For i = numberOfPeriods - resto + 1 To numberOfPeriods
                                aux(i, 4) = lag(aux(numberOfPeriods - resto + 1, 1), fixingLag)
                            Next i
                        End If
                End Select
            End If

            'aux = ContraeMatriz(Of Date)(aux)
            result = DContrae(aux)
        Catch ex As Exception
            result = Nothing
        Finally
        End Try
        Return result
    End Function
    <ExcelFunction(category:=Categoria, Description:="")> _
    Shared Function forwardCurveFromCurrency(ByVal ccy As String) As String
        Dim retornar As String = ""

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ccy = UCase(ccy)
            If ccy = "USD" Then retornar = "cld" Else retornar = ccy
        Catch ex As Exception
        Finally
        End Try
        Return retornar
    End Function
#End Region
#Region "Funciones Auxiliares"
    Private Shared Function ContraeMatriz(Of t)(ByVal matriz(,) As t) As t(,)
        Dim NuevaMatriz(,) As t
        Dim i As Integer
        Dim j As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ReDim NuevaMatriz(matriz.GetUpperBound(0) - 1, matriz.GetUpperBound(1) - 1)
            For i = 1 To matriz.GetUpperBound(0)
                For j = 1 To matriz.GetUpperBound(1)
                    NuevaMatriz(i - 1, j - 1) = matriz(i, j)
                Next
            Next
        Catch ex As Exception
            NuevaMatriz = Nothing
        Finally
        End Try
        Return NuevaMatriz
    End Function
    Private Shared Function DContrae(ByVal matriz(,) As Date)
        Dim NuevaMatriz(,) As Object
        Dim i As Integer
        Dim j As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ReDim NuevaMatriz(matriz.GetUpperBound(0) - 1, matriz.GetUpperBound(1) - 1)
            For i = 1 To matriz.GetUpperBound(0)
                For j = 1 To matriz.GetUpperBound(1)
                    NuevaMatriz(i - 1, j - 1) = matriz(i, j)
                Next
            Next
        Catch ex As Exception
            NuevaMatriz = Nothing
        Finally
        End Try
        Return NuevaMatriz
    End Function
    Private Shared Function DContrae(ByVal matriz(,) As Double)
        Dim NuevaMatriz(,) As Double
        Dim i As Integer
        Dim j As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ReDim NuevaMatriz(matriz.GetUpperBound(0) - 1, matriz.GetUpperBound(1) - 1)
            For i = 1 To matriz.GetUpperBound(0)
                For j = 1 To matriz.GetUpperBound(1)
                    NuevaMatriz(i - 1, j - 1) = matriz(i, j)
                Next
            Next
        Catch ex As Exception
            NuevaMatriz = Nothing
        Finally
        End Try
        Return NuevaMatriz
    End Function
    Private Shared Function DContraeObj(ByVal matriz(,) As Object)
        Dim NuevaMatriz(,) As Object
        Dim i As Integer
        Dim j As Integer

        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If

        Try
            ReDim NuevaMatriz(matriz.GetUpperBound(0) - 1, matriz.GetUpperBound(1) - 1)
            For i = 1 To matriz.GetUpperBound(0)
                For j = 1 To matriz.GetUpperBound(1)
                    NuevaMatriz(i - 1, j - 1) = matriz(i, j)
                Next
            Next
        Catch ex As Exception
            NuevaMatriz = Nothing
        Finally
        End Try
        Return NuevaMatriz
    End Function
    Private Shared Function Range2Array(Of t)(ByVal Range As Object) As t()
        Dim Result() As t = Nothing
        Dim value As Object
        Dim Cont As Integer = 1
        Try
            If IsArray(Range) Then
                For Each x In Range
                    ReDim Preserve Result(Cont)
                    value = x
                    Result(Cont - 1) = CType(value, t)
                    Cont += 1
                Next
            Else
                ReDim Result(Cont)
                Result(0) = Range
            End If
        Catch ex As Exception
        End Try
        Return Result
    End Function
    Private Function Ver() As String
        Return "1.0.0"
    End Function
    Private Shared Shadows Function Right(ByVal texto As String, Optional ByVal largo As Integer = 1) As String
        Dim result As String
        If Not ValidaLicencia() Then
            MsgBox("Esta version no ha sido licenciada" & vbCrLf & "Favor comunicarse con Creasys S.A.")
            Return Nothing
        End If
        Try
            result = texto.Substring(texto.Length - largo)
        Catch ex As Exception
            result = Nothing
        End Try
        Return result
    End Function
    Private Shared Function IsMissing(ByVal o As Object) As Boolean
        MsgBox(o.ToString.ToLower())
        Return o.ToString.ToLower() = "exceldna.integration.excelmissing"
    End Function
#End Region
#Region "Validacion Licencia"
    Private Shared Function ValidaLicencia() As Boolean
        Return GenerarClave(MostrarInformacionDeDisco) = LeeRegistro()
    End Function
    ''' <summary>
    ''' Genera una clave del tipo xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx a partir de un string.
    ''' </summary>
    ''' <param name="sClave">Clave a Encriptar</param>
    ''' <returns>String del tipo xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx en hexadecimal</returns>
    ''' <remarks></remarks>
    Private Shared Function GenerarClave(ByVal sClave As String) As String
        ' Obtenemos la longitud de la cadena de usuario
        Dim longitud As Integer = sClave.Length
        Dim valorEntrada As Long = 0
        For I As Integer = 0 To longitud - 1
            valorEntrada += Asc(sClave.Substring(I, 1))
        Next
        valorEntrada \= longitud
        Dim valorBase As Long = valorEntrada * longitud
        Dim key As String = ""
        Dim valor As String = Hex(valorBase + (123 * 10000))
        key &= valor.Substring(valor.Length - 6, 6)
        valor = Hex(valorBase + (98 * 12500))
        key &= "-" & valor.Substring(0, 6)
        valor = Hex(valorBase + (77 * 15000))
        key &= "-" & valor.Substring(valor.Length - 6, 6)
        valor = Hex(valorBase + (121 * 17500))
        key &= "-" & valor.Substring(0, 6)
        valor = Hex(valorBase + (55 * 20000))
        key &= "-" & valor.Substring(valor.Length - 6, 6)
        valor = Hex(valorBase + (134 * 22500))
        key &= "-" & valor.Substring(0, 6)
        valor = Hex(valorBase + (63 * 25000))
        key &= "-" & valor.Substring(valor.Length - 6, 6)
        valor = Hex(valorBase + (117 * 27500))
        key &= "-" & valor.Substring(0, 6)
        Return key
    End Function
    Private Shared Function MostrarInformacionDeDisco() As String

        Dim sGetVol As New Volume.GetVol

        Dim sSerial As String

        sSerial = sGetVol.GetVolumeSerial("C")

        MostrarInformacionDeDisco = sSerial

        Exit Function

MensajeError:

        MostrarInformacionDeDisco = ""

    End Function
    Private Shared Function LeeRegistro() As String

        Return GetSetting("Derivados", "Licencia", "Value")

    End Function
    Private Sub GrabaRegistro()
        SaveSetting("Derivados", "Licencia", "Value", GenerarClave(MostrarInformacionDeDisco))
    End Sub
    Private Sub BorraRegistro()
        DeleteSetting("Derivados", "Licencia", "Value")
    End Sub
#End Region




End Class
