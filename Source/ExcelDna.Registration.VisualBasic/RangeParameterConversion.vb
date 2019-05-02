Imports System.Linq.Expressions
Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration

Public Module RangeParameterConversion

    Function ReferenceToRange(xlInput As Object) As Range

        Dim reference As ExcelReference = xlInput   ' Will throw some Exception if not valid, which will be returned as #VALUE
        Dim app As Application = ExcelDnaUtil.Application

        Dim sheetName As String = XlCall.Excel(XlCall.xlSheetNm, reference)
        Dim index As Integer = sheetName.LastIndexOf("]")
        sheetName = sheetName.Substring(index + 1)
        Dim ws As Worksheet = app.Sheets(sheetName)
        Dim target As Range = app.Range(ws.Cells(reference.RowFirst + 1, reference.ColumnFirst + 1),
                                        ws.Cells(reference.RowLast + 1, reference.ColumnLast + 1))

        For iInnerRef As Long = 1 To reference.InnerReferences.Count - 1
            Dim innerRef As ExcelReference = reference.InnerReferences(iInnerRef)
            Dim innerTarget As Range = app.Range(ws.Cells(innerRef.RowFirst + 1, innerRef.ColumnFirst + 1),
                                                 ws.Cells(innerRef.RowLast + 1, innerRef.ColumnLast + 1))
            target = app.Union(target, innerTarget)
        Next
        Return target
    End Function

    Private Function UpdateAttributesForRangeParameters(reg As ExcelFunctionRegistration) As ExcelFunctionRegistration

        Dim rangeParams = From parWithIndex In reg.FunctionLambda.Parameters.Select(Function(par, i) New With {.Parameter = par, .Index = i})
                          Where parWithIndex.Parameter.Type.IsEquivalentTo(GetType(Range))
                          Select parWithIndex

        For Each param In rangeParams
            reg.ParameterRegistrations(param.Index).ArgumentAttribute.AllowReference = True
        Next

        Return reg
    End Function

    ' Must be run before the parameter conversions
    <Extension()>
    Function UpdateRegistrationsForRangeParameters(regs As IEnumerable(Of ExcelFunctionRegistration)) As IEnumerable(Of ExcelFunctionRegistration)
        Return regs.Select(AddressOf UpdateAttributesForRangeParameters)
    End Function

    ' NOTE: This parameter conversion should be registered to run for all types (with 'Nothing' as the TypeFilter)
    ' so that the COM-friendly type equivalence check here can be done, instead of exact type check.
    Function ParameterConversion(paramType As Type, paramRegistration As ExcelParameterRegistration)
        If paramType.IsEquivalentTo(GetType(Range)) Then
            Return CType(Function(input As Object) ReferenceToRange(input), Expression(Of Func(Of Object, Range)))
        Else
            Return Nothing
        End If
    End Function
End Module
