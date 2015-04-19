Imports System.Linq.Expressions
Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ExcelDna.Registration

Public Module RangeParameterConversion

    Function ReferenceToRange(ByVal xlInput As Object) As Range

        Dim xlRef As ExcelReference = xlInput   ' Will throw some Exception if not valid, which will be returned as #VALUE

        Dim cntRef As Long
        Dim strText As String
        Dim strAddress As String

        strAddress = XlCall.Excel(XlCall.xlfReftext, xlRef.InnerReferences(0), True)
        For cntRef = 1 To xlRef.InnerReferences.Count - 1
            strText = XlCall.Excel(XlCall.xlfReftext, xlRef.InnerReferences(cntRef), True)
            strAddress = strAddress & "," & Mid(strText, strText.LastIndexOf("!") + 2) ' +2 because IndexOf starts at 0
        Next
        ReferenceToRange = CType(ExcelDnaUtil.Application, Application).Range(strAddress)
    End Function

    Private Function UpdateAttributesForRangeParameters(reg As ExcelFunctionRegistration) As ExcelFunctionRegistration

        Dim rangeParams = From parWithIndex In reg.FunctionLambda.Parameters.Select(Function(par, i) New With {.Parameter = par, .Index = i})
                          Where parWithIndex.Parameter.Type.IsEquivalentTo(GetType(Range))
                          Select parWithIndex

        Dim hasRangeParam As Boolean = False
        For Each param In rangeParams
            reg.ParameterRegistrations(param.Index).ArgumentAttribute.AllowReference = True
            hasRangeParam = True
        Next

        If hasRangeParam Then
            reg.FunctionAttribute.IsMacroType = True
        End If

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
