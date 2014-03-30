Imports System.Runtime.CompilerServices
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

Public Module RangeParameterConversion

    Function ReferenceToRange(ByVal xlRef As ExcelReference) As Range
        Dim cntRef As Long, strText As String, strAddress As String
        strAddress = XlCall.Excel(XlCall.xlfReftext, xlRef.InnerReferences(0), True)
        For cntRef = 1 To xlRef.InnerReferences.Count - 1
            strText = XlCall.Excel(XlCall.xlfReftext, xlRef.InnerReferences(cntRef), True)
            strAddress = strAddress & "," & Mid(strText, strText.LastIndexOf("!") + 2) ' +2 because IndexOf starts at 0
        Next
        ReferenceToRange = CType(ExcelDnaUtil.Application, Application).Range(strAddress)
    End Function

    Private Function UpdateAttributesForRangeParameters(reg As ExcelFunctionRegistration) As ExcelFunctionRegistration
        
        Dim rangeParams = From parWithIndex In reg.FunctionLambda.Parameters.Select(Function(par, i) New With { .Parameter = par, .Index = i})
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

End Module
