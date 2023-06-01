Attribute VB_Name = "ComObjectProperties"
Option Explicit

' https://renenyffenegger.ch/notes/development/languages/VBA/Useful-object-libraries/TypeLib-Information/index?fbclid=IwAR1_9eTnZcZ9CySJW_Exf63RA_zzKraw0MOc0eoTDd5R1nl_GHfoIRm8V5s
' Had issuess regarding the type library, worked after adding component service
' https://stackoverflow.com/questions/42569377/tlbinf32-dll-in-a-64bits-net-application/42581513#42581513

Public Sub propertiesOfObj(ByVal obj As Object) ' {

    Dim tlApp  As New TLI.TLIApplication
    Dim tlInfo As TLI.TypeInfo

    Dim attributes() As String
    Dim ix           As Long
    Dim nofAttrs     As Long

    Set tlInfo = tlApp.InterfaceInfoFromObject(obj)

    nofAttrs = tlInfo.AttributeStrings(attributes)

    Debug.Print "Name             = " & tlInfo.Name
    Debug.Print "GUID             = " & tlInfo.Guid
    Debug.Print "Kind             = " & tlInfo.TypeKindString
    Debug.Print "AttributeMask    = " & tlInfo.AttributeMask

    For ix = LBound(attributes) To UBound(attributes)
        Debug.Print "                   " & attributes(ix)
    Next ix

    Debug.Print "nof Interfaces   = " & tlInfo.Interfaces.Count
    Debug.Print "-----------------------"

    Dim mbrInfo As TLI.MemberInfo

    Dim i  As Long
    For Each mbrInfo In tlInfo.Members ' {
        i = i + 1
        Debug.Print lpad(mbrInfo.MemberId, 11) & " " & tlMemberKind(mbrInfo) & rpad(mbrInfo.Name, 40) & ": " & tlTypeName(mbrInfo.ReturnType)

        Dim parInfo As TLI.ParameterInfo
        For Each parInfo In mbrInfo.Parameters ' {

            Debug.Print "   " & tlParamKind(parInfo) & " " & rpad(parInfo.Name, 40) & ": " & tlTypeName(parInfo.VarTypeInfo)

        Next parInfo ' }

'       debug.print "   " & callingConvention(mbrInfo.callConv)
'       if i > 20 then exit sub

        Debug.Print ""

    Next mbrInfo ' }

End Sub ' }

Private Function tlMemberKind(mbr As TLI.MemberInfo) As String ' {

    Select Case mbr.DescKind ' {
    Case TLI.DESCKIND_FUNCDESC
           Select Case mbr.InvokeKind
           Case TLI.INVOKE_FUNC   ' {

                  If mbr.ReturnType.VarType = VT_VOID Then
                     tlMemberKind = "sub              "
                  Else
                     tlMemberKind = "function         "
                  End If


           Case TLI.INVOKE_PROPERTYGET: tlMemberKind = "property get     "
           Case TLI.INVOKE_PROPERTYPUT: tlMemberKind = "property put     "
           Case Else: tlMemberKind = "?                "
           End Select    ' }

    Case TLI.DESCKIND_VARDESC: tlMemberKind = "variable     "
    Case TLI.DESCKIND_NONE: tlMemberKind = "             "
    Case Else: tlMemberKind = "?            "
    End Select    ' }

End Function ' }

Private Function tlParamKind(par As TLI.ParameterInfo) As String ' {

    If par.Flags And PARAMFLAG_FOPT Then ' {
       tlParamKind = "optional "
    Else
       tlParamKind = ".        "
    End If ' }

    If par.Flags And PARAMFLAG_FOUT Then ' {
       tlParamKind = tlParamKind & "byRef "
    Else
       tlParamKind = tlParamKind & "byVal "
    End If ' }

End Function ' }

Private Function tlTypeName(var As TLI.VarTypeInfo) As String ' {
    Dim vType As TliVarType

    vType = var.VarType

    If vType And VT_ARRAY Then
       tlTypeName = "()"
       vType = vType And Not VT_ARRAY
    End If

    Select Case vType
    Case TLI.TliVarType.VT_BOOL: tlTypeName = tlTypeName & "boolean"
    Case TLI.TliVarType.VT_BSTR: tlTypeName = tlTypeName & "string"
    Case TLI.TliVarType.VT_CY: tlTypeName = tlTypeName & "currency"
    Case TLI.TliVarType.VT_DATE: tlTypeName = tlTypeName & "date"
    Case TLI.TliVarType.VT_DISPATCH: tlTypeName = tlTypeName & "object"
    Case TLI.TliVarType.VT_HRESULT: tlTypeName = tlTypeName & "HRESULT"
    Case TLI.TliVarType.VT_I2: tlTypeName = tlTypeName & "integer"
    Case TLI.TliVarType.VT_I4: tlTypeName = tlTypeName & "long"
    Case TLI.TliVarType.VT_R4: tlTypeName = tlTypeName & "single"
    Case TLI.TliVarType.VT_R8: tlTypeName = tlTypeName & "double"
    Case TLI.TliVarType.VT_UI1: tlTypeName = tlTypeName & "byte"
    Case TLI.TliVarType.VT_UNKNOWN: tlTypeName = tlTypeName & "IUnknown"
    Case TLI.TliVarType.VT_VARIANT: tlTypeName = tlTypeName & "variant"
    Case TLI.TliVarType.VT_VOID: tlTypeName = tlTypeName & "any"

    Case TliVarType.VT_EMPTY ' {

        If Not var.TypeInfo Is Nothing Then
            tlTypeName = tlTypeName & var.TypeInfo.Name
        End If

    End Select ' }

End Function ' }

Private Function callingConvention(cc As Long) As String ' {

    Select Case cc
    Case TLI.CC_CDECL: callingConvention = "cdecl   "
    Case TLI.CC_FASTCALL: callingConvention = "fastcall"
    Case TLI.CC_STDCALL: callingConvention = "stdcall "
    Case TLI.CC_SYSCALL: callingConvention = "syscall "
    Case Else: callingConvention = "TODO: implement me!"
    End Select

End Function ' }

Function rpad(text As String, length As Integer, Optional padChar As String = " ") ' {
 '
 '   https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/Text
 '
    rpad = text & String(length - Len(text), padChar)
End Function ' }

Function lpad(text As String, length As Integer, Optional padChar As String = " ") ' {
    lpad = String(length - Len(text), padChar) & text
End Function ' }
