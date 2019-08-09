Sub ArrumaData_Excel()
'
' converte data . para / - SAP
'
    
    Selection.TextToColumns DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4), TrailingMinusNumbers:=True
    
End Sub