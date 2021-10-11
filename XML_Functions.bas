Attribute VB_Name = "XML_Functions"
'
'Sub xml_flags()
'Dim xml_file As String
'Dim xml_text As String
'Dim textline As String
'Dim xml_obj As MSXML2.DOMDocument
'Dim child_node As MSXML2.IXMLDOMNode
'Dim accnums As String
'
'xml_file = "C:\Users\adanner\Desktop\xml.txt"
'Open xml_file For Input As #1
'Do Until EOF(1)
'    Line Input #1, textline
'    xml_text = xml_text & textline
'Loop
'Close #1
'
'Set xml_obj = load_xml(xml_text)
'
'If xml_obj Is Nothing Then
'    MsgBox "XML error"
'    Exit Sub
'End If
'
'' loop through each item and return Y/N flags
'accnums = node_value(xml_obj.DocumentElement.ChildNodes, "item", "accnum")
'Debug.Print accnums
'End Sub
'Function node_value(this_node As MSXML2.IXMLDOMNodeList, elmt As String, Optional sub_attr As String) As String
'Dim xNode As MSXML2.IXMLDOMNode
'Dim attr As MSXML2.IXMLDOMNode
'Dim keeper As String
'
'If IsMissing(sub_attr) Then sub_attr = ""
'
'For Each xNode In this_node
'    If xNode.nodeName = elmt And sub_attr = "" Then
'        keeper = add_val(xNode.nodeName, keeper)
'    ElseIf xNode.nodeName = elmt Then
'        Set attr = xNode.Attributes.getNamedItem(sub_attr)
'        If attr.nodeName = sub_attr Then
'            keeper = add_val(attr.NodeValue, keeper)
'        End If
'    End If
'
'    If xNode.HasChildNodes Then
'        'keeper = add_val(recursive_nodes(xNode.ChildNodes, elmt, sub_attr), keeper)
'    End If
'Next xNode
'recursive_nodes = keeper
'End Function
'Function add_val(adding As String, adding_to As String, Optional delimiter As String)
'If delimiter = "" Then delimiter = "|"
'
'If adding_to = "" Then
'    add_val = adding
'Else
'    add_val = adding_to & delimiter & adding
'End If
'
'End Function
'Sub PARCC_XML_Keys()
'Dim xl_file_name As Variant
'Dim n As Long
'Dim wb As Workbook
'Dim rept As Worksheet
'
'xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
'    "Browse for file to be XML keyed", MultiSelect:=True)
'
'Application.Workbooks.Add
'Set rept = ActiveWorkbook.Sheets(1)
'rept.Cells(1, 1) = "File name"
'rept.Cells(1, 2) = "Processed"
'
'Application.DisplayAlerts = False
'Application.ScreenUpdating = False
'
'If IsArray(xl_file_name) Then
'    For n = LBound(xl_file_name) To UBound(xl_file_name)
'        Application.Workbooks.Open Filename:=xl_file_name(n), ReadOnly:=False
'        rept.Cells(n + 1, 1) = xl_file_name(n)
'        Set wb = ActiveWorkbook
'        rept.Cells(n + 1, 2) = "Incomplete"
'        If wb.ReadOnly Then
'            rept.Cells(n + 1, 2) = "Workbook is read only"
'            wb.Close savechanges:=False
'        Else
'            Call XML_Keys(wb)
'            rept.Cells(n + 1, 2) = "Complete"
'            wb.Close savechanges:=True
'        End If
'
'    Next n
'End If
'
'Application.DisplayAlerts = False
'Application.ScreenUpdating = False
'
'End Sub
'Function XML_Keys(ByVal wb As Workbook)
'Dim parse As Worksheet
'Dim key As Worksheet
'Dim key_xml As String
'Dim part As String
'
'Set key = wb.Sheets(1)
'
'wb.Sheets.Add After:=wb.Sheets(ActiveWorkbook.Sheets.Count)
'Set parse = wb.Sheets(wb.Sheets.Count)
'
'With parse
'.Cells(1, 1) = "Form"
'.Cells(1, 2) = "Entity ID"
'.Cells(1, 3) = "All Answer Keys"
'.Cells(1, 4) = "Number of choices that the item/part has"
'.Cells(1, 5) = "Question No"
'.Cells(1, 6) = "QTI_INTERACTIONTYPE"
'.Cells(1, 7) = "QTI_CARDINALITY"
'.Cells(1, 8) = "QTI_ResponseId"
'.Cells(1, 9) = "Item Type"
'
'For n = 2 To key.UsedRange.Rows.Count
'    ' Form (Column A)
'    .Cells(n, 1) = key.Cells(n, 1)
'    key_xml = key.Cells(n, 128)
'
'    If key.Cells(n, 74) <> "" Then
'        ' need to get number of choices total
'        ' report is multipart
'        ' All Answer Keys (Parsed from “Grid Key 1” column DX)
'        ' Number of choices that the item/part has (Parsed from “Grid Key 1” column DX)
'        .Cells(n, 2) = key.Cells(n, 74)
'
'        If (Len(key_xml) - Len(Replace(key_xml, "correctResponse", ""))) / Len("correctResponse") > 2 Then
'            .Cells(n, 3) = "Multipart item"
'        End If
'    Else
'        ' All Answer Keys (Parsed from “Grid Key 1” column DX)
'        ' Number of choices that the item/part has (Parsed from “Grid Key 1” column DX)
'
'        .Cells(n, 2) = key.Cells(n, 76)
'    End If
'
'
'    ' Question No (Column BQ)
'    .Cells(n, 5) = key.Cells(n, 69)
'
'    ' QTI_INTERACTIONTYPE (Column IR)
'    .Cells(n, 6) = key.Cells(n, 252)
'
'    ' QTI_CARDINALITY (Column IT)
'    .Cells(n, 7) = key.Cells(n, 254)
'
'    ' QTI_ResponseId (Column IW)
'    .Cells(n, 8) = key.Cells(n, 257)
'
'    ' item type (Column CF)
'    .Cells(n, 9) = key.Cells(n, 84)
'
'    ' parse the keys only if they item is not multipart
'    If .Cells(n, 3) <> "Multipart item" And .Cells(n, 6) = "choiceInteraction" Then
'        part = .Cells(n, 8)
'        .Cells(n, 3) = get_item_keys(key_xml, part)
'        .Cells(n, 4) = num_choices(key_xml, part)
'    ElseIf .Cells(n, 3) <> "Multipart item" And .Cells(n, 6) <> "choiceInteraction" Then
'        .Cells(n, 3) = "Key Check Unsupported"
'        .Cells(n, 4) = "Key Check Unsupported"
'    End If
'
'Next n
'
'
'End With
'End Function
'Function load_xml(xml_str As String) As MSXML2.DOMDocument
'Dim xml_obj As MSXML2.DOMDocument
'Dim xml_nodes As Variant
'Dim n As Long
'Dim node_list As Collection
'Dim math_list() As Variant
'
'Set xml_obj = New MSXML2.DOMDocument
'xml_obj.validateOnParse = False
'
'' error should have already happened in images
'If Not xml_obj.LoadXML(xml_str) Then 'strXML is the string with XML'
'    Err.Raise xml_obj.parseError.ErrorCode, , xml_obj.parseError.reason
'    Exit Function
'End If
'Set load_xml = xml_obj
'End Function
'Function get_item_keys(xml_str As String, part As String) As String
'Dim xml_obj As MSXML2.DOMDocument
'Dim xml_nodes As Variant
'Dim n As Long
'Dim node_list As Collection
'Dim math_list() As Variant
'
'Set xml_obj = New MSXML2.DOMDocument
'xml_obj.validateOnParse = False
'
'' error should have already happened in images
'If Not xml_obj.LoadXML(xml_str) Then 'strXML is the string with XML'
'    Err.Raise xml_obj.parseError.ErrorCode, , xml_obj.parseError.reason
'    get_item_keys = "Parse error"
'    Exit Function
'End If
'
'get_item_keys = get_keys(xml_obj.DocumentElement.ChildNodes, part)
'End Function
'Function num_choices(xml_str As String, part As String) As Long
'Dim xml_obj As MSXML2.DOMDocument
'Dim xml_nodes As Variant
'Dim n As Long
'Dim node_list As Collection
'Dim math_list() As Variant
'
'Set xml_obj = New MSXML2.DOMDocument
'xml_obj.validateOnParse = False
'
'' error should have already happened in images
'If Not xml_obj.LoadXML(xml_str) Then 'strXML is the string with XML'
'    Err.Raise xml_obj.parseError.ErrorCode, , xml_obj.parseError.reason
'    num_choices = 0
'    Exit Function
'End If
'
'num_choices = get_choices(xml_obj.DocumentElement.ChildNodes, part)
'End Function
'Function get_choices(node_list As MSXML2.IXMLDOMNodeList, part As String, Optional sub_node As Boolean) As Long
'Dim xNode As MSXML2.IXMLDOMNode
'Dim num As Long
'Dim node_chk As MSXML2.IXMLDOMNode
'Dim arr() As Variant
'
'If IsMissing(sub_node) Then sub_node = False
'
'num = 0
'For Each xNode In node_list
'
'    If is_interactionType(xNode.nodeName) Then
'        If xNode.Attributes.getNamedItem("responseIdentifier").NodeValue = part Then
'            sub_node = True
'        Else
'            sub_node = False
'        End If
'    End If
'
'    If (xNode.nodeName = "simpleChoice" Or xNode.nodeName = "simpleAssociableChoice" Or xNode.nodeName = "extendedTextInteraction" Or _
'       xNode.nodeName = "associableHotspot" Or xNode.nodeName = "inlineChoice" Or xNode.nodeName = "customOption") And sub_node Then
'        num = num + 1
'    End If
'
'    If xNode.HasChildNodes Then
'        num = num + get_choices(xNode.ChildNodes, part, sub_node)
'    End If
'Next xNode
'get_choices = num
'End Function
'Function is_interactionType(node_name As String) As Boolean
'Dim arr() As Variant
'Dim elmt As Variant
'
'arr = Array("extendedTextInteraction", "textEntryInteraction", "choiceInteraction", "customInteraction", _
'    "hotspotInteraction", "matchInteraction", "gapMatchInteraction", "graphicGapMatchInteraction", _
'    "orderInteraction", "graphicOrderInteraction", "inlineChoiceInteraction")
'
'is_interactionType = False
'
'For Each elmt In arr
'    If node_name = elmt Then
'        is_interactionType = True
'        Exit Function
'    End If
'Next elmt
'
'
'End Function
'
'Function get_keys(node_list As MSXML2.IXMLDOMNodeList, part As String, Optional sub_node As Boolean) As String
'Dim xNode As MSXML2.IXMLDOMNode
'Dim keys As String
'Dim node_chk As MSXML2.IXMLDOMNode
'If IsMissing(sub_node) Then sub_node = False
'
'For Each xNode In node_list
'
'    If xNode.nodeName = "responseDeclaration" Then
'        If xNode.Attributes.getNamedItem("identifier").NodeValue = part Then
'            sub_node = True
'        Else
'            sub_node = False
'        End If
'    End If
'
'    If xNode.nodeName = "correctResponse" And sub_node Then
'        If keys = "" Then
'            keys = xNode.Text
'        Else
'            keys = keys & "|" & xNode.Text
'        End If
'    End If
'
'    If xNode.HasChildNodes Then
'        pot_key = get_keys(xNode.ChildNodes, part, sub_node)
'
'        If pot_key <> "" Then
'            If keys = "" Then
'                keys = pot_key
'            Else
'                keys = keys & "|" & pot_key
'            End If
'        End If
'    End If
'Next xNode
'get_keys = keys
'End Function
'Function Display_Node_math(Nodes As MSXML2.IXMLDOMNodeList, _
' math_cnt As Long, Optional math_list As Variant)
'
'Dim parent_list As String
'
'If IsMissing(math_list) Then
'    ReDim math_list(0 To 2, 0 To 0) As Variant
'End If
'For Each xNode In Nodes
'
'    If xNode.nodeName = "math" Then
'        parent_list = get_parent_nodes(xNode)
'        ReDim Preserve math_list(0 To 2, 0 To math_cnt)
'        math_list(0, math_cnt) = xNode.xml
'        math_list(1, math_cnt) = parent_list
'        math_list(2, math_cnt) = xNode.ParentNode.nodeName
'        math_cnt = math_cnt + 1
'    End If
'
'    If xNode.HasChildNodes Then
'        Display_Node_math xNode.ChildNodes, math_cnt, math_list
'    End If
'Next xNode
'Display_Node_math = math_list
'End Function
'
'
'
'Sub make_pretty()
'Dim xml_obj As MSXML2.DOMDocument
'Dim xml_str As String
'Dim n As Long
'
'For n = 2 To 3
'    xml_str = Cells(n, 128)
'    ' error should have already happened in images
'    If Not xml_obj.LoadXML(xml_str) Then 'strXML is the string with XML'
'        Err.Raise xml_obj.parseError.ErrorCode, , xml_obj.parseError.reason
'        Exit Sub
'    End If
'
'    Cells(n, 128) = PrettyPrintDocument(xml_obj)
'    Set xml_obj = Nothing
'Next n
'End Sub
'Function PrettyPrintDocument(Doc As MSXML2.DOMDocument) As String
'  PrettyPrintDocument = PrettyPrintXML(Doc.xml)
'End Function
'Function PrettyPrintXML(xml As String) As String
'Dim Reader As New SAXXMLReader60
'Dim Writer As New MXXMLWriter60
'
'Writer.Indent = True
'Writer.standalone = False
'Writer.omitXMLDeclaration = False
'Writer.Encoding = "utf-8"
'
'Set Reader.contentHandler = Writer
'Set Reader.dtdHandler = Writer
'Set Reader.errorHandler = Writer
'
'Call Reader.putProperty("http://xml.org/sax/properties/declaration-handler", _
'        Writer)
'Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", _
'        Writer)
'
'Call Reader.parse(xml)
'
'PrettyPrintXML = Writer.output
'
'End Function
'
'Sub xml_parse()
'Dim xml_sht As Worksheet
'Dim parse_sht As Worksheet
'
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'
'Set xml_sht = ActiveWorkbook.Sheets("XML Results")
'Set parse_sht = ActiveWorkbook.Sheets("XML Parse")
'
'parse_sht.Range(parse_sht.Rows(2), parse_sht.Rows(parse_sht.Rows.Count)).EntireRow.Delete
'
'Call get_imgs_mathml(xml_sht, parse_sht)
'
'Application.Calculation = xlCalculationAutomatic
'Application.ScreenUpdating = True
'End Sub
'Sub batcher()
'Dim xml_sht As Worksheet
'Dim deets_sht As Worksheet
'Dim alt_sht As Worksheet
'Dim details_sht As Worksheet
'Dim sht As Worksheet
'
'Dim batch_col As Integer
'batch_col = 15
'Application.ScreenUpdating = False
'Set alt_sht = ActiveWorkbook.Sheets("Authored Alt Text")
'Set deets_sht = ActiveWorkbook.Sheets("Summary Tracker")
'Set xml_sht = ActiveWorkbook.Sheets("XML Parse")
'Set details_sht = ActiveWorkbook.Sheets("Accessibility Detail")
'Call split_batches(deets_sht, xml_sht, alt_sht, batch_col, details_sht)
'
'For n = 1 To ActiveWorkbook.Sheets.Count
'    Set sht = ActiveWorkbook.Sheets(n)
'    If (sht.AutoFilterMode And sht.FilterMode) Or sht.FilterMode Then
'        sht.ShowAllData
'    End If
'Next n
'Application.ScreenUpdating = True
'End Sub
'Function get_unique_values(this_sht As Worksheet, this_col As Integer) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''                                This Macro was written by                                '''
''''                                      Asher Danner                                       '''
''''                                       03/14/2012                                        '''
''''The purpose is to return all unique values in a column as an array.                      '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim row_count As Long
'Dim next_col As Long
'With this_sht
'    next_col = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
'
'    .Columns(this_col).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Cells(1, next_col), unique:=True
'    row_count = .Cells(.Rows.Count, next_col).End(xlUp).Row
'
'    get_unique_values = WorksheetFunction.Transpose(.Range(.Cells(2, next_col), .Cells(row_count, next_col)))
'    .Columns(next_col).EntireColumn.Delete
'
'End With
'End Function
'Function split_batches(deets_sht As Worksheet, xml_sht As Worksheet, _
'    alt_sht As Worksheet, batch_col As Integer, details_sht As Worksheet)
'Dim new_book As Workbook
'Dim this_sht As Worksheet
'Dim this_col As Integer
'Dim acc_sht As Worksheet
'Dim alt_text_sht As Worksheet
'Dim rng As Range
'
'Set this_sht = deets_sht
'
'file_path = "C:\Batcher"
'this_col = batch_col
'
'this_array = get_unique_values(deets_sht, batch_col)
'For Each elmt In this_array
'    Application.Workbooks.Add
'    Set new_book = ActiveWorkbook
'    this_sht.UsedRange.AutoFilter Field:=this_col, Criteria1:=elmt
'    this_sht.Rows(1).Copy
'    new_book.Sheets(1).Cells(1, 1).PasteSpecial xlPasteColumnWidths
'    this_sht.UsedRange.SpecialCells(xlCellTypeVisible).Copy
'    new_book.Sheets(1).Cells(1, 1).PasteSpecial xlPasteValues
'    If new_book.Sheets.Count < 2 Then new_book.Sheets.Add After:=new_book.Sheets(new_book.Sheets.Count)
'    alt_sht.Rows(1).Copy
'    new_book.Sheets(2).Cells(1, 1).PasteSpecial xlPasteColumnWidths
'    this_sht.UsedRange.SpecialCells(xlCellTypeVisible).Copy
'    new_book.Sheets(1).Cells(1, 1).PasteSpecial xlPasteValues
'
'    If elmt <> "" Then
'        elmt = Replace(elmt, "?", "")
'        elmt = Replace(elmt, "\", "")
'        elmt = Replace(elmt, "/", "")
'        elmt = Replace(elmt, ":", "")
'        elmt = Replace(elmt, "<", "")
'        elmt = Replace(elmt, ">", "")
'        elmt = Replace(elmt, "|", "")
'
'        new_book.Sheets(1).Name = "Accessibility Details"
'        new_book.Sheets(2).Name = "Alt Text"
'        ' get images and mathml list
'        Call batch_img_mathml(new_book, xml_sht)
'        Set acc_sht = new_book.Sheets(1)
'        Set alt_text_sht = new_book.Sheets(2)
'
'
'        acc_sht.Rows(1).EntireRow.Delete
'        details_sht.Range(details_sht.Rows(1), details_sht.Rows(2)).Copy
'        acc_sht.Rows(1).Insert Shift:=xlDown
'        details_sht.Rows(2).Copy
'        acc_sht.Rows(2).PasteSpecial xlPasteColumnWidths
'        acc_sht.Range(acc_sht.Cells(3, 12), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 24)).Delete
'
'        ' now add all the formulas and drop downs
'        ' accessibility drop down
'        acc_sht.Range(acc_sht.Cells(3, 12), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 12)).ClearContents
'        With acc_sht.Range(acc_sht.Cells(3, 12), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 12)).Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'            xlBetween, Formula1:= _
'            "NA,1 - Non Essential,2 - Essential: Accessible with Alt-Text or  Video Transcription or text Transcription,3 - Essential: Accessible with Alt-Text or  Video Transcription or text Transcription and supplement,4 - Inaccessible"
'            .IgnoreBlank = True
'            .InCellDropdown = True
'            .InputTitle = ""
'            .ErrorTitle = ""
'            .InputMessage = ""
'            .ErrorMessage = ""
'            .ShowInput = True
'            .ShowError = True
'        End With
'
'        ' Yes/No drop downs
'        acc_sht.Range(acc_sht.Cells(3, 13), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 19)).ClearContents
'        With acc_sht.Range(acc_sht.Cells(3, 13), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 19)).Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'            xlBetween, Formula1:= _
'            "Yes,No"
'            .IgnoreBlank = True
'            .InCellDropdown = True
'            .InputTitle = ""
'            .ErrorTitle = ""
'            .InputMessage = ""
'            .ErrorMessage = ""
'            .ShowInput = True
'            .ShowError = True
'        End With
'
'        'Inaccessible Formula
'        With acc_sht.Range(acc_sht.Cells(3, 20), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 20))
'            .FormulaR1C1 = "=IF(COUNTIF(RC[-7]:RC[-1],""No"")=7,""Yes"",""No"")"
'            .Interior.Color = RGB(217, 217, 217)
'        End With
'
'        'Alt Text Needed? Formula
'        With acc_sht.Range(acc_sht.Cells(3, 21), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 21))
'            .FormulaR1C1 = "=IF(AND(OR(RC[-12]=""Yes"",RC[-10]>0)<>0,OR(RC[-5]=""Yes"",RC[-4]=""Yes""),OR(LEFT(RC[-9],1)=""2"",LEFT(RC[-9],1)=""3"")),""Yes"",""No"")"
'            .Interior.Color = RGB(217, 217, 217)
'        End With
'
'        With acc_sht.Range(acc_sht.Cells(3, 22), acc_sht.Cells(acc_sht.UsedRange.Rows.Count, 23)).Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'                xlBetween, Formula1:="Yes,No"
'            .IgnoreBlank = True
'            .InCellDropdown = True
'            .InputTitle = ""
'            .ErrorTitle = ""
'            .InputMessage = ""
'            .ErrorMessage = ""
'            .ShowInput = True
'            .ShowError = True
'        End With
'
'        alt_text_sht.Range(alt_text_sht.Cells(2, 7), alt_text_sht.Cells(alt_text_sht.UsedRange.Rows.Count, 8)).ClearContents
'
'        For n = 1 To 13
'            ' setup cell greying for No on Alt text
'            Set rng = alt_text_sht.Range(alt_text_sht.Cells(2, n), alt_text_sht.Cells(alt_text_sht.UsedRange.Rows.Count, n))
'
'            If n = 1 Then
'                rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
'                    "=IF(VLOOKUP(RC,'Accessibility Details'!C[2]:C[21],19,)=""No"",TRUE,FALSE)"
'            Else
'                rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
'                    "=IF(VLOOKUP(RC[-" & n - 1 & "],'Accessibility Details'!C[" & 2 - n + 1 & "]:C[" & 21 - n + 1 & "],19,)=""No"",TRUE,FALSE)"
'            End If
'            With rng.FormatConditions(1).Interior
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = -0.349986266670736
'            End With
'
'        Next n
'
'
'        new_book.Sheets(1).Select
'
'        new_book.SaveAs Filename:=file_path & "\" & elmt & ".xlsx", FileFormat:=xlOpenXMLWorkbook
'    End If
'    new_book.Close savechanges:=False
'
'Next elmt
'End Function
'Function batch_img_mathml(this_bk As Workbook, xml_sht As Worksheet)
'Dim alt_sht As Worksheet
'Dim deets_sht As Worksheet
'Dim n As Long
'Dim lr As Long
'Dim acc_col As Integer
'Dim accnums() As String
'
'acc_col = 3
'
'Set alt_sht = this_bk.Sheets("Alt Text")
'Set deets_sht = this_bk.Sheets("Accessibility Details")
'lr = deets_sht.Cells(deets_sht.Rows.Count, 1).End(xlUp).Row
'ReDim accnums(0 To lr - 1)
'For n = 2 To lr
'    accnums(n - 2) = deets_sht.Cells(n, acc_col)
'Next n
'
'' filter by accnums
'xml_sht.UsedRange.AutoFilter Field:=1, Criteria1:=accnums, Operator:=xlFilterValues ' arrays need Operators to filter
'
'' find the last visible row
'lr = xml_sht.Cells(xml_sht.Rows.Count, 1).End(xlUp).Row
'xml_sht.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=alt_sht.Cells(1, 1)
'
'End Function
'
'Function get_imgs_mathml(xml_sht As Worksheet, parse_sht As Worksheet)
'Dim img_list() As Variant
'Dim math_list As Variant
'Dim xml_str As String
'Dim i As Long
'Dim rRow As Long
'Dim n As Long
'
'rRow = 2
'For n = 2 To xml_sht.UsedRange.Rows.Count
'    xml_str = CStr(xml_sht.Cells(n, 4))
'
'    img_list = string_xml_image_parse(xml_str)
'
'    ' if there was an XML error return an error
'    If img_list(0, 0) = "Error" Then
'        parse_sht.Cells(rRow, 1) = xml_sht.Cells(n, 2) ' accnum
'        parse_sht.Cells(rRow, 2) = "XML Error"
'        parse_sht.Cells(rRow, 3) = "XML Error"
'        parse_sht.Cells(rRow, 4) = "XML Error"
'        parse_sht.Cells(rRow, 5) = "XML Error"
'        parse_sht.Cells(rRow, 6) = "XML Error"
'        parse_sht.Cells(rRow, 7) = xml_sht.Cells(n, 6)
'        parse_sht.Cells(rRow, 8) = xml_sht.Cells(n, 7)
'        rRow = rRow + 1
'    Else
'        If img_list(0, 0) <> "No Image" Then
'            ' loop over the image list and report information
'            For i = 0 To UBound(img_list, 2)
'                parse_sht.Cells(rRow, 1) = xml_sht.Cells(n, 2) ' accnum
'                If img_list(1, i) <> "Zone" Then
'                    parse_sht.Cells(rRow, 2) = "Image" ' Image/MathML
'                Else
'                    parse_sht.Cells(rRow, 2) = "Zone" ' Image/MathML
'                End If
'                parse_sht.Cells(rRow, 3) = img_list(0, i) ' image name
'                parse_sht.Cells(rRow, 4) = img_list(1, i) ' mime_types
'                parse_sht.Cells(rRow, 5) = img_list(2, i) ' Full path
'                parse_sht.Cells(rRow, 6) = img_list(3, i) ' Parent
'                parse_sht.Cells(rRow, 7) = xml_sht.Cells(n, 6)
'                parse_sht.Cells(rRow, 8) = xml_sht.Cells(n, 7)
'                rRow = rRow + 1
'            Next i
'        End If
'        math_list = string_xml_math_parse(xml_str)
'        If math_list(0, 0) <> "No Math" Then
'            For i = 0 To UBound(math_list, 2)
'                parse_sht.Cells(rRow, 1) = xml_sht.Cells(n, 2) ' accnum
'                parse_sht.Cells(rRow, 2) = "MathML"
'                parse_sht.Cells(rRow, 3) = math_list(0, i)
'                parse_sht.Cells(rRow, 4) = "~"
'                parse_sht.Cells(rRow, 5) = math_list(1, i)
'                parse_sht.Cells(rRow, 6) = math_list(2, i)
'                parse_sht.Cells(rRow, 7) = xml_sht.Cells(n, 6)
'                parse_sht.Cells(rRow, 8) = xml_sht.Cells(n, 7)
'                rRow = rRow + 1
'            Next i
'        End If
'    End If
'Next n
'
'End Function
'Function string_xml_math_parse(xml_str As String) As Variant
'Dim xml_obj As MSXML2.DOMDocument
'Dim xml_nodes As Variant
'Dim n As Long
'Dim node_list As Collection
'Dim math_list() As Variant
'
'Set xml_obj = New MSXML2.DOMDocument
'xml_obj.validateOnParse = False
'
'' error should have already happened in images
'If Not xml_obj.LoadXML(xml_str) Then 'strXML is the string with XML'
'    Err.Raise xml_obj.parseError.ErrorCode, , xml_obj.parseError.reason
'    Exit Function
'End If
'
'' get images in XML
'math_list = Display_Node_math(xml_obj.DocumentElement.ChildNodes, 0)
'If math_list(0, 0) = "" Then
'    math_list(0, 0) = "No Math"
'End If
'string_xml_math_parse = math_list
'End Function
'
'Function get_parent_nodes(ByVal pNode As MSXML2.IXMLDOMNode) As String
'Dim pList() As Variant
'Dim pCnt As Long
'Dim break_loop As Boolean
'
'' loop over all parent nodes until the at the top of the XML - item or set_leader
'break_loop = False
'Do Until break_loop = True
'    ReDim Preserve pList(0 To pCnt)
'    pList(pCnt) = pNode.ParentNode.nodeName
'
'    If pNode.ParentNode.nodeName = "item" Or pNode.ParentNode.nodeName = "set_leader" Then
'        break_loop = True
'    Else
'        Set pNode = pNode.ParentNode
'    End If
'
'    pCnt = pCnt + 1
'Loop
'Set pNode = Nothing
'
'' the above list creates a backwards list of parents
'' reorganize and send as final.
'For n = UBound(pList) To LBound(pList) Step -1
'    If parent_list = "" Then
'        parent_list = pList(n)
'    Else
'        parent_list = parent_list & " : " & pList(n)
'    End If
'Next n
'get_parent_nodes = parent_list
'End Function
'Function string_xml_image_parse(xml_str As String) As Variant
'Dim xml_obj As MSXML2.DOMDocument
'Dim xml_nodes As Variant
'Dim n As Long
'Dim node_list As Collection
'Dim img_list() As Variant
'Dim file_names() As Variant
'Dim file_names_unique() As Variant
'Dim img_list_unique() As Variant
'Dim zones As Long
'Dim first_zone As Boolean
'first_zone = True
'
'zones = 0
'Set xml_obj = New MSXML2.DOMDocument
'xml_obj.validateOnParse = False
'
'' check for errors in XML
'If Not xml_obj.LoadXML(xml_str) Then 'strXML is the string with XML'
'    ReDim img_list(0 To 0, 0 To 0)
'    img_list(0, 0) = "Error"
'    string_xml_image_parse = img_list
'    Exit Function
'End If
'
'' get images in XML
'img_list = Display_Node_Img(xml_obj.DocumentElement.ChildNodes, "", 0)
'If img_list(0, 0) = "" Then
'    img_list(0, 0) = "No Image"
'    string_xml_image_parse = img_list
'    Exit Function
'End If
'ReDim file_names(0 To UBound(img_list, 2))
'
'' create array of accnums
'For n = 0 To UBound(img_list, 2)
'    file_names(n) = img_list(0, n)
'Next n
'
'
'' get unique accnum values
'file_names_unique = unique(file_names)
'
''loop over file_names_unique and img_list to compare to the list of images, combine mimetypes as new array
'ReDim img_list_unique(0 To 3, 0 To UBound(file_names_unique))
'
'For n = 0 To UBound(img_list, 2) ' n gives the location in img_list
'    For i = 0 To UBound(file_names_unique) ' i gives the location in the unique list
'        If img_list(0, n) = file_names_unique(i) And img_list(0, n) <> "Zone" Then
'            ' each zone element needs to be listed many times
'            If img_list_unique(1, i) = "" Then
'                img_list_unique(0, i) = file_names_unique(i)
'                img_list_unique(1, i) = img_list(1, n)
'                img_list_unique(2, i) = img_list(2, n)
'                img_list_unique(3, i) = img_list(3, n)
'            Else
'                img_list_unique(1, i) = img_list_unique(1, i) & "," & img_list(1, n)
'            End If
'        End If
'    Next i
'Next n
'
'' add lines for zones
'For n = 0 To UBound(img_list, 2) ' n gives the location in img_list
'    If img_list(1, n) = "Zone" Then
'        If first_zone = True Then
'            first_zone = False
'        Else
'            zones = zones + 1
'            ReDim Preserve img_list_unique(0 To 3, 0 To UBound(file_names_unique) + zones)
'        End If
'        img_list_unique(0, UBound(file_names_unique) + zones) = "Zone"
'        img_list_unique(1, UBound(file_names_unique) + zones) = "Zone"
'        img_list_unique(2, UBound(file_names_unique) + zones) = img_list(2, n)
'        img_list_unique(3, UBound(file_names_unique) + zones) = img_list(3, n)
'    End If
'Next n
'
'string_xml_image_parse = img_list_unique
'End Function
'Function Display_Node_Img(Nodes As MSXML2.IXMLDOMNodeList, _
'ByRef parent_list As String, img_cnt As Long, Optional img_list As Variant)
'
'Dim xNode As MSXML2.IXMLDOMNode
'Dim attrib As Variant
'If IsMissing(img_list) Then
'    ReDim img_list(0 To 3, 0 To 0) As Variant
'End If
'
'For Each xNode In Nodes
'    If xNode.nodeName = "graphic" Then
'        parent_list = get_parent_nodes(xNode)
'        ReDim Preserve img_list(0 To 3, 0 To img_cnt)
'        Set attrib = xNode.Attributes.getNamedItem("external_file_name")
'        If attrib Is Nothing Then
'            img_list(0, img_cnt) = xNode.Attributes.getNamedItem("name").NodeValue & "." & xNode.Attributes.getNamedItem("mimetype").NodeValue
'        Else
'            img_list(0, img_cnt) = xNode.Attributes.getNamedItem("external_file_name").NodeValue
'        End If
'        img_list(1, img_cnt) = xNode.Attributes.getNamedItem("mimetype").NodeValue
'        img_list(2, img_cnt) = parent_list
'        img_list(3, img_cnt) = xNode.ParentNode.nodeName
'
'        'remove mimetype from external filename
'        img_list(0, img_cnt) = Replace(img_list(0, img_cnt), "." & img_list(1, img_cnt), "")
'
'        img_cnt = img_cnt + 1
'    ElseIf xNode.nodeName = "zone_choice" Then
'        parent_list = get_parent_nodes(xNode)
'        ReDim Preserve img_list(0 To 3, 0 To img_cnt)
'
'        img_list(0, img_cnt) = "Zone" ' filename
'        img_list(1, img_cnt) = "Zone" ' file type
'        img_list(2, img_cnt) = parent_list ' parent list
'        img_list(3, img_cnt) = xNode.ParentNode.nodeName '
'        img_cnt = img_cnt + 1
'    End If
'
'    If xNode.HasChildNodes Then
'        Display_Node_Img xNode.ChildNodes, parent_list, img_cnt, img_list
'    End If
'Next xNode
'Display_Node_Img = img_list
'End Function
'Function unique(aFirstArray) As Variant
'  Dim arr As New Collection, a
'  Dim i As Long
'  Dim uni_arr() As Variant
'
'  On Error Resume Next
'  For Each a In aFirstArray
'     arr.Add a, a
'  Next
'
'  For i = 1 To arr.Count
'     ReDim Preserve uni_arr(0 To q)
'     uni_arr(q) = arr(i)
'     q = q + 1
'  Next
'unique = uni_arr
'End Function
'Sub metadata_from_search_XML()
'last_col = ActiveSheet.UsedRange.Columns.Count
'last_row = ActiveSheet.UsedRange.Rows.Count
'For n = 2 To last_row
'    xml_txt = Cells(n, 9)
'    xml_txt = clean_xml(CStr(xml_txt))
'    xml_lines = Split(xml_txt, Chr(10))
'    i = last_col + 1
'    For Each xml_line In xml_lines
'        xml_line = Trim(xml_line)
'        If InStr(xml_line, "<metadata") <> 0 Or InStr(xml_line, "<operator") <> 0 _
'        Or InStr(xml_line, "<value") <> 0 Or InStr(xml_line, "<join") <> 0 Then
'            If has_begin_and_end(CStr(xml_line)) Then
'                metadata = extract_tag(CStr(xml_line))
'                Cells(n, i) = Trim(metadata)
'            ElseIf InStr(xml_line, "<join/>") = 0 Then
'                Cells(n, i) = "XML read error"
'            End If
'            i = i + 1
'        End If
'    Next xml_line
'
'Next n
'End Sub
'Function clean_xml(xml_txt As String) As String
'
'xml_txt = Replace(xml_txt, Chr(10), "")
'xml_txt = Replace(xml_txt, Chr(13), "")
'xml_txt = Replace(xml_txt, "<metadata", Chr(10) & "<metadata")
'xml_txt = Replace(xml_txt, "<operator", Chr(10) & "<operator")
'xml_txt = Replace(xml_txt, "<value", Chr(10) & "<value")
'xml_txt = Replace(xml_txt, "<join", Chr(10) & "<join")
'clean_xml = xml_txt
'
'End Function
'Sub metadata_XML()
'last_col = ActiveSheet.UsedRange.Columns.Count
'last_row = ActiveSheet.UsedRange.Rows.Count
'For n = 2 To last_row
'    xml_txt = Cells(n, 9)
'    xml_txt = clean_xml(CStr(xml_txt))
'    xml_lines = Split(xml_txt, Chr(10))
'    i = last_col + 1
'    For Each xml_line In xml_lines
'        xml_line = Trim(xml_line)
'        If InStr(xml_line, "<metadata") <> 0 Then
'            If has_begin_and_end(CStr(xml_line)) Then
'                metadata = extract_tag(CStr(xml_line))
'                Cells(n, i) = Trim(metadata)
'            ElseIf InStr(xml_line, "<join/>") = 0 Then
'                Cells(n, i) = "XML read error"
'            End If
'            i = i + 1
'        End If
'    Next xml_line
'
'Next n
'End Sub
'Function extract_tag(xml_txt As String) As String
'extract_tag = ""
'If xml_txt = "" Then Exit Function
'
'If InStr(xml_txt, "</") <> 0 Then
'    extract_tag = Mid(xml_txt, InStr(xml_txt, ">") + 1, InStr(xml_txt, "</") - 1 - InStr(xml_txt, ">"))
'End If
'End Function
'Function has_begin_and_end(xml_txt As String) As Boolean
'has_begin_and_end = False
'If InStr(xml_txt, "<") <> 0 And InStr(xml_txt, ">") <> 0 Then
'    xml_tag = Mid(xml_txt, InStr(xml_txt, "<") + 1, InStr(xml_txt, ">") - 1 - InStr(xml_txt, "<"))
'    If InStr(xml_tag, " ") <> 0 Then
'        xml_tag = Left(xml_tag, InStr(xml_tag, " ") - 1)
'    End If
'    If count_of_string(xml_txt, CStr(xml_tag)) > 1 Then
'        has_begin_and_end = True
'    End If
'End If
'
'End Function
'Function count_of_string(full_string As String, count_string As String) As Long
'count_of_string = (Len(full_string) - Len(Replace(full_string, count_string, ""))) / Len(count_string)
'End Function
'Sub Fill_IC_Names()
'Dim endrow As Long
'Dim i As Long
'Dim start_col As Long
'Dim end_col As Long
'
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'
'start_col = 1
'end_col = 1
'
'endrow = ActiveSheet.UsedRange.Rows.Count
'For i = start_col To end_col
'
'this_name = Cells(2, i)
'
'For n = 2 To endrow
'    If Cells(n, i) = "" Then
'        Cells(n, i) = this_name
'    Else
'        this_name = Cells(n, i)
'    End If
'Next n
'Next i
'For n = endrow To 2 Step -1
'    If Cells(n, 3) = "" Then
'        Cells(n, i).EntireRow.Delete
'    End If
'Next n
'
'Application.ScreenUpdating = True
'Application.Calculation = xlCalculationAutomatic
'End Sub
'
'
