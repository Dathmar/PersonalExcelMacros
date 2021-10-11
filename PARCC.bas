Attribute VB_Name = "PARCC"
Option Explicit
Option Base 0
Sub PARCC_Tech_Checks()
Dim row_count As Long
Dim tech_sht As Worksheet
Dim n As Long
Dim has_err As Boolean
Dim images As String
Dim art As Variant
Dim i As Integer
Dim gif_cnt As Integer
Dim svg_cnt As Integer
Dim eps_cnt As Integer
Dim types() As Variant
Dim sets As Long
Dim chk As String

Set tech_sht = ActiveSheet
row_count = tech_sht.UsedRange.Rows.Count

types = Array(".gif", ".svg", ".eps")

tech_sht.Range(tech_sht.Rows(2), tech_sht.Rows(row_count)).Interior.ColorIndex = xlNone
' Math specific tech checks
For n = 2 To row_count
has_err = False
With tech_sht
    Application.StatusBar = "Processing... " & (n - 1) / (row_count - 1) & " complete"
    ' Notes legend
    ' * means not a macro check
    ' %%%%% means further clarification needed
    ' & means new check not in QC document yet
    ' $ means check was done in previous steps step noted in ()
    ' @ means checks have been tested and verified working.
    
    ' @#C1. Description field is blank.
    If .Cells(n, 6) <> vbNullString Then
        .Cells(n, 6).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' #@C2. Alt text is present and correct for all SVG images (or GIF in the case of background images)
    ' and committee-confirmed glossing is applied to requested words.
    ' ----------------------------Notes------------------------------
    ' Include classifications of alt test [1], [2], [3] - this is per svg or gif file
    ' columns map 25 -> 26 | 28 -> 29 | 30 -> 31 | 32 -> 33
    
    ' columns check 25 -> 26 ONLY checks per SVG
    ' Compare SVGs to alt text
    If Not image_match_text(.Cells(n, 25), ".svg", .Cells(n, 26), "|") Then
        .Cells(n, 26).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' compare SVGs to number of classifications
    svg_cnt = (Len(.Cells(n, 25)) - Len(Replace(.Cells(n, 25), ".svg", ""))) / 4
    If Not check_classifications(.Cells(n, 26), svg_cnt) Then
        .Cells(n, 26).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' columns check 28 -> 29 ONLY checks per SVG
    ' Compare SVGs to alt text
    If Not image_match_text(.Cells(n, 28), ".svg", .Cells(n, 29), "|") Then
        .Cells(n, 29).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' compare SVGs to number of classifications
    svg_cnt = (Len(.Cells(n, 28)) - Len(Replace(.Cells(n, 28), ".svg", ""))) / 4
    If Not check_classifications(.Cells(n, 29), svg_cnt) Then
        .Cells(n, 29).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' columns check 30 -> 31 ONLY checks per SVG
    ' Compare SVGs to alt text
    If Not image_match_text(.Cells(n, 30), ".svg", .Cells(n, 31), "|") Then
        .Cells(n, 30).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' compare SVGs to number of classifications
    svg_cnt = (Len(.Cells(n, 30)) - Len(Replace(.Cells(n, 30), ".svg", ""))) / 4
    If Not check_classifications(.Cells(n, 31), svg_cnt) Then
        .Cells(n, 31).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' columns check 32 -> 33 ONLY checks per GIF
    ' Compare GIFs to alt text
    If Not image_match_text(.Cells(n, 32), ".gif", .Cells(n, 33), "|") Then
        .Cells(n, 32).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' compare SVGs to number of classifications
    svg_cnt = (Len(.Cells(n, 32)) - Len(Replace(.Cells(n, 32), ".gif", ""))) / 4
    If Not check_classifications(.Cells(n, 33), svg_cnt) Then
        .Cells(n, 33).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' *C3. Item is keyed correctly and scoring follows the guidelines in the PARCC Scoring Cheat Sheet (Math).
    ' ----------------------------Notes------------------------------
    ' Does not check in Macro this is a content check
    
    ' #C4. Max Points for each item is correctly entered either in the scoring guide (Type I Equation Editor,
    ' Type II, or Type III items) or in the Content tab of the item (all other items, except Numeric which
    ' does not have a UI field for Max Points and defaults to 1).
    ' ----------------------------Notes------------------------------
    ' if 15 has ExtendedText or Composite CR column 58 should match right most character of 45
    ' otherwise 20 matches should match right most character of 45
    ' -------------------------Assumptions---------------------------
    ' Assumes single digit score.  Conversts score from text to integer.
    If (InStr(.Cells(n, 15), "ExtendedText") <> 0 Or InStr(.Cells(n, 15), "Composite CR") <> 0) Then
        If IsNumeric(Trim(.Cells(n, 58))) And IsNumeric(Right(.Cells(n, 45), 1)) Then
            If CInt(Right(.Cells(n, 45), 1)) <> CInt(Trim(.Cells(n, 58))) Then
                .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
                .Cells(n, 58).Interior.Color = RGB(255, 0, 0)
                has_err = True
            End If
        Else
            .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
            .Cells(n, 58).Interior.Color = RGB(255, 0, 0)
            has_err = True
        End If
    ElseIf (InStr(.Cells(n, 15), "ExtendedText") = 0 Or InStr(.Cells(n, 15), "Composite CR") = 0) Then
        If IsNumeric(Trim(.Cells(n, 20))) And IsNumeric(Right(.Cells(n, 45), 1)) Then
            If CInt(Right(.Cells(n, 45), 1)) <> CInt(Trim(.Cells(n, 20))) Then
                .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
                .Cells(n, 20).Interior.Color = RGB(255, 0, 0)
                has_err = True
            End If
        Else
            .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
            .Cells(n, 20).Interior.Color = RGB(255, 0, 0)
            has_err = True
        End If
    End If
    
    
    ' #C5. There are no 1-part Composite items.
    If InStr(.Cells(n, 15), "Composite") <> 0 And .Cells(n, 15) = 1 Then
        .Cells(n, 15).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C6. Each Part ID in Composite items matches the Part ID Ref for that item part in the scoring information
    ' section of the item.
    ' ----------------------------Notes------------------------------
    ' if composite item then 18 = 19 This check should not take into account composite CRs
    If InStr(.Cells(n, 15), "Composite") <> 0 And InStr(.Cells(n, 15), "Composite CR") = 0 And .Cells(n, 18) <> .Cells(n, 19) Then
        .Cells(n, 18).Interior.Color = RGB(255, 0, 0)
        .Cells(n, 19).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' *C7. Each InlineChoiceList in an item has the same number of choices. If it is necessary to use a blank
    ' choice, it must be at the end of the list.
    ' ----------------------------Notes------------------------------
    ' Does not check in Macro this is a content check
    
    
    ' #C8. There is no more than 1 Source Key ID per Source Key.
    If .Cells(n, 21) = "YES" Then
        .Cells(n, 21).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' ^C9. Each Type I Equation Editor, Type II, or Type III item has an item-level Scoring Guide object, and the
    ' final Scoring Guide PDF is attached to the Scoring Guide Attachments tab as Type=Scoring Guide and
    ' Mimetype=Portable Document Format as outlined in the PARCC Scoring Objects (Math) document.
    ' ----------------------------Notes------------------------------
    ' For 44 is Type II or Type III
    
    
    ' &@#C10. column 56 contains the correct accnum
    If InStr(.Cells(n, 56), .Cells(n, 4)) = 0 And .Cells(n, 56) <> vbNullString Then
        .Cells(n, 56).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    '************
    ' #C1. The correct classification scheme is used in the item as listed in PARCC Current Classifications.
    ' ----------------------------Notes------------------------------
    ' found in column 40
    ' #Mathematics - Algebra I__FINAL
    ' #Mathematics - Algebra II_FINAL
    ' @#Mathematics - Geometry_FINAL
    ' #Mathematics - 3rd Grade_FINAL
    ' #Mathematics - 4th Grade__FINAL
    ' #Mathematics - 5th Grade__FINAL
    ' #Mathematics - 6th Grade_final-9/27
    ' #Mathematics - 7th Grade FINAL-9/28
    ' #Mathematics - 8th Grade_FINAL
    ' #Integrated Math
    If .Cells(n, 40) <> "Mathematics - Algebra I__FINAL" And .Cells(n, 40) <> "Mathematics - Algebra II_FINAL" And .Cells(n, 40) <> "Mathematics - Geometry_FINAL" _
    And .Cells(n, 40) <> "Mathematics - 3rd Grade_FINAL" And .Cells(n, 40) <> "Mathematics - 4th Grade__FINAL" And .Cells(n, 40) <> "Mathematics - 5th Grade__FINAL" _
    And .Cells(n, 40) <> "Mathematics - 6th Grade_final-9/27" And .Cells(n, 40) <> "Mathematics - 7th Grade FINAL-9/28" And .Cells(n, 40) <> "Mathematics - 8th Grade_FINAL" _
    And .Cells(n, 40) <> "Integrated Math" Then
        .Cells(n, 40).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C2. Cla R1C1 must not be blank.
    If .Cells(n, 41) = vbNullString Then
        .Cells(n, 41).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C3. Cla R1C2 must not be blank.
    If .Cells(n, 42) = vbNullString Then
        .Cells(n, 42).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' ----------------------------Notes------------------------------
    ' R1C3 - column 43 does not get checked
    
    ' #C4. Cla R1C4 must not be blank.
    If .Cells(n, 44) = vbNullString Then
        .Cells(n, 44).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C5. Cla R1C5 must not be blank.
    If .Cells(n, 45) = vbNullString Then
        .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' #The first character in R1C5 must match the character in R1C4.
    If .Cells(n, 45) = vbNullString Or .Cells(n, 44) = vbNullString Then
        .Cells(n, 44).Interior.Color = RGB(255, 0, 0)
        .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
        has_err = True
    ElseIf Left(.Cells(n, 45), InStr(.Cells(n, 45), ".") - 1) <> Right(.Cells(n, 44), Len(Left(.Cells(n, 45), InStr(.Cells(n, 45), ".") - 1))) Then
        .Cells(n, 44).Interior.Color = RGB(255, 0, 0)
        .Cells(n, 45).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    ' $( Tech checks C4) The number after the period is the number of points.
    ' $( Tech checks C4) For non-CRs number of points should match max points from the Scoring Information on the Content tab.
    ' $( Tech checks C4) For CRs number of points should match max points/score from scoring object report.


    ' #C6. Cla R1C6 must not be blank.
    If .Cells(n, 46) = vbNullString Then
        .Cells(n, 46).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C7. Cla R1C7 must not be blank.
    If .Cells(n, 47) = vbNullString Then
        .Cells(n, 47).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #R1C7 should never have the value Mid-Year.
    If .Cells(n, 47) = "Mid-Year" Then
        .Cells(n, 47).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' #If R1C4 is II or III, then R1C7 should be Performance Based.
    ' ---------------------------Change------------------------------
    ' I used R1C5 for the Performance Based check.
    If .Cells(n, 45) = vbNullString Or .Cells(n, 47) = vbNullString Then
        .Cells(n, 47).Interior.Color = RGB(255, 0, 0)
        has_err = True
    ElseIf (Left(.Cells(n, 45), InStr(.Cells(n, 45), ".") - 1) = "II" Or Left(.Cells(n, 45), InStr(.Cells(n, 45), ".") - 1) = "III") And .Cells(n, 47) <> "Performance Based" Then
        .Cells(n, 47).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C8. Cla R1C8 must not be blank.
    If .Cells(n, 48) = vbNullString Then
        .Cells(n, 48).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C9. Cla R3C1 must not be blank.
    If .Cells(n, 49) = vbNullString Then
        .Cells(n, 49).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' #Value for R3C1 should be 2.5 times the value after the period in R1C5.
    If CStr(CDbl(.Cells(n, 49) / 2.5)) <> CStr(Right(.Cells(n, 45), 1)) Then
        .Cells(n, 49).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C10. Cla R3C8 must not be blank.
    If .Cells(n, 50) = vbNullString Then
        .Cells(n, 50).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C11. Wildcard 2 must not be blank.
    If .Cells(n, 51) = vbNullString Then
        .Cells(n, 51).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' #If R1C4 is II or III, Wildcard 2 should be Human Scoring.
    ' ---------------------------Change------------------------------
    ' I used R1C5 for the Performance Based check.
    If .Cells(n, 45) = vbNullString Then
        .Cells(n, 47).Interior.Color = RGB(255, 0, 0)
        has_err = True
    ElseIf (Left(.Cells(n, 45), InStr(.Cells(n, 45), ".") - 1) = "II" Or Left(.Cells(n, 45), InStr(.Cells(n, 45), ".") - 1) = "III") And .Cells(n, 51) <> "Human Scoring" Then
        .Cells(n, 47).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C12. Wildcard 3 must not be blank.
    If .Cells(n, 52) = vbNullString Then
        .Cells(n, 52).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #C13. User Code 10 must not be blank IF the item has enemies.
    ' #If the item has enemies, enemy accnums should be filled in using a colon to delimit multiple enemies. Example: VF123456:VH987654
    If .Cells(n, 53) <> vbNullString Then
        If (Len(.Cells(n, 53)) - Len(Replace(.Cells(n, 53), "V", ""))) - (Len(.Cells(n, 53)) - Len(Replace(.Cells(n, 53), ":", ""))) - 1 > 0 Then
            .Cells(n, 53).Interior.Color = RGB(255, 0, 0)
            has_err = True
        End If
    End If
    
    ' #Make sure the item is not it's own enemy
    If InStr(.Cells(n, 53), .Cells(n, 4)) <> 0 Then
        .Cells(n, 53).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #@QC1. Items with art contain an SVG and an EPS file.
    ' ----------------------------Notes------------------------------
    ' need to check 3 file types svg, gif-(will go away), eps
    ' columns 25 and 28 should contain gif, svg and EPS for each set of images
    ' does not need to check attachments
    
    ' #@column 25 check
    sets = verify_image_sets(.Cells(n, 25), types, ", ", .Cells(n, 4))
    If sets <> CInt(.Cells(n, 24)) And .Cells(n, 25) <> vbNullString And .Cells(n, 24) <> 0 Then
        .Cells(n, 24).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    If sets = 0 Then
        .Cells(n, 25).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' #@column 28 check
    
    sets = verify_image_sets(.Cells(n, 28), types, "| ", .Cells(n, 4))
    If sets <> CInt(.Cells(n, 27)) And .Cells(n, 28) <> vbNullString And .Cells(n, 27) <> 0 Then
        .Cells(n, 27).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    If sets = 0 Then
        .Cells(n, 28).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    ' QC2. Items with background images have the GIF file loaded in the background image graphic, and the SVG and EPS files are attached on the item Attachments tab. SVG attachment has a Description of "svg replacement". EPS attachment has a Description of "eps export".
    ' ----------------------------Notes------------------------------
    ' need to check 3 file types svg, gif-(will go away), eps
    ' --- need to check in the attachments (39) --- if eps should have | and eps export if svg should have | and svg replacement
    ' columns 30, 32
    If (InStr(.Cells(n, 39), ".eps") <> 0 And InStr(.Cells(n, 39), "eps export") = 0) Or _
       (InStr(.Cells(n, 39), ".svg") <> 0 And InStr(.Cells(n, 39), "svg replacement") = 0) Then
        .Cells(n, 39).Interior.Color = RGB(255, 0, 0)
        has_err = True
    ElseIf .Cells(n, 32) <> vbNullString Or .Cells(n, 39) <> vbNullString Then
        chk = Replace(.Cells(n, 39), "|eps export", "")
        chk = Replace(chk, "|svg replacement", "")
        chk = Replace(chk, ",", "|")
        If .Cells(n, 32) <> vbNullString And .Cells(n, 39) <> vbNullString Then
            chk = .Cells(n, 32) & "|" & chk
        End If
        sets = verify_image_sets(chk, types, "|", .Cells(n, 4))
        If sets = 0 Then
            .Cells(n, 32).Interior.Color = RGB(255, 0, 0)
            has_err = True
        End If
    End If
    
    
    ' $( QC2 and C2 ) QC3. Multiple graphics for the same piece of art (SVG, EPS, etc) reside in the same Standalone Image element as separate Graphic elements.
    ' ----------------------------Notes------------------------------
    ' check 27 matches number of sets of images in 28
    ' check 24 matches number of sets of images in 25
    
    
    ' QC4. Final videos are in mp4 format and uploaded in the content tab of the item.
    ' ----------------------------Notes------------------------------
    ' %%%%% is there always only 1 video %%%%%
    If InStr(.Cells(n, 34), ".mp4") = 0 And .Cells(n, 34) <> vbNullString Then
        .Cells(n, 34).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' QC5. There are no Order Match, Complex Table, Interactive-Clock Object, Interactive-Scale item types (or item parts in Composite items).
    ' ----------------------------Notes------------------------------
    ' item types is column 15 check that the above is not in 17
    ' %%%%% confirm that the above list is comprehensive to how the list exports %%%%%
    If InStr(.Cells(n, 17), "Order Match") <> 0 Or InStr(.Cells(n, 17), "Complex Table") <> 0 Or InStr(.Cells(n, 17), "Interactive-Clock Object") <> 0 _
    Or InStr(.Cells(n, 17), "Interactive-Scale") <> 0 Then
        .Cells(n, 17).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' #@QC6. No restricted elements are authored in the following item types/item parts: Numeric, Match, Grid, Fill In Blank, Extended Text.
    ' NOTE: Content must follow these restrictions for new items.
    If .Cells(n, 64) = "YES" Then
        .Cells(n, 64).Interior.Color = RGB(255, 0, 0)
        has_err = True
    End If
    
    
    ' &QC7. Check columns 25, 28, 30, 32 to make sure there are no other filetypes than gif, svg, and eps
    chk = .Cells(n, 25) & .Cells(n, 28) & .Cells(n, 30) & .Cells(n, 32)
    art = Split(chk, ".")
    For i = LBound(art) + 1 To UBound(art)
        If Left(Trim(art(i)), 3) <> "svg" And Left(Trim(art(i)), 3) <> "gif" And Left(Trim(art(i)), 3) <> "eps" Then
            If Left(Trim(art(i)), 2) <> "AR" Then
                .Cells(n, 25).Interior.Color = RGB(255, 0, 0)
                .Cells(n, 28).Interior.Color = RGB(255, 0, 0)
                .Cells(n, 30).Interior.Color = RGB(255, 0, 0)
                .Cells(n, 32).Interior.Color = RGB(255, 0, 0)
                has_err = True
            End If
        End If
    Next i


    ' $( QC2 ) QC8. If Image names in columns 25, 28, 30, 32, or 39 contain VH or VF check to make sure the accnum matches for each occurance of VH or VF in image cells
    
    
    If has_err Then .Cells(n, 4).Interior.Color = RGB(255, 0, 0)
End With
Next n
Application.StatusBar = False
End Sub
Function image_match_text(img As String, img_type As String, alt As String, alt_type As String) As Boolean
Dim svg_cnt As Integer
Dim clas_cnt As Integer

image_match_text = False
If alt = vbNullString And img = vbNullString Then
    image_match_text = True
    Exit Function
End If

svg_cnt = (Len(img) - Len(Replace(img, img_type, ""))) / Len(img_type)
clas_cnt = (Len(alt) - Len(Replace(alt, alt_type, ""))) + 1

If svg_cnt = clas_cnt Then
    image_match_text = True
    Exit Function
End If

End Function
Function verify_image_sets(images As String, types() As Variant, split_val As String, accnum As String) As Long
Dim img() As String
Dim img_name As Long
Dim elmt As Variant
Dim img_arr As Variant

If images = vbNullString Then
    verify_image_sets = -1
    Exit Function
End If

img = Split(images, split_val)

' if the number of images is not a multiple of the number of
' types then there is not an image per type.
If (UBound(img) + 1) Mod (UBound(types) + 1) <> 0 Then
    verify_image_sets = 0
    Exit Function
End If

' Group images by primary name
For img_name = LBound(img) To UBound(img)
    For Each elmt In types
        img(img_name) = Replace(img(img_name), CStr(elmt), "")
    Next elmt
Next img_name
'img_arr = eliminateDuplicate(img)

For img_name = LBound(img) To UBound(img)
    If InStr(img(img_name), "VH") <> 0 Or InStr(img(img_name), "VF") <> 0 Then
        If InStr(img(img_name), accnum) = 0 Then
            verify_image_sets = 0
            Exit Function
        End If
    End If
Next img_name
verify_image_sets = LBound(img_arr) + 1
End Function
Function check_classifications(class As String, svg_cnt As Integer) As Boolean
Dim classifications As String

check_classifications = False

If class = vbNullString And svg_cnt = 0 Then
    check_classifications = True
    Exit Function
End If

classifications = Replace(class, "[1]", "")
classifications = Replace(classifications, "[2]", "")
classifications = Replace(classifications, "[3]", "")
If svg_cnt - ((Len(class) - Len(classifications)) / 3) = 0 Then
    check_classifications = True
End If
End Function
