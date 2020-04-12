Attribute VB_Name = "Reinforcement_functions"
Option Explicit

Function stirrup_length(vert_bar_ctrs As Double, horiz_bar_ctrs As Double, stirrup_dia As Double, enclosed_bar_dia As Double, Optional main_bars As Boolean = False, Optional deformed_bars As Boolean = False)

'function to calculate the length of a minimum length rectangular closed stirrup with 135 degree hooks

'vert_bar_ctrs = dimension between centers of enclosed bars in the vertical direction
'horiz_bar_ctrs = dimension between centers of enclosed bars in the horizontical direction
'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'enclosed_bar_dia = diameter of bar which 'stirrup' bar encloses
'main_bars = TRUE ('stirrup' bars are considered main bars for minimum bend diameters), FALSE ('stirrup' bars are considered stirrup or tie bars for minimum bend diameters)
'deformed_bars =TRUE ('stirrup' bars are considered as deformed bars for minimum bend diameters and hook extensions), FALSE ('stirrup' bars are considered as plain round bars for minimum bend diameters and hook extensions)

    Dim pi As Double
    pi = 3.14159265358979

    Dim h    'ctr to ctr height dimension of stirrup hook bends around main bars
    Dim B    'ctr to ctr width dimension of stirrup hook bends around main bars
    Dim r    ' minimum centreline bend radius
    Dim e    'std hook extension

    h = vert_bar_ctrs + enclosed_bar_dia + stirrup_dia
    B = horiz_bar_ctrs + enclosed_bar_dia + stirrup_dia
    r = r_min(stirrup_dia, enclosed_bar_dia, main_bars, deformed_bars)
    e = std_hook_ext(135, stirrup_dia, deformed_bars)
    stirrup_length = 2 * (B - 2 * r) + 2 * (h - 2 * r) + 2 * e + 3 * pi * r

End Function
Function link_length(bar_ctrs As Double, stirrup_dia As Double, enclosed_bar_dia As Double, hook_type As Double, Optional main_bars As Boolean = False, Optional deformed_bars As Boolean = False)

'function to calculate the length of a minimum length link with 135 or 180 degree standard hooks

'bar_ctrs = dimension between centers of enclosed bars
'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'enclosed_bar_dia = diameter of 'stirrup' bar which stirrup encloses
'hook_type = 135 or180 degree standard hook
'main_bars = TRUE ('stirrup' bars are considered main bars for minimum bend diameters), FALSE ('stirrup' bars are considered stirrup or tie bars for minimum bend diameters)
'deformed_bars =TRUE ('stirrup' bars are considered as deformed bars for minimum bend diameters and hook extensions), FALSE ('stirrup' bars are considered as plain round bars for minimum bend diameters and hook extensions)

    Dim pi As Double
    pi = 3.14159265358979

    Dim h    'ctr to ctr dimension of stirrup hook bends around main bars
    Dim r    ' minimum centreline bend radius
    Dim e    'std hook extension
    Dim m    'multiplier for angle

    If hook_type = 135 Then
        m = 3 / 2
    ElseIf hook_type = 180 Then
        m = 2
    Else
        link_length = "135/180 required for hook_type"
        Exit Function
    End If

    h = bar_ctrs + enclosed_bar_dia + stirrup_dia
    r = r_min(stirrup_dia, enclosed_bar_dia, main_bars, deformed_bars)
    e = std_hook_ext(hook_type, stirrup_dia, deformed_bars)
    link_length = (h - 2 * r) + 2 * e + m * pi * r

End Function

Private Function std_hook_ext(hook_type As Double, stirrup_dia As Double, deformed_bars As Boolean)

'function to calculate the minimum straight length extension for a standard hook

'hook_type = 90/135/180 degree standard hook
'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'deformed_bars =TRUE ('stirrup' bars are considered as deformed bars for minimum bend diameters and hook extensions), FALSE ('stirrup' bars are considered as plain round bars for minimum bend diameters and hook extensions)

    Select Case hook_type

    Case 90
        If deformed_bars Then
            std_hook_ext = 12 * stirrup_dia
        Else
            std_hook_ext = 16 * stirrup_dia
        End If

    Case 135
        If deformed_bars Then
            std_hook_ext = 6 * stirrup_dia
        Else
            std_hook_ext = 8 * stirrup_dia
        End If

    Case 180
        std_hook_ext = WorksheetFunction.Max(4 * stirrup_dia, 65)

    End Select

End Function

Private Function r_min(stirrup_dia As Double, enclosed_bar_dia As Double, main_bars As Boolean, deformed_bars As Boolean)

'function to calculate minimum bend radius to stirrup centreline

'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'enclosed_bar_dia = diameter of 'stirrup' bar which stirrup encloses
'main_bars = TRUE ('stirrup' bars are considered main bars for minimum bend diameters), FALSE ('stirrup' bars are considered stirrup or tie bars for minimum bend diameters)
'deformed_bars = TRUE ('stirrup' bars are considered as deformed bars for minimum bend diameters and hook extensions), FALSE ('stirrup' bars are considered as plain round bars for minimum bend diameters and hook extensions)

    Dim d_i    'min bend diameter
    Dim n    'multiplier for deformed bars
    Dim m    'value for minimum bend diameter for main bars or stirrups
    If deformed_bars Then
        n = 2
    Else
        n = 1
    End If

    If main_bars Then
        n = 1
        If stirrup_dia >= 24 Then
            m = 6
        Else
            m = 5
        End If
    Else
        If stirrup_dia >= 24 Then
            m = 3
        Else
            m = 2
        End If
    End If

    d_i = n * m * stirrup_dia

    r_min = WorksheetFunction.Max(enclosed_bar_dia / 2, d_i / 2) + stirrup_dia / 2

End Function

Function generate_stirrup(vert_bar_ctrs As Double, horiz_bar_ctrs As Double, stirrup_dia As Double, enclosed_bar_dia As Double, Optional flip_vertically = False, Optional flip_horizontally = False, Optional main_bars As Boolean = False, Optional deformed_bars As Boolean = False, Optional horizontal_offset = 0, Optional vertical_offset = 0) As Variant

'function to generate coordinates for rectangular stirrup

'vert_bar_ctrs = dimension between centers of enclosed bars in vertical direction
'horiz_bar_ctrs = dimension between centers of enclosed bars in horizontal driection
'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'enclosed_bar_dia = diameter of 'stirrup' bar which stirrup encloses
'horizontal_offset = horizontal distance to offset link
'vertical_offset = vertical distance to offset link

    Dim i
    Dim results

    ReDim coords_x(1 To 19)
    ReDim coords_y(1 To 19)

    Dim r
    Dim e
    Dim v
    Dim h

    r = r_min(stirrup_dia, enclosed_bar_dia, main_bars, deformed_bars)    'stirrup_dia / 2 + enclosed_bar_dia / 2
    e = std_hook_ext(135, stirrup_dia, deformed_bars)
    v = vert_bar_ctrs + enclosed_bar_dia + stirrup_dia - 2 * r
    h = horiz_bar_ctrs + enclosed_bar_dia + stirrup_dia - 2 * r

    'working clockwise
    'pt1
    coords_x(1) = -h / 2 - r / (2 ^ 0.5) + e / (2 ^ 0.5)
    coords_y(1) = v / 2 - r / (2 ^ 0.5) - e / (2 ^ 0.5)
    'pt2
    coords_x(2) = -h / 2 - r / (2 ^ 0.5)
    coords_y(2) = v / 2 - r / (2 ^ 0.5)
    'pt3

    coords_x(3) = -h / 2 - r
    coords_y(3) = v / 2
    'pt4
    coords_x(4) = -h / 2 - r / (2 ^ 0.5)
    coords_y(4) = v / 2 + r / (2 ^ 0.5)
    'pt5
    coords_x(5) = -h / 2
    coords_y(5) = v / 2 + r
    'pt6
    coords_x(6) = h / 2
    coords_y(6) = v / 2 + r
    'pt7
    coords_x(7) = h / 2 + r / (2 ^ 0.5)
    coords_y(7) = v / 2 + r / (2 ^ 0.5)
    'pt8
    coords_x(8) = h / 2 + r
    coords_y(8) = v / 2
    'pt9
    coords_x(9) = h / 2 + r
    coords_y(9) = -v / 2
    'pt10
    coords_x(10) = h / 2 + r / (2 ^ 0.5)
    coords_y(10) = -v / 2 - r / (2 ^ 0.5)
    'pt11
    coords_x(11) = h / 2
    coords_y(11) = -v / 2 - r
    'pt12
    coords_x(12) = -h / 2
    coords_y(12) = -v / 2 - r
    'pt13
    coords_x(13) = -h / 2 - r / (2 ^ 0.5)
    coords_y(13) = -v / 2 - r / (2 ^ 0.5)
    'pt14
    coords_x(14) = -h / 2 - r
    coords_y(14) = -v / 2
    'pt15
    coords_x(15) = -h / 2 - r
    coords_y(15) = v / 2
    'pt16
    coords_x(16) = -h / 2 - r / (2 ^ 0.5)
    coords_y(16) = v / 2 + r / (2 ^ 0.5)
    'pt17
    coords_x(17) = -h / 2
    coords_y(17) = v / 2 + r
    'pt18
    coords_x(18) = -h / 2 + r / (2 ^ 0.5)
    coords_y(18) = v / 2 + r / (2 ^ 0.5)
    'pt19
    coords_x(19) = -h / 2 + r / (2 ^ 0.5) + e / (2 ^ 0.5)
    coords_y(19) = v / 2 + r / (2 ^ 0.5) - e / (2 ^ 0.5)

    'swap hook orientation/direction if specified, by default hooks in top left
    If flip_vertically Then
        For i = 1 To UBound(coords_y)
            coords_y(i) = -coords_y(i)
        Next i
    End If

    If flip_horizontally Then
        For i = 1 To UBound(coords_x)
            coords_x(i) = -coords_x(i)
        Next i
    End If

    'assemble final array
    ReDim results(1 To UBound(coords_x), 1 To 2)

    For i = 1 To UBound(coords_x)
        results(i, 1) = coords_x(i) + horizontal_offset
        results(i, 2) = coords_y(i) + vertical_offset
    Next i

    'write results
    generate_stirrup = results

End Function

Function generate_link(bar_ctrs As Double, stirrup_dia As Double, enclosed_bar_dia As Double, hook_type As Double, vertical_link As Boolean, hook_orientation As Boolean, _
                       Optional main_bars As Boolean = False, Optional deformed_bars As Boolean = False, Optional horizontal_offset = 0, Optional vertical_offset = 0) As Variant

'function to generate coordinates for link stirrup

'bar_ctrs = dimension between centers of enclosed bars
'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'enclosed_bar_dia = diameter of 'stirrup' bar which stirrup encloses
'hook_type = 135 or 180 degree standard hook
'vertical_link = TRUE (vertical link), FALSE (horizontal link)
'hook_orientation =
'main_bars = TRUE ('stirrup' bars are considered main bars for minimum bend diameters), FALSE ('stirrup' bars are considered stirrup or tie bars for minimum bend diameters)
'deformed_bars = TRUE ('stirrup' bars are considered as deformed bars for minimum bend diameters and hook extensions), FALSE ('stirrup' bars are considered as plain round bars for minimum bend diameters and hook extensions)
'horizontal_offset = horizontal distance to offset link
'vertical_offset = vertical distance to offset link

'consider adding end offset to move hooks inwards, could simply change c to be smaller and follow normal link generation

    Dim pi As Double
    pi = 3.14159265358979
    Dim results
    Dim temp_coords
    Dim i

    Dim count
    If hook_type = 135 Then
        count = 10
    Else
        count = 12
    End If
    ReDim coords_x(1 To count)
    ReDim coords_y(1 To count)

    Dim r
    Dim e
    Dim v
    Dim C

        r = r_min(stirrup_dia, enclosed_bar_dia, main_bars, deformed_bars)
        C = bar_ctrs + enclosed_bar_dia + stirrup_dia - 2 * r

    e = std_hook_ext(hook_type, stirrup_dia, deformed_bars)

    'working clockwise with vertical link with hooks to left
    If hook_type = 135 Then
        'pt1
        coords_x(1) = -r / (2 ^ 0.5) - e / (2 ^ 0.5)
        coords_y(1) = C / 2 + r / (2 ^ 0.5) - e / (2 ^ 0.5)
        'pt2
        coords_x(2) = -r / (2 ^ 0.5)
        coords_y(2) = C / 2 + r / (2 ^ 0.5)
        'pt3
        coords_x(3) = 0
        coords_y(3) = C / 2 + r
        'pt4
        coords_x(4) = r / (2 ^ 0.5)
        coords_y(4) = C / 2 + r / (2 ^ 0.5)
        'pt5
        coords_x(5) = r
        coords_y(5) = C / 2
        'pt6
        coords_x(6) = r
        coords_y(6) = -C / 2
        'pt7
        coords_x(7) = r / (2 ^ 0.5)
        coords_y(7) = -C / 2 - r / (2 ^ 0.5)
        'pt8
        coords_x(8) = 0
        coords_y(8) = -C / 2 - r
        'pt9
        coords_x(9) = -r / (2 ^ 0.5)
        coords_y(9) = -C / 2 - r / (2 ^ 0.5)
        'pt10
        coords_x(10) = -r / (2 ^ 0.5) - e / (2 ^ 0.5)
        coords_y(10) = -C / 2 - r / (2 ^ 0.5) + e / (2 ^ 0.5)
    ElseIf hook_type = 180 Then
        'pt1
        coords_x(1) = -r
        coords_y(1) = C / 2 - e
        'pt2
        coords_x(2) = -r
        coords_y(2) = C / 2
        'pt3
        coords_x(3) = -r / (2 ^ 0.5)
        coords_y(3) = C / 2 + r / (2 ^ 0.5)
        'pt4
        coords_x(4) = 0
        coords_y(4) = C / 2 + r
        'pt5
        coords_x(5) = r / (2 ^ 0.5)
        coords_y(5) = C / 2 + r / (2 ^ 0.5)
        'pt6
        coords_x(6) = r
        coords_y(6) = C / 2
        'pt7
        coords_x(7) = r
        coords_y(7) = -C / 2
        'pt8
        coords_x(8) = r / (2 ^ 0.5)
        coords_y(8) = -C / 2 - r / (2 ^ 0.5)
        'pt9
        coords_x(9) = 0
        coords_y(9) = -C / 2 - r
        'pt10
        coords_x(10) = -r / (2 ^ 0.5)
        coords_y(10) = -C / 2 - r / (2 ^ 0.5)
        'pt11
        coords_x(11) = -r
        coords_y(11) = -C / 2
        'pt12
        coords_x(12) = -r
        coords_y(12) = -C / 2 + e
    Else
        generate_link = "135/180 required for hook_type"
        Exit Function
    End If

    'swap hook orientation/direction if specified
    If hook_orientation Then
        For i = 1 To UBound(coords_x)
            coords_x(i) = -coords_x(i)
        Next i
    End If

    'swap link orientation if required between vertical and horizontal
    If Not vertical_link Then
        For i = 1 To UBound(coords_x)
            temp_coords = coords_x(i)
            coords_x(i) = coords_y(i)
            coords_y(i) = temp_coords
        Next i
    End If

    'assemble final array
    ReDim results(1 To UBound(coords_x), 1 To 2)

    For i = 1 To UBound(coords_x)
        results(i, 1) = coords_x(i) + horizontal_offset
        results(i, 2) = coords_y(i) + vertical_offset
    Next i

    'write results
    generate_link = results

End Function

Function generate_stirrup_set(set_type_vert As Integer, set_type_horiz As Integer, vert_bar_ctrs As Double, horiz_bar_ctrs As Double, stirrup_dia As Double, _
                              enclosed_bar_dia As Double, hook_type As Double, number_legs_vert As Integer, number_legs_horiz As Integer, _
                              Optional flip_vertically = False, Optional flip_horizontally = False, _
                              Optional deformed_bars As Boolean = False, _
                              Optional horizontal_offset = 0, Optional vertical_offset = 0) As Variant

'function to generate coordinates for full rectangular stirrup set using plain round or deformed bars
'set_type_horiz & set_type_vert:-
' type '1' = full outer stirrup + individual links in given direction
' type '2' = full outer stirrup + smaller full stirrups + link if required in given direction
' type '3' = full outer stirrup + full overlapping stirrups + link if required in given direction
'vert_bar_ctrs = dimension between centers of outermost enclosed bars in vertical direction
'horiz_bar_ctrs = dimension between centers of outermost enclosed bars in horizontal driection
'stirrup_dia = diameter of 'stirrup' bar for which length is being calculated
'enclosed_bar_dia = diameter of 'stirrup' bar which stirrup encloses
'hook_type = 135 or 180 degree standard hook for links
'number_legs_vert = number of full depth vertical legs
'number_legs_horiz = number of full width horizontal legs
'flip_vertically = (TRUE/FALSE) mirror stirrup vertically to reorientate stirrup hook location
'flip_vertically = (TRUE/FALSE) mirror stirrup horizontally to reorientate stirrup hook location
'deformed_bars = TRUE ('stirrup' bars are considered as deformed bars for minimum bend diameters and hook extensions), FALSE ('stirrup' bars are considered as plain round bars for minimum bend diameters and hook extensions)
'horizontal_offset = horizontal distance to offset link
'vertical_offset = vertical distance to offset link
    Dim external
    Dim internal

    Dim results()
    ReDim results(1 To 1, 1 To 2)
    results(1, 1) = CVErr(xlErrNA)
    results(1, 2) = CVErr(xlErrNA)
    Dim spacer_array
    spacer_array = results
    Dim i As Integer
    Dim n As Double
    Dim horiz_link_location As Double
    Dim vert_link_location As Double
    Dim number_internal_legs_vert As Integer
    Dim number_internal_legs_horiz As Integer
    Dim number_internal_stirrups_vert As Integer
    Dim number_internal_links_vert As Integer
    Dim number_internal_stirrups_horiz As Integer
    Dim number_internal_links_horiz As Integer

    Dim hook_orientation As Boolean
    Dim vertical_link As Boolean
    Dim main_bars As Boolean
    main_bars = False

    'create external stirrup
    external = generate_stirrup(vert_bar_ctrs, horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
    results = CombineArrays(external, spacer_array)

    '_______________________________________________________
    'type '1' = full outer stirrup + individual links
    If set_type_vert = 1 Then
        'create internal vertical links
        vertical_link = True

        If number_legs_vert > 2 Then
            number_internal_legs_vert = number_legs_vert - 2

            For i = 1 To number_internal_legs_vert
                If number_internal_legs_vert = 1 Then
                    horiz_link_location = horiz_bar_ctrs / 2
                Else
                    horiz_link_location = i * (horiz_bar_ctrs / (number_legs_vert - 1))
                End If

                If -horiz_bar_ctrs / 2 + horiz_link_location > 0 Then
                    hook_orientation = False
                Else
                    hook_orientation = True
                End If

                internal = generate_link(vert_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset - horiz_bar_ctrs / 2 + horiz_link_location, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            Next i
        End If
    End If

    If set_type_horiz = 1 Then
        'create internal horizontal links
        vertical_link = False

        If number_legs_horiz > 2 Then
            number_internal_legs_horiz = number_legs_horiz - 2

            For i = 1 To number_internal_legs_horiz
                If number_internal_legs_horiz = 1 Then
                    vert_link_location = vert_bar_ctrs / 2
                Else
                    vert_link_location = i * (vert_bar_ctrs / (number_legs_horiz - 1))
                End If

                If -vert_bar_ctrs / 2 + vert_link_location > 0 Then
                    hook_orientation = False
                Else
                    hook_orientation = True
                End If

                internal = generate_link(horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset - vert_bar_ctrs / 2 + vert_link_location)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            Next i
        End If
    End If
    '_______________________________________________________
    'type '2' = full outer stirrup + smaller full stirrups + link if required

    If set_type_vert = 2 Then
        'create vertical stirrups
        hook_orientation = Not flip_horizontally
        vertical_link = True

        number_internal_legs_vert = number_legs_vert - 2
        If number_internal_legs_vert Mod 2 = 0 Then
            number_internal_stirrups_vert = CInt(number_internal_legs_vert / 2)    ' this needs to be changed so it works out correct number for both odd and even values
        Else
            number_internal_stirrups_vert = CInt(number_internal_legs_vert \ 2)
        End If
        number_internal_links_vert = number_internal_legs_vert Mod 2
        If number_internal_links_vert = 1 Then
            If number_internal_stirrups_vert Mod 2 Then
                'generate double width internal stirrup + link
                internal = generate_stirrup(vert_bar_ctrs, 2 * horiz_bar_ctrs / (number_legs_vert - 1), stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)

                internal = generate_link(vert_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)

                number_internal_stirrups_vert = number_internal_stirrups_vert - 1
            Else
                'generate central internal link
                internal = generate_link(vert_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)

            End If
        Else
            If number_internal_stirrups_vert Mod 2 Then
                'generate normal sized internal stirrup
                internal = generate_stirrup(vert_bar_ctrs, horiz_bar_ctrs / (number_legs_vert - 1), stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
                number_internal_stirrups_vert = number_internal_stirrups_vert - 1
            End If
        End If
        'generate even number side stirrups
        horiz_link_location = -0.5 * (horiz_bar_ctrs / (number_legs_vert - 1))
        For i = 1 To number_internal_stirrups_vert / 2
            If number_internal_stirrups_vert = 1 Then
                horiz_link_location = horiz_bar_ctrs / 2
            Else
                horiz_link_location = horiz_link_location + 2 * (horiz_bar_ctrs / (number_legs_vert - 1))
            End If

            If -horiz_bar_ctrs / 2 + horiz_link_location > 0 Then
                hook_orientation = False
            Else
                hook_orientation = True
            End If

            internal = generate_stirrup(vert_bar_ctrs, horiz_bar_ctrs / (number_legs_vert - 1), stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset - horiz_bar_ctrs / 2 + horiz_link_location, vertical_offset)
            'generate_link(vert_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, True, hook_orientation, False, False, horizontal_offset - horiz_bar_ctrs / 2 + horiz_link_location, vertical_offset)
            internal = CombineArrays(internal, spacer_array)
            results = CombineArrays(results, internal)
            internal = generate_stirrup(vert_bar_ctrs, horiz_bar_ctrs / (number_legs_vert - 1), stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset + horiz_bar_ctrs / 2 - horiz_link_location, vertical_offset)
            'generate_link(vert_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, True, hook_orientation, False, False, horizontal_offset - horiz_bar_ctrs / 2 + horiz_link_location, vertical_offset)
            internal = CombineArrays(internal, spacer_array)
            results = CombineArrays(results, internal)
        Next i
    End If

    If set_type_horiz = 2 Then
        'create horizontal stirrups
        hook_orientation = flip_vertically
         vertical_link = False

        number_internal_legs_horiz = number_legs_horiz - 2
        If number_internal_legs_horiz Mod 2 = 0 Then
            number_internal_stirrups_horiz = CInt(number_internal_legs_horiz / 2)    ' this needs to be changed so it works out correct number for both odd and even values
        Else
            number_internal_stirrups_horiz = CInt(number_internal_legs_horiz \ 2)
        End If
        number_internal_links_horiz = number_internal_legs_horiz Mod 2
        If number_internal_links_horiz = 1 Then
            If number_internal_stirrups_horiz Mod 2 Then
                'generate double width internal stirrup + link
                internal = generate_stirrup(2 * vert_bar_ctrs / (number_legs_horiz - 1), horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)

                internal = generate_link(horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)

                number_internal_stirrups_horiz = number_internal_stirrups_horiz - 1
            Else
                'generate central internal link
                internal = generate_link(horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            End If
        Else
            If number_internal_stirrups_horiz Mod 2 Then
                'generate normal sized internal stirrup
                internal = generate_stirrup(vert_bar_ctrs / (number_legs_horiz - 1), horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
                number_internal_stirrups_horiz = number_internal_stirrups_horiz - 1
            End If
        End If
        'generate even number top/bottom stirrups
        vert_link_location = -0.5 * (vert_bar_ctrs / (number_legs_horiz - 1))
        For i = 1 To number_internal_stirrups_horiz / 2
            If number_internal_stirrups_horiz = 1 Then
                vert_link_location = vert_bar_ctrs / 2
            Else
                vert_link_location = vert_link_location + 2 * (vert_bar_ctrs / (number_legs_horiz - 1))
            End If

            If -vert_bar_ctrs / 2 + vert_link_location > 0 Then
                hook_orientation = False
            Else
                hook_orientation = True
            End If

            internal = generate_stirrup(vert_bar_ctrs / (number_legs_horiz - 1), horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset - vert_bar_ctrs / 2 + vert_link_location)
            internal = CombineArrays(internal, spacer_array)
            results = CombineArrays(results, internal)
            internal = generate_stirrup(vert_bar_ctrs / (number_legs_horiz - 1), horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset + vert_bar_ctrs / 2 - vert_link_location)
            internal = CombineArrays(internal, spacer_array)
            results = CombineArrays(results, internal)
        Next i
    End If
    '_______________________________________________________
    'type '3' = full outer stirrup + full overlapping stirrups + link if required

    If set_type_vert = 3 Then
        'create vertical stirrups
        hook_orientation = Not flip_horizontally
        vertical_link = True

        number_internal_legs_vert = number_legs_vert - 2
        If number_internal_legs_vert Mod 2 = 0 Then
            number_internal_stirrups_vert = CInt(number_internal_legs_vert / 2)    ' this needs to be changed so it works out correct number for both odd and even values
        Else
            number_internal_stirrups_vert = CInt(number_internal_legs_vert \ 2)
        End If
        number_internal_links_vert = number_internal_legs_vert Mod 2
        If number_internal_links_vert = 1 Then
            'generate central internal link
            internal = generate_link(vert_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset)
            internal = CombineArrays(internal, spacer_array)
            results = CombineArrays(results, internal)
            'generate full stirrups
            For i = 1 To number_internal_stirrups_vert
                internal = generate_stirrup(vert_bar_ctrs, 2 * i * horiz_bar_ctrs / (number_legs_vert - 1), stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            Next i
        Else
            'generate full stirrups with no link
            For n = 0.5 To number_internal_stirrups_vert Step 1
                internal = generate_stirrup(vert_bar_ctrs, 2 * n * horiz_bar_ctrs / (number_legs_vert - 1), stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            Next n
        End If
    End If

    If set_type_horiz = 3 Then
        'create horizontal stirrups
        hook_orientation = flip_vertically
        vertical_link = False

        number_internal_legs_horiz = number_legs_horiz - 2
        If number_internal_legs_horiz Mod 2 = 0 Then
            number_internal_stirrups_horiz = CInt(number_internal_legs_horiz / 2)    ' this needs to be changed so it works out correct number for both odd and even values
        Else
            number_internal_stirrups_horiz = CInt(number_internal_legs_horiz \ 2)
        End If
        number_internal_links_horiz = number_internal_legs_horiz Mod 2
        If number_internal_links_horiz = 1 Then
            'generate central internal link
            internal = generate_link(horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, hook_type, vertical_link, hook_orientation, main_bars, deformed_bars, horizontal_offset, vertical_offset)
            internal = CombineArrays(internal, spacer_array)
            results = CombineArrays(results, internal)
            'generate full stirrups
            For i = 1 To number_internal_stirrups_horiz
                internal = generate_stirrup(2 * i * vert_bar_ctrs / (number_legs_horiz - 1), horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            Next i
        Else
            'generate full stirrups with no link
            For n = 0.5 To number_internal_stirrups_horiz Step 1
                internal = generate_stirrup(2 * n * vert_bar_ctrs / (number_legs_horiz - 1), horiz_bar_ctrs, stirrup_dia, enclosed_bar_dia, flip_vertically, flip_horizontally, main_bars, deformed_bars, horizontal_offset, vertical_offset)
                internal = CombineArrays(internal, spacer_array)
                results = CombineArrays(results, internal)
            Next n
        End If
    End If

    'return results as dynamic array
    generate_stirrup_set = results
End Function

Function CombineArrays(A As Variant, B As Variant, Optional stacked As Boolean = True) As Variant
'assumes that A and B are 2-dimensional variant arrays
'if stacked is true then A is placed on top of B
'in this case the number of rows must be the same,
'otherwise they are placed side by side A|B
'in which case the number of columns are the same
'LBound can be anything but is assumed to be
'the same for A and B (in both dimensions)
'False is returned if a clash

    Dim lb As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long, k As Long
    Dim C As Variant

    If TypeName(A) = "Range" Then A = A.Value
    If TypeName(B) = "Range" Then B = B.Value

    lb = LBound(A, 1)
    m_A = UBound(A, 1)
    n_A = UBound(A, 2)
    m_B = UBound(B, 1)
    n_B = UBound(B, 2)

    If stacked Then
        m = m_A + m_B + 1 - lb
        n = n_A
        If n_B <> n Then
            CombineArrays = False
            Exit Function
        End If
    Else
        m = m_A
        If m_B <> m Then
            CombineArrays = False
            Exit Function
        End If
        n = n_A + n_B + 1 - lb
    End If
    ReDim C(lb To m, lb To n)
    For i = lb To m
        For j = lb To n
            If stacked Then
                If i <= m_A Then
                    C(i, j) = A(i, j)
                Else
                    C(i, j) = B(lb + i - m_A - 1, j)
                End If
            Else
                If j <= n_A Then
                    C(i, j) = A(i, j)
                Else
                    C(i, j) = B(i, lb + j - n_A - 1)
                End If
            End If
        Next j
    Next i
    CombineArrays = C
End Function
