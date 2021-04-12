Sub Generate_Code()

'Macro Name: Generate_Code
'Author: Alex Holburn https://www.alexholburn.com
'Last Edit: 03/31/21
'Edit Purpose: Better commenting.
'License: MIT License. Copyright 2021, Alex Holburn https://www.alexholburn.com

'Macro Purpose: The purpose of this macro is to rapidly generate code and tag information for import
'               into a Studio 5000 project. The utility was created to save time when retrofitting 75
'               machines to avoid the use of a true 3D Vision System. The principle behind why this works
'               is the intercept theoreom. This theroem can only be applied if the distortion of the lens.
'               has been calibrated out.
'    *******************************BEGIN VARIABLE DECLARATION*******************************

    Dim cal_target_distance As Double        'Distance of the calibration target from the distance sensor.
    Dim is_target_calibrated As Variant      'Yes/No input of whether the system has been calibrated.
    Dim is_transformed As Variant            'Yes/No input of whether the lens distortion has been transformed.
    Dim is_distance_calibrated As Variant    'Yes/No input of whether the distance sensor has been calibrated.

    Dim cal_target_distance_tag As Variant   'Calibration Target Distance Tag
    Dim apparent_dimension_tag As Variant    'Apparent Distance Tag
    Dim part_distance_tag As Variant         'Part Distance Tag
    Dim comp_dimension_tag As Variant        'Compensated Dimension Tag
    
    Dim apparent_dimension As Double        'Tolerance check calculator apparent dimension
    Dim part_distance As Double             'Tolerance check calculator part distance
    Dim acceptable_tolerance As Double      'Tolerance check acceptable tolerance %
    Dim actual_dimension As Double          'Tolerance check actual part dimension
    
    Dim compensated_dimension As Double     'Tolerance check compensated dimension
    Dim upper_tolerance As Double           'Tolerance check upper tolerance
    Dim lower_tolerance As Double           'Tolerance check lower tolerance
    Dim within_tolerance As Double          'Tolerance check Yes/No within tolerance spec
    
    Dim rung_1 As Variant                    'Rung 1 output code
    Dim rung_2 As Variant                    'Rung 2 output code
    Dim rung_txt_file As String              'Path to the output rung text file.
    Dim Output_Template As String            'Path to tag import CSV Template
    Dim Output_filename As String            'Path to the tag import CSV output file.
    

    

'    *************************************BEGIN MAIN CODE************************************
    
        'Set Calibration Target Distance
        cal_target_distance = Sheets("User_Inputs").Range("B9").Value
        
        'Set User Yes/No Checklist Variables
        is_target_calibrated = Sheets("User_Inputs").Range("B10").Value
        is_transformed = Sheets("User_Inputs").Range("B11").Value
        is_distance_calibrated = Sheets("User_Inputs").Range("B12").Value
        
        'Set The Tag Data Varaibles
        cal_target_distance_tag = Sheets("User_Inputs").Range("B13").Value
        apparent_dimension_tag = Sheets("User_Inputs").Range("B14").Value
        part_distance_tag = Sheets("User_Inputs").Range("B15").Value
        comp_dimension_tag = Sheets("User_Inputs").Range("B16").Value
        
        'Set Calibration Calculator inputs
        apparent_dimension = Sheets("User_Inputs").Range("H9").Value
        part_distance = Sheets("User_Inputs").Range("H10").Value
        acceptable_tolerance = Sheets("User_Inputs").Range("H11").Value
        actual_dimension = Sheets("User_Inputs").Range("H12").Value
        
        'Set the workbook output paths
        Output_Template = Worksheets("Tool_Setup").Range("H6")
        Output_filename = Worksheets("Tool_Setup").Range("H5")

        'Set simulation output data range
        Set out_data = Worksheets("Tool_Setup").Range("E6")
        
        If is_target_calibrated = "YES" And is_transformed = "YES" And is_distance_calibrated = "YES" Then
            'Calculate Compensated Dimension
            Worksheets("User_Inputs").Range("H13") = (apparent_dimension / part_distance) * cal_target_distance
            
            'Calculate Upper tolerance
            Worksheets("User_Inputs").Range("H14") = actual_dimension + ((actual_dimension / 100) * acceptable_tolerance)

            'Calculate Lower tolerance
            Worksheets("User_Inputs").Range("H15") = actual_dimension - ((actual_dimension / 100) * acceptable_tolerance)
            
            'Determine if Apparent Dimension is within tolerance of actual size
            If Worksheets("User_Inputs").Range("H13") >= Worksheets("User_Inputs").Range("H15") And Worksheets("User_Inputs").Range("H13") <= Worksheets("User_Inputs").Range("H14") Then
                Worksheets("User_Inputs").Range("H16") = "YES"
            
            Else
                Worksheets("User_Inputs").Range("H16") = "NO"
            
            End If
            
            'Generate Output Rung #1
            rung_1 = "MOV(" & cal_target_distance & "," & cal_target_distance_tag & ");"
            
            'Generate Output Rung #2
            rung_2 = "CPT(" & comp_dimension_tag & "," & "(" & apparent_dimension_tag & "/" & part_distance_tag & ")*" & cal_target_distance_tag & ");"
        
            'Write The Rung Text File
            rung_text_file = Worksheets("Tool_Setup").Range("H5")
        
            Open rung_text_file For Output As #1
                Print #1, rung_1
                Print #1, rung_2
            Close #1
            
            'Create the Tag Import CSV
            Workbooks.Add (Output_Template)
            
            'Set Type
            Workbooks(2).Worksheets("Tag_Import_Template").Range("A8:A11") = "TAG"
            
            'Set Datatype
            Workbooks(1).Worksheets("User_Inputs").Range("C13:C16").Copy _
                Workbooks(2).Worksheets("Tag_Import_Template").Range("E8")
            
            'Set Attributes
            Workbooks(1).Worksheets("Tool_Setup").Range("H7").Copy _
                Workbooks(2).Worksheets("Tag_Import_Template").Range("G8:G11")
    
            
            'Copy the tag names to the template csv.
            Workbooks(1).Worksheets("User_Inputs").Range("B13:B16").Copy _
                Workbooks(2).Worksheets("Tag_Import_Template").Range("C8")
                
            'Save the template workbook as the output filename.
            Output_filename = Workbooks(1).Worksheets("Tool_Setup").Range("H4")
            Workbooks(2).SaveAs (Output_filename)
            
            'Close the output data workbook
            Workbooks(2).Close
        
            'Display a message to the user so they know the tool worked.
            MsgBox "Rung Text and Tag Import Files Generated!"
            
        Else
            'Display a Message to the user so they know an error occurred.
            MsgBox "All systems must be calibrated and/or Transformed, for this code to this work!"
 
        End If

    
                
End Sub