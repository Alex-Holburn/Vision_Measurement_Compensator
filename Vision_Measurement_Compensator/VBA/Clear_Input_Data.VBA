Sub Clear_Input_Data()

'Macro Name: Clear_Input_Data
'Author: Alex Holburn https://www.alexholburn.com
'Last Edit: 03/31/21
'Edit Purpose: Better commenting.
'License: MIT License. Copyright 2021, Alex Holburn https://www.alexholburn.com

'Macro Purpose: The purpose of this macro is to clear the simulation input data with a single button click

'    *************************************BEGIN MAIN CODE************************************
        
        'Write Blanks into the project info range
        Worksheets("Motion_Analysis").Range("G2:G5").Value = ""
        
        'Write Blanks into the input data range
        Worksheets("Motion_Analysis").Range("B9:16").Value = ""
        Worksheets("Motion_Analysis").Range("H9:12").Value = ""

        'Display a Message to the user so they know the analysis was run.
        MsgBox "Simulation Input Data Cleared!"
    
End Sub