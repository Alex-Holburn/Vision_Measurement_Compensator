Sub Clear_Output_Data()

'Macro Name: Clear_Output_Data
'Author: Alex Holburn https://www.alexholburn.com
'Last Edit: 03/31/21
'Edit Purpose: Better commenting.
'License: MIT License. Copyright 2021, Alex Holburn https://www.alexholburn.com

'Macro Purpose: The purpose of this macro is to clear the simulation output data with a single button click

'    *************************************BEGIN MAIN CODE************************************
        
        'Write Blanks into the Motion Data Output data ranges
        Worksheets("Motion_Data").Range(out_data).Value = ""
        
        'Write Blanks into the Motion Analysis Data Range
        Worksheets("Motion_Analysis").Range("H13:H16").Value = ""

        'Display a Message to the user so they know the analysis was run.
        MsgBox "Simulation Output Data Cleared!"

End Sub