// Get input array from a specific xlsx
let
    Source = Excel.Workbook(File.Contents("Y:\01_Component Librairies\70_NdC et Tolérances\Battery Pack\Bounding Box\Master-Tool_Bounding-Box_240313.xlsm"), null, true),
    Input_Table = Source{[Item="Input",Kind="Table"]}[Data],
    #"Type modifié" = Table.TransformColumnTypes(Input_Table,{{"Input_Name", type any}, {"Input_Value", type any}}),
    inputNames = Text.Combine(#"Type modifié"[Input_Name]," ; ")
in
    inputNames


// Get Output array from a specific xlsx
let
    Source = Excel.Workbook(File.Contents("Y:\01_Component Librairies\70_NdC et Tolérances\Battery Pack\Bounding Box\Master-Tool_Bounding-Box_240313.xlsm"), null, true),
    Input_Table = Source{[Item="Output",Kind="Table"]}[Data],
    inputNames = Text.Combine(Input_Table[Param Name]," ; ")
in
    inputNames


//Function : 
let
    Source = (filePath) =>
    //[INSERT FUNCTION]
in
    Source

