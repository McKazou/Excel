let
    // Récupère le classeur Excel actuel
    Source = Excel.CurrentWorkbook(),
    //ListTable in Excel
    ListTable = Table.FromList(Source[Name], Splitter.SplitByNothing(), {"Name"}),
    // Récupère le tableau "StartingPoint"
    GetTopAssyName = Source{[Name="StartingPoint"]}[Content],
    // Récupère la première ligne du tableau "StartingPoint"
    firstRow = GetTopAssyName{0},
    // Récupère le nom de l'assemblage de départ, ou une valeur par défaut si la colonne "StartingPoint" n'existe pas
    TopAssyName = try Record.Field(firstRow, "StartingPoint") otherwise "In_Context",
    // Récupère l'id de l'assemblage de départ, ou 1 si la colonne "id" n'existe pas
    TopAssyId = try Record.Field(firstRow, "id") otherwise 1,
    // Récupère la quantité de l'assemblage de départ, ou 1 si la colonne "Qty" n'existe pas
    TopAssyQty = try Record.Field(firstRow, "Qty") otherwise 1,
    // Définit la profondeur de départ, ou 1 si la colonne "Depth" n'existe pas
    TopAssyDepth = try Record.Field(firstRow, "Depth") otherwise 5,
    // Définit la profondeur de départ, ou 1 si la colonne "Depth" n'existe pas
    TopAssyBranchQty = try Record.Field(firstRow, "BranchQty") otherwise TopAssyQty,
    // Crée le tableau inputTable
    inputTable = #table(type table [id=text, Name=text, Qty=number, Depth=number, BranchQty=number], {{TopAssyId, TopAssyName, TopAssyQty, TopAssyDepth,TopAssyBranchQty}}),
    // Appelle la fonction récursive avec le tableau inputTable
    Result = AssemblyRecursion(inputTable,Source),
    // Inverse les valeurs de la colonne "Depth"
    MaxDepth = List.Max(Result[Depth]),
    #"Inverted Depth" = Table.TransformColumns(Result, {"Depth", each MaxDepth - _, type number}),
    #"Added Index" = Table.AddIndexColumn(#"Inverted Depth", "Index", 0, 1, Int64.Type),
    #"Added NextDepth" = Table.AddColumn(#"Added Index", "NextDepth", each try #"Added Index"{[Index]+1}[Depth] otherwise null),
    // Modifie la colonne "ArrowSymbol" pour changer le symbole en fonction de la valeur de "Depth" et "NextDepth", et n'ajoute pas de symbole si "Depth" est 0
    #"Added ArrowSymbol" = Table.AddColumn(#"Added NextDepth", "ArrowSymbol", each if [Depth] = 0 then "" else if [Depth] = [NextDepth] then " ├─> " else " └─> "),
    #"Removed Columns" = Table.RemoveColumns(#"Added ArrowSymbol",{"NextDepth"}),
    // Modifie la colonne "RepeatedSymbol" pour répéter " ┃   " en fonction de la valeur de "Depth", et n'ajoute pas de "tabulation" si "Depth" est 0
    #"Added RepeatedSymbol" = Table.AddColumn(#"Removed Columns", "RepeatedSymbol", each if [Depth] = 0 then "" else Text.Repeat(" │   ", [Depth]-1)),
    // Fusionne les colonnes "RepeatedSymbol" et "ArrowSymbol" pour créer la colonne "Arrow"
    #"Added Arrow Column" = Table.CombineColumns(#"Added RepeatedSymbol",{"RepeatedSymbol", "ArrowSymbol"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Arrow"),    #"Merged Columns" = Table.CombineColumns(#"Added Arrow Column",{"Arrow", "Name"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Name"),
    #"Reordered Columns" = Table.ReorderColumns(#"Merged Columns",{"Depth", "id", "Name", "Qty", "BranchQty"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Reordered Columns",{"Index", "Depth", "id", "Name", "Qty", "BranchQty"})
in
    #"Reordered Columns1"