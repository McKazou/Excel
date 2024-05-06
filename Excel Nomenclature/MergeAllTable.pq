let
    // Récupère le classeur Excel actuel
    Source = Excel.CurrentWorkbook(),
    //Obtiens la liste des tableaux généré par l'utilisateur :
    AssemblyTables = Frontend_Assembly_Members,
    // Convertir la colonne de table en liste
    AssemblyTablesList = Table.Column(AssemblyTables, "Assembly Members"),
    // Filtrer les tableaux en fonction de la liste obtenue à partir de "AssemblyTables"
    FilteredRows = Table.SelectRows(Source, each List.Contains(AssemblyTablesList, [Name])),
    #"Expanded Content" = Table.ExpandTableColumn(FilteredRows, "Content", {"id", "Name", "Qty"}, {"Content.id", "Content.Name", "Content.Qty"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Content",{{"Content.Qty", "Qty per assembly"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Content.id"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Name", "Content.Name", "Qty per assembly"}),
    #"Add Timestamps" = Table.AddColumn(#"Reordered Columns", "Timestamp", each DateTime.LocalNow())
in
    #"Add Timestamps"