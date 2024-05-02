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
    // Remove the "Depth" column
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"Depth"}),

    // Group by "Name" and perform necessary operations on other columns
    Grouped = Table.Group(#"Removed Columns", {"Name"}, {
        {"BranchQty", each if List.AllTrue(List.Transform([BranchQty], each Type.Is(Value.Type(_), Number.Type))) then List.Sum([BranchQty]) else "#Text", type any},
        {"Qty", each Text.Combine(List.Transform([Qty], Text.From), ","), type text},
        {"id", each if List.Count(List.Distinct([id])) > 1 then "#errorIdentité" else List.First([id]), type text}

    }),

    #"Reordered Columns" = Table.ReorderColumns(Grouped,{"id", "Name", "Qty", "BranchQty"})
in
    #"Reordered Columns"