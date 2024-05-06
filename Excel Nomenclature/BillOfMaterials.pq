let
    // Récupère le classeur Excel actuel
    Source = Backend_Parts_List_Data,
    #"Removed Columns1" = Table.RemoveColumns(Source,{"TimeStamps"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns1", "Index", 0, 1, Int64.Type),
    // Remove the "Depth" column
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"Depth"}),

    // Group by "Name" and perform necessary operations on other columns
    Grouped = Table.Group(#"Removed Columns", {"Name"}, {
        {"BranchQty", each if List.AllTrue(List.Transform([BranchQty], each Type.Is(Value.Type(_), Number.Type))) then List.Sum([BranchQty]) else "#Text", type any},
        {"Qty", each Text.Combine(List.Transform([Qty], Text.From), ","), type text},
        {"id", each if List.Count(List.Distinct([id])) > 1 then "#errorIdentité" else List.First([id]), type text}

    }),
    #"Added Custom" = Table.AddColumn(Grouped, "Timestamp", each DateTime.LocalNow())
in
    #"Added Custom"