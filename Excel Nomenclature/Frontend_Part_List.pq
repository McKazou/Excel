let
    // Récupère le classeur Excel actuel
    Source = Backend_Parts_List_Data,
    #"Remove Timestamps" = Table.RemoveColumns(Source,{"TimeStamps"}),
    #"Added Index" = Table.AddIndexColumn(#"Remove Timestamps", "Index", 0, 1, Int64.Type),
    #"Added NextDepth" = Table.AddColumn(#"Added Index", "NextDepth", each try #"Added Index"{[Index]+1}[Depth] otherwise null),
    // Modifie la colonne "ArrowSymbol" pour changer le symbole en fonction de la valeur de "Depth" et "NextDepth", et n'ajoute pas de symbole si "Depth" est 0
   #"Added ArrowSymbol" = Table.AddColumn(#"Added NextDepth", "ArrowSymbol", each if [Depth] = 0 then "" else if [Depth] = [NextDepth] then " ├─> " else " └─> "),
    #"Removed Columns" = Table.RemoveColumns(#"Added ArrowSymbol",{"NextDepth"}),
    // Votre code existant...
    #"Added RepeatedSymbol" = Table.AddColumn(#"Removed Columns", "RepeatedSymbol", each if [Depth] = 0 then "" else Text.Repeat(" │  ", [Depth]-1)),
    #"Added Arrow Column" = Table.CombineColumns(#"Added RepeatedSymbol",{"RepeatedSymbol", "ArrowSymbol"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Arrow"),
    #"Merged Columns" = Table.CombineColumns(#"Added Arrow Column",{"Arrow", "Name"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Arrow&Name"),
    #"Reordered Columns" = Table.ReorderColumns(#"Merged Columns",{"Depth", "id", "Arrow&Name", "Qty", "BranchQty"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Reordered Columns",{"Index", "Depth", "id", "Arrow&Name", "Qty", "BranchQty"})
in
    #"Reordered Columns1"