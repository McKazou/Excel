let
    // Récupère le classeur Excel actuel
    Source = Excel.CurrentWorkbook(),
    // Définit le nom de la colonne
    columnName = "Name",
    // Définit la fonction récursive
    recursivefunction = (tableName as text, depth as number) as any =>
        let
            // Essaye de récupérer la table avec le nom donné, sinon retourne null
            CurrentTable = try Source{[Name=tableName]}[Content] otherwise null,
            // Si la table n'existe pas ou si la profondeur est de 0, retourne une table avec le nom et la profondeur
            ApplyFunction = if CurrentTable is null or depth = 0 then
                #table({columnName, "Depth"}, {{tableName, depth}})
            else
                let
                    // Crée une ligne avec le nom et la profondeur
                    CurrentRow = #table({columnName, "Depth"}, {{tableName, depth}}),
                    // Transforme la colonne "Name" en appelant récursivement la fonction avec le nom et la profondeur - 1
                    RecursiveTable = Table.TransformColumns(CurrentTable, {{columnName, each recursivefunction(_, depth - 1), type table}}),
                    // Étend la colonne "Name" pour inclure les colonnes "Name" et "Depth" de la table récursive
                    ExpandedTable = Table.ExpandTableColumn(RecursiveTable, columnName, {columnName, "Depth"}),
                    // Combine la ligne courante et la table étendue
                    ResultTable = Table.Combine({CurrentRow, ExpandedTable})
                in
                    // Retourne la table résultante
                    ResultTable
        in
            // Applique la fonction
            ApplyFunction
in
    // Retourne la fonction récursive
    recursivefunction