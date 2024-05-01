let
    AssemblyRecursion = (inputTable as table, source as table) as table =>
    let
        result = Table.Buffer(inputTable),
        processRow = (row as record) as table =>
            let
                depth = row[Depth],
                name = row[Name],
                currentTable = try source{[Name=name]}[Content] otherwise null,
                newTable = if currentTable <> null and depth > 0 then
                    let
                        addedDepth = Table.AddColumn(currentTable, "Depth", each depth - 1),
                        recursionResult = AssemblyRecursion(addedDepth,source)
                    in
                        recursionResult
                else
                    #table({"id", "Name", "Qty", "Depth"}, {{row[id], row[Name], row[Qty], row[Depth]}})
            in
                if newTable <> null then newTable else #table(type table [id=text, Name=text, Qty=number, Depth=number], {})
    in
        List.Accumulate(Table.ToRecords(result), #table(type table [id=text, Name=text, Qty=number, Depth=number], {}), (state, current) => Table.Combine({state, processRow(current)}))
in
    AssemblyRecursion