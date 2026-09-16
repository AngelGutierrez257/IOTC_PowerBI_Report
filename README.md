# IOTC Catch and Effort Analysis - Power BI Semantic Model and Report

## Synopsis:

This report was developed with the purpose of applying the skills acquired through my self-directed study for the PL_300 certification to a dataset that relates to my experience as a fisheries observer in the IOTC Regulatory Area.

The main objective is to analyse catch and economic data (effort, fuel price and catch price of sale per kgs) to determine efficiency and quota sustainability by fleet, by main commercial species and by gear employed. For model performance and for practicality, the analysis was limited to the period between 2000 and 2024.

For self-study, I asked Claude to prepare manuals for Dax and Power Query based on the Freidman technique and organize them based on the old Skillpipe manuals, in order to prepare for the exam. The visualization and administration skills were studied trough the Microsoft Learning Path (https://learn.microsoft.com/en-us/credentials/certifications/data-analyst-associate/?practice-assessment-type=certification) and the book PL-300 Exam Ref, of author Daniil Maslyuk.

The Claude Manuals are available for download in main.


## Repository structure:

- IOTC/ contains the semantic model and report json files.
- assets/ contains illustrative images of each page of the report and of the semantic model.
- README.md contains the project explanation.

## 1. Data Sources:

### Yearly Catch Data:

  - Original Title: IOTC-DATASETS-2026-05-27-RC-SCI-1950-2024.csv
  - Model Title: Catch_Estimates
  - Origin: https://iotc.org/data/datasets/Retained catches by year, main IOTC area, fleet, and gear for all IOTC and bycatch species
  - Connection type: Csv with with parameter RC-SCI_1950-2024_PATH

    ```powerquery
    Fonte = Csv.Document(File.Contents(#"RC-SCI_1950-2024_PATH"))
    ```

### Yearly fleet statistics:

  - Original Title: Fishing_Craft_Statistics_20250702.zip
  - Model Title: Fleet_Statistics
  - Origin: https://iotc.org/data/datasets/Annual Number of Vessels by Fishing Fleet, Gear, Architecture, Mechanisation type, Size class, and Fish Preservation Method (Fishing Craft Statistics)
  - Connector type: Csv with parameter Conexao Pasta as Folder Connector.

     ```powerquery
     Fonte = Csv.Document(File.Contents(#"Conexao Pasta" & "\IOTC-DATASETS-2026-02-23-CE-1952-2024.csv"))
     ```

### Monthly Fishing Effort

  - Original Title: IOTC-DATASETS-2026-02-23-CE-1952-2024.csv
  - Model Title: Fishing_Effort
  - Origin: https://iotc.org/data/datasets/latest/CE/All
  - Connector type: Csv with parameter parameter Conexao Pasta as Folder Connector.

     ```powerquery
     Fonte = Csv.Document(File.Contents(#"Conexao Pasta" & "\IOTC-DATASETS-2026-02-23-CE-1952-2024.csv")),
     ```

### Tuna Import Prices

  - Original Title: FFA_import_price_tuna_time_series.xlsx
  - Model Title: Tuna_Import_Prices
  - Origin: https://iotc.org/data/datasets/latest/SD/TUNAS
  - Connector type: Csv with parameter Fish_Prices_Excel_Path.

    ```powerquery
    Fonte = Excel.Workbook(File.Contents(#"Fish_Prices_Excel_Path")),
    ```

### Crude Prices

  - Original Title: FFA_crude_oil_price_time_series.xlsx
  - Model Title: Crude_Prices
  - Origin: https://iotc.org/data/datasets/latest/SD/FUEL
  -  Connector type: Csv with parameter  Crude_Oil_Prices_Excel_Path.

      ```powerquery
      Fonte = Excel.Workbook(File.Contents(Crude_Oil_Prices_Excel_Path)),
      ```

### IOTC Main Areas (Geographical data)

  - Original Title: IOTC_MAIN_AREAS_10.0.0.csv
  - Model Title: IOTC_Major_Geo_Areas
  - Origin: https://data.iotc.org/reference/latest/domain/admin/#geospatialData
  - Connector type: Csv with parameter Geographical_Data_Connection.

     ```powerquery
     Fonte = Csv.Document(File.Contents(Geographical_Data_Connection)),
     ```


The use of parameters is preferred has it handles changes in source locations more gracefully and easily and also allow the use of deployment pipelines.


## 2 Data Preparation - Power Query:

  Mai tasks: Data normalization, null handling, prepare data for modelling, filter and define data types, rows and columns relevant for the model.

### Power Query Parameter:

  Besides the connectors, a parameter was defined to uniformly select rows >= the year 2000.
  Parameter Name: StartYear



## Fact Tables:


### Catch_Estimates (staging, not imported to the model)

1. Basic transformations: Promote headers, Remove Spaces, alter data types, column title uniformization:

      ``` powerquery
      Fonte = Csv.Document(File.Contents(#"RC-SCI_1950-2024_PATH")),
    PromoverNomes = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),
    RemoverEspaços = Table.TransformColumns(PromoverNomes, {}, Text.Trim),
    AlterarTipos =  Table.TransformColumnTypes(RemoverEspaços, {{"YEAR", type number}, {"FISHING_GROUND_CODE", type text}, {"FISHING_GROUND", type text}, {"FLEET_CODE", type text}, {"FLEET", type text},   {"FISHERY_TYPE_CODE", type text}, {"FISHERY_TYPE", type text}, {"FISHERY_GROUP_CODE", type text},{"FISHERY_GROUP", type text}, {"FISHERY_CODE", type text}, {"FISHERY", type text}, {"GEAR_CODE", type text}, {"GEAR", type text}, {"SPECIES_CATEGORY_CODE", type text}, {"SPECIES_CATEGORY", type text}, {"SPECIES_CODE", type text}, {"SPECIES", type text}, {"SPECIES_SCIENTIFIC", type text}, {"FATE_TYPE_CODE", type text}, {"FATE_TYPE", type text}, {"FATE_CODE", type text}, {"FATE", type text}, {"CATCH", type number}}),
    UniformizaçãoTitulos = Table.RenameColumns(AlterarTipos, {{"YEAR", "Year"}, {"FISHING_GROUND_CODE", "Fishing Ground Code"}, {"FISHING_GROUND", "Fishing Ground"}, {"FLEET_CODE", "Fleet Code"}, {"FLEET", "Fleet"}, {"FISHERY_TYPE_CODE", "Fishery Type Code"}, {"FISHERY_TYPE", "Fishery Type"}, {"FISHERY_GROUP_CODE", "Fishery Group Code"},{"FISHERY_GROUP", "Fishery Group"}, {"FISHERY_CODE", "Fishery Code"}, {"FISHERY", "Fishery"}, {"GEAR_CODE", "Gear FAO Code"}, {"GEAR", "Gear"}, {"SPECIES_CATEGORY_CODE", "Species Category Code"}, {"SPECIES_CATEGORY", "Species Category"}, {"SPECIES_CODE", "Species Code"}, {"SPECIES", "Species Name"}, {"SPECIES_SCIENTIFIC", "Species Scientific Name"}, {"FATE_TYPE_CODE", "Fate Type Code"}, {"FATE_TYPE", "Fate Type"}, {"FATE_CODE", "Fate Code"}, {"FATE", "Fate"}, {"CATCH", "Catch Weight"}}),
   ```

1. Selecting only records >= year 2000:

    ```powerquery
    SelecionarLinhas = Table.SelectRows(UniformizaçãoTitulos, each [Year]>=StartYear),
    ```

1. Adding index to allow for running totals and previous-rows calculations, ordered operations and increase table relationship performance:

    ```powerquery
    Index = Table.AddIndexColumn(SelecionarLinhas,"Index",1,1, Int64.Type),
    ```

1. To normalize the country(fleet) data,  as for exemple could be presented as EU.Portugal, the following was done:

    ```powerquery
     NormalizarTaiwan = Table.TransformColumns( Index, {{"Fleet", each if Text.Contains(Text.Trim(_),"Taiwan,China") then Text.BeforeDelimiter(_,"," ) else _, type text}}),
    NormalizarEU = Table.ReplaceValue(Table.TransformColumns(NormalizarTaiwan, {{"Fleet", each if Text.Contains(Text.Trim(_), "EU") then Text.AfterDelimiter(_,"(") else _, type text}}), ")","", Replacer.ReplaceText, {"Fleet"}),
    ```

1. Some fleets referenced specific regions that constituted different fleets, as for example French Polynesia, and to properly analyze those, specifically, a SubFleet column was added:

    ```powerquery
     AdicionarSubFleet = Table.AddColumn(NormalizarEU, "SubFleet", each if Text.Contains(Text.Trim([Fleet]), ",") then Text.AfterDelimiter([Fleet], ",") else "", type text),
    ```

1. Handling nulls in SubFleet:

    ```powerquery
     SubstituirNullsSubFleet = Table.ReplaceValue(AdicionarSubFleet, "", "Not applicable", Replacer.ReplaceValue, {"SubFleet"}),
    ```

1. Normalize fleet column after extracting subregions

    ```powerquery
    NormalizarFleet = Table.TransformColumns(SubstituirNullsSubFleet,{{"Fleet",each  if Text.Contains(_,",") then Text.BeforeDelimiter(_,",") else _, type text}}),
    ```

1. Normalize NEI values in the fleet column (NEI respects to data of extinct or not known enttities):

    ```powerquery
    NormalizarNEI =  Table.ReplaceValue(NormalizarFleet, "(","", Replacer.ReplaceText, {"Fleet"}),
    ```

1. Adding EU indicator: If country is part of the EU, then Yes, else No (This is relevant because the EU manages fisheries has a block through the EU's Common Fisheries Policy):

    ```powerquery
    AdicionarIdentificadorEU = Table.AddColumn(NormalizarNEI, "EU Fleet", each if Text.Contains([Fleet Code], "EU") then "Yes" else "No", type text),
    ```

1. Round catch to normalize number of digits:

    ```powerquery
    ArrendondarCatch = Table.TransformColumns(AdicionarIdentificadorEU,{{"Catch Weight", each Number.Round(_,2), type number}}),
    ```

1. Reorder columns:

    ```powerquery
    ReordenarColunas = Table.ReorderColumns(ArrendondarCatch, {"Index","Year","Fishing Ground Code","Fishing Ground","Fleet Code","Fleet","SubFleet","EU Fleet","Fishery Type Code","Fishery Type","Fishery Group Code","Fishery Group","Fishery Code","Fishery","Gear FAO Code", "Gear","Species Category Code","Species Category","Species Code", "Species Name", "Species Scientific Name",  "Fate Type Code", "Fate Type", "Fate Code", "Fate", "Catch Weight"})
    ```

  At this stage, the Catch_Estimates fact table has 26 columns. This number will be lower after the creation of the Dim Tables.



### Fleet_Statistics (staging, not imported to the model):


1. Basic transformations: Promote headers, Remove Spaces, alter data types, column title uniformization:


    ```powerquery
    Fonte = Csv.Document(File.Contents(#"Fishing_Craft_Statistics_Path")),
    PromoverNomes = Table.PromoteHeaders(Fonte, [PromoteAllScalars = true]),
    RemoverEspaços = Table.TransformColumns(PromoverNomes, {}, Text.Trim),
    RemoverColunas = Table.RemoveColumns(RemoverEspaços, {"FlSort","Flotte", "TypePêcherie", "Engin", "GrGroupe", "GrMult"}),
    AlterarTitulos = Table.RenameColumns(RemoverColunas, {{"FlCde", "Fleet Code"}, {"Year/An", "Year"}, {"TFCde", "Fishery Type Code"}, {"TypeFishery", "Fishery Type"}, {"GrSort", "Gear Code"},{"GrCde", "Gear FAO Code"}, {"Gear", "Gear"}, {"GrGroup", "Gear Group"}, {"ClassLow", "Class Lower Length"}, {"ClassHigh", "Class Upper Length"}, {"ClassTypeCode", "Class Type Code"}, {"CLass type", "Class Type"}, {"noBoats_nBateaux", "Number of Vessels"}, {"Fgrounds", "Fishing Grounds"}}),
    AlterarTipos = Table.TransformColumnTypes(AlterarTitulos, { {"Fleet Code", type text}, {"Fleet", type text}, {"Year", type number}, {"Fishery Type Code", type text}, {"Fishery Type", type text}, {"Gear Code", type number}, {"Gear FAO Code", type text}, {"Gear", type text}, {"Gear Group", type text}, {"Class Lower Length", type number}, {"Class Upper Length", type number}, {"Class Type Code", type text},{"Class Type", type text}, {"Number of Vessels", type number}, {"Fishing Grounds", type text}}),
    ```

2. Selecting only records >= year 2000:

    ```powerquery
    SelecionarLinhas = Table.SelectRows(AlterarTipos, each [Year]>=StartYear),
    ```
    
3. Adding index:

    ```powerquery
    Index = Table.AddIndexColumn(SelecionarLinhas, "Index",1,1,Int64.Type),
    ```

4. Normalize capital letters in the fleet column:

    ```powerquery
     ProperFleet = Table.TransformColumns(Index, {{"Fleet", each Text.Proper(_)}}),
    ```

5. As the catch estimates fact table, the fleet data needed to be normalized to extract only the actual fleet:

    ```powerquery
    NormalizarTaiwan = Table.TransformColumns( ProperFleet, {{"Fleet", each if Text.Contains(Text.Trim(_),"Taiwan,China") then Text.BeforeDelimiter(_,"," ) else _, type text}}),
    NormalizarFleet = Table.ReplaceValue(NormalizarTaiwan, each [Fleet], each if Text.Contains(Text.Trim([Fleet]), ".") then   Text.Replace([Fleet], ".", " ") else [Fleet], Replacer.ReplaceValue, {"Fleet"}),
    NormalizarEU = Table.TransformColumns(NormalizarFleet, {{"Fleet", each if Text.Contains(Text.Trim(_), "Eu ") then Text.AfterDelimiter(_," ") else _, type text}}),
    // Quando a operação incide só sobre uma coluna sem referir valores noutra, usar TransformColumn
    ```

  Here the use of transform columns is detailed, as the use of replace value gave an error and it was due because TransformColumns is used to iterate the values of one column.

6. As well, the sub regions present in the fleet column needed to be extracted to normalize all fleet values:

    ```powerquery
    AdicionarSubFleet = Table.AddColumn(NormalizarEU, "SubFleet", each if Text.Contains(Text.Trim([Fleet]), " ") and Text.Contains(Text.Trim([Fleet Code]), "EU") or Text.Contains(Text.Trim([Fleet]), "Uk")   then   Text.AfterDelimiter([Fleet], " ") else "", type text),
    SubstituirNullsSubFleet = Table.ReplaceValue(AdicionarSubFleet, "", "Not applicable", Replacer.ReplaceValue, {"SubFleet"}),
    ```

 7. The EU fleet required a second step after the first one to normalized correctly:

    ```powerquery
    NormalizarEU2 = Table.ReplaceValue(SubstituirNullsSubFleet, each [Fleet], each if Text.Contains(Text.Trim([Fleet Code]), "EU") then Text.BeforeDelimiter(_," ") else [Fleet], Replacer.ReplaceValue, {"Fleet"}),
    ```

8. It also required a new correction to the data type:

   ```powerquery
   CorrigirTipoFleet = Table.TransformColumnTypes(NormalizarEU2, {{"Fleet", type text}}),
   ```

9. Add EU Fleet identifier:

   ```powerquery
   AdicionarEUFleet = Table.AddColumn(CorrigirTipoFleet, "EU Fleet", each if Text.Contains(Text.Trim([Fleet Code]), "EU") then "Yes" else "No", type text),
   ```

10. Finally, reorder columns on the Fleet_Statistics fact table. It contains 18 columns for now, as the dim description columns will be retired after the creation of the dim tables:

   ```powerquery
   ReordenarColunas = Table.ReorderColumns(AdicionarEUFleet, {"Index","Year","Fleet Code", "Fleet","SubFleet", "EU Fleet","Fishery Type Code", "Fishery Type", "Gear Code", "Gear FAO Code", "Gear", "Gear Group", "Class Lower Length",  "Class Upper Length", "Class Type Code", "Class Type", "Number of Vessels", "Fishing Grounds"})
   ```



### Fishing_Effort (staging, not imported to the model)

1. As always, basic transformations were performed first:

   ```powerquery
   Fonte = Csv.Document(File.Contents(#"Conexao Pasta" & "\IOTC-DATASETS-2026-02-23-CE-1952-2024.csv")),
    PromoverTitulos = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),
   ```

2. Contrary to the previous tables, some columns were selected immediately to lower the size of the import and optimize subsequent transformations, as the columns left out were clearly not needed. To always select the same columns, applied a dynamic selection by creating a list, select the values of that list that were present in the columns names of the table, also converted as list, and selecting columns by referencing the original table with the PromoteHeaders and the list compared in the previous step:

   ```powerquery
   ColunasPretendidas = {"YEAR", "QUARTER", "FISHING_GROUND_CODE", "FLEET_CODE", "FLEET", "FISHERY_TYPE_CODE", "FISHERY_TYPE", "FISHERY_GROUP_CODE", "FISHERY_GROUP", "FISHERY_CODE", "FISHERY", "GEAR_CODE", "GEAR", "EFFORT_SCHOOL_TYPE_CODE", "CATCH_SCHOOL_TYPE_CODE", "EFFORT", "EFFORT_UNIT_CODE", "SPECIES_CATEGORY_CODE", "SPECIES_CATEGORY", "SPECIES_CODE", "SPECIES", "CATCH_UNIT_CODE", "FATE_TYPE", "FATE_CODE", "FATE", "CATCH"  },
    ColunasPresentes = List.Select(ColunasPretendidas, each List.Contains(Table.ColumnNames(PromoverTitulos),_)),
    SelecionarColunas = Table.SelectColumns(PromoverTitulos, ColunasPresentes),
   ```

3. Replace string values in catch column:

   ```powerquery
    AlterarNAColCatch = Table.ReplaceValue(SelecionarColunas, "NA", 0.0,Replacer.ReplaceValue, {"CATCH"}),
   ```

4. Basic transformation, alter column data types on the dynamically selected columns:

   ```powerquery
   AlterarTipos = Table.TransformColumnTypes(AlterarNAColCatch,{{"YEAR", type number}, {"QUARTER", type text}, {"FISHING_GROUND_CODE",type number}, {"FLEET_CODE", type text},{"FLEET", type text}, {"FISHERY_TYPE_CODE", type text}, {"FISHERY_TYPE", type text}, {"FISHERY_GROUP_CODE", type text}, {"FISHERY_GROUP", type text}, {"FISHERY_CODE", type text}, {"FISHERY", type text}, {"GEAR_CODE", type text}, {"GEAR", type text}, {"EFFORT_SCHOOL_TYPE_CODE",type text} , {"CATCH_SCHOOL_TYPE_CODE",type text},{"EFFORT",type number}, {"EFFORT_UNIT_CODE", type text}, {"SPECIES_CATEGORY_CODE", type text}, {"SPECIES_CATEGORY", type text}, {"SPECIES_CODE", type text}, {"SPECIES", type text}, {"CATCH_UNIT_CODE", type text}, {"FATE_TYPE", type text}, {"FATE_CODE", type text}, {"FATE", type text}, {"CATCH",type number}}),
   ```

5. Filter rows to only select years >= 2000:

   ```powerquery
   SelecionarValores = Table.SelectRows(AlterarTipos, each [YEAR] >= StartYear),
   ```

6. As the previous tables, added an index:

   ```powerquery
   Index = Table.AddIndexColumn(SelecionarValores,"Index", 1,1, Int64.Type),
   ```

7. After the index, the fleet column required normalization. Firstly by taking out parenthesis:
   ```powerquery
   NormalizarFleet = Table.ReplaceValue(Table.ReplaceValue(Index, "(", "", Replacer.ReplaceText, {"FLEET"}), ")", "", Replacer.ReplaceText, {"FLEET"}),
   ```

8. Secondly, by taking out the EU reference:

   ```powerquery
   NormalizarEU = Table.TransformColumns(NormalizarFleet, {{"FLEET", each if Text.Contains(Text.Trim(_), "EU") then Text.AfterDelimiter(_," ") else _, type text}}),
   ```

9. An EU identifier column is added to allow EU fleet analysis with normalized fleet column:

    ```powerquery
     AdicionarIdentificadorEU = Table.AddColumn(NormalizarEU, "EU Fleet", each if Text.Contains([FLEET_CODE], "EU") then "Yes" else "No", type text),
    ```

10. The original table had quarters and years as time-related columns. A QuarterYear key was added to allow for analysis at that granularity level and to connect to the QuarterYear key in the Dim_Date table.

    ```powerquery
    AdicionarChaveQuarterYear = Table.AddColumn(AdicionarIdentificadorEU, "YEARQUARTER", each  Text.From([YEAR]) & [QUARTER], type text),
    ```

11. Lastly, the columns were reordered.

    ```powerquery
    ReordenarColunas = Table.ReorderColumns(AdicionarChaveQuarterYear, {"Index","YEAR", "QUARTER","YEARQUARTER", "FISHING_GROUND_CODE", "FLEET_CODE", "FLEET", "EU Fleet","FISHERY_TYPE_CODE", "FISHERY_TYPE", "FISHERY_GROUP_CODE", "FISHERY_GROUP", "FISHERY_CODE", "FISHERY", "GEAR_CODE", "GEAR", "EFFORT_SCHOOL_TYPE_CODE", "CATCH_SCHOOL_TYPE_CODE", "EFFORT", "EFFORT_UNIT_CODE", "SPECIES_CATEGORY_CODE", "SPECIES_CATEGORY", "SPECIES_CODE", "SPECIES", "CATCH_UNIT_CODE", "FATE_TYPE", "FATE_CODE", "FATE", "CATCH"})
    ```

These tables were left int he Other Queries folder - they are not directly imported. They are the basis for creating the dim tables. After those are created, they are buffered and the, as said earlier, the descriptive columns related to dims are removed before upload to the model.

The tuna_import_prices, the crude_prices and the IOTC_major_geo_areas are imported and uploaded directly and are contained in the Queries to Upload folder inside Power Query.



### Tuna Import Prices (Imported to the model)


1.  The basic transformations required in this table are different. As the information was separated in various sheets with the same columns, those should be combined and exclude hidden sheets:

   ```powerquery
   Fonte = Excel.Workbook(File.Contents(#"Fish_Prices_Excel_Path")),
    FiltrarFolhasOcultas = Table.SelectRows(Fonte, each ([Kind] = "Sheet") and ([Hidden] = false)),
    PromoverCabeçalhos = Table.TransformColumns(FiltrarFolhasOcultas, {{"Data", Table.PromoteHeaders}}),
    CombinarFolhas = Table.Combine(PromoverCabeçalhos[Data]),
   ```

2. The data in the column PROVIDER wasn´t needed, so the column was removed:

   ```powerquery
    RemoverProvider = Table.RemoveColumns(CombinarFolhas, "PROVIDER"),
   ```

3. Another basic transformations were also required: change data types of columns, rename columns and handle nulls in quantitative columns:

   ```powerquery
   AlterarTipos = Table.TransformColumnTypes(RemoverProvider, {{"YEAR", type number}, {"MONTH", type number}, {"SOURCE", type text}, {"UNIT", type text},{"SKJ", type number}, {"YFT", type number}, {"BET", type number}, {"ALB", type number}}),
    RenomearColunas = Table.RenameColumns(AlterarTipos, {{"YEAR", "Year"},{"MONTH", "Month"}, {"SOURCE", "Source"}, {"UNIT", "Unit"}}),
    AlterarNulls = Table.ReplaceValue(RenomearColunas, null, 0.0, Replacer.ReplaceValue, {"SKJ","YFT", "BET", "ALB"}),
   ```

4. The price value columns required normalization on decimal values. As Table.TransformColumns reset the column data type, they should be explicitly declared:

   ```powerquery
   ArredondarValores = Table.TransformColumns(AlterarNulls, {{"SKJ", each Number.Round(_,2), type number},{"YFT", each Number.Round(_,2), type number}, {"BET",  each Number.Round(_,2), type number}, {"ALB",  each Number.Round(_,2), type number}}),
   ```

5. Next, a less straightforward transformation was required. To be able to do meaningful analysis, the dimensions contained in the source column needed to be separated. They contained fish condition, gear used, product market and geographical import/export market. The source column was separated into Fisheries, Market Description, Fish Condition and gear columns. These transformations were encapsulated in a nested let...in statement:

    ```powerquery
   AdicionarColunasCategoricas = let
        AdicionarTipoPescaria = Table.AddColumn(ArredondarValores, "Fisheries" , each Text.Split(Text.Trim([Source]), "|"){0}, type text),
        AdicionarDescricaoMercado = Table.AddColumn(AdicionarTipoPescaria, "Market Description", each Text.Split(Text.Trim([Source]), "|"){1}, type text),
        AdicionarEstadoPescado = Table.AddColumn(AdicionarDescricaoMercado, "Fish Condition", each Text.Split(Text.Trim([Fisheries]), " "){0}, type text ),
        AdicionarAparelhoCaptura =  Table.AddColumn(AdicionarEstadoPescado, "Gear", each Text.AfterDelimiter(Text.Trim([Fisheries])," ", 0), type text ),
   ```

6. The next transformations done inside the nested let...in statement were simpler, normalizing the data inside the gear column, remove now redundant columns and reorder the resulting columns:

   ```powerquery
   AdicionarMaiusculas = Table.TransformColumns(AdicionarAparelhoCaptura, {{"Gear", each Text.Proper(_), type text}}),
        RemoverColunas = Table.RemoveColumns(AdicionarMaiusculas, {"Source", "Fisheries"}),
        ReordenarColunas = Table.ReorderColumns(RemoverColunas, {"Year", "Month", "Market Description" ,"Fish Condition", "Gear", "Unit", "SKJ", "YFT", "BET", "ALB"})
    in
    ReordenarColunas,
   ```


7. The next transformation required several tries. Tried Table.ReplaceValue in several ways and never worked properly, as i needed to iterate each row to normalized the values of weight ie where tons then / 1000 to bring down to kilos. The only way it worked was using List.Accumulate, which allows to alter values based on conditions row by row and only alter when value meet conditions, otherwise maintaining the previous state:

    ```powerquery
    UniformizarPreços_TonsParaKgs = List.Accumulate({"SKJ", "YFT", "BET", "ALB"}, AdicionarColunasCategoricas,  (state,current) => Table.ReplaceValue ( state, each Record.Field(_,current), each if [Unit] = "USD per tonne" and Record.Field(_,current) <> 0.0 then Record.Field(_,current) / 1000
        else Record.Field(_,current), Replacer.ReplaceValue, {current})),
    ```

8. The value description column also needed uniformization row by row, from  tonne to kgs:

    ```powerquery
    UniformizarUnit = Table.ReplaceValue(UniformizarPreços_TonsParaKgs, "USD per tonne", "USD per kg", Replacer.ReplaceValue, {"Unit"}),
    ```

9. The next normalization task concerned currency normalization, as the Japanese prices were in Yen. For a meaningful comparison, these values needed to be converted to USD (the converting values ca be found here: https://www.federalreserve.gov/RELEASES/H10/hist/dat00_ja.htm):

    ```powerquery
    ConvertYen =  List.Accumulate({"SKJ", "YFT", "BET", "ALB"}, UniformizarUnit, (state,current) => Table.ReplaceValue ( state, each Record.Field(_,current), each if [Unit] = "YEN per kg" and Record.Field(_,current) <> 0.0 then Record.Field(_,current) * 0.0063
        else Record.Field(_,current), Replacer.ReplaceValue, {current})),
    UniformizarUnitYENtoUSD = Table.ReplaceValue(ConvertYen, "YEN per kg", "USD per kg", Replacer.ReplaceValue, {"Unit"}),
    ```

10. As with the previous tables, the rows which year >= 2000 were selected:

    ```powerquery
    SelecionarLinhas = Table.SelectRows(UniformizarUnitYENtoUSD, each [Year]>=StartYear),
    ```

11. Next, we unpivot the columns added in the nested let...in statement. UnpivotOtherColumns is preferable to NpivotColumns because, if new columns are added to the source table, the transformation picks them up automatically:

    ```powerquery
    UnpivotColunas = Table.UnpivotOtherColumns(SelecionarLinhas,{"Year", "Month", "Market Description","Fish Condition", "Gear", "Unit"}, "Species Code", "Price per kg"),
    ```


12. The last transformation on the Tubna_Import_Prices table are the explicit definition of the column type for the Price per kg column, adding an YearMonth key and finally reorder de columns:

    ```powerquery
    DefinirTipoValCol = Table.TransformColumnTypes(UnpivotColunas, {{"Price per kg", type number}}),
    AdicionarYearMonth = Table.AddColumn(DefinirTipoValCol, "YearMonth", each [Year]*100 + [Month], type number),
    ReordenarColunas = Table.ReorderColumns(AdicionarYearMonth, {"Year", "Month","YearMonth" ,"Market Description" ,"Fish Condition", "Gear", "Unit", "Species Code", "Price per kg"})
    ```



### Crude_Prices (Imported to the model)


1. Similarly to the previous table, the crude_prices source csv contained several sheets, albeit we only need one. The first transformations are to open that sheet:

   ```powerquery
   Fonte = Excel.Workbook(File.Contents(Crude_Oil_Prices_Excel_Path)),
    FiltrarFolhasOcultas = Table.SelectRows(Fonte, each ([Kind] = "Sheet") and ([Hidden] = false)),
    AbrirFolha = FiltrarFolhasOcultas { [Item = "FUEL | CRUDE OIL SPOT", Kind = "Sheet"]}[Data],
   ```

2. Several basic transformations were performed after opening the sheet, promote headers, remove the redundant provider column, correct the cost column title, explicitly change data types and round the quantitative cost column. As the last transformation uses the TransformColumn, the column data type needs to be explicitly stated (i always forget these):

   ```powerquery
   PromoverCabeçalhos = Table.PromoteHeaders(AbrirFolha, [PromoteAllScallar=true]),
    RemoverColuna = Table.RemoveColumns(PromoverCabeçalhos, {"PROVIDER"}),
    AlterarTitulo = Table.RenameColumns(RemoverColuna, {"COS", "COST"}),
    AlterarTipo = Table.TransformColumnTypes(AlterarTitulo, {{"YEAR", type number}, {"MONTH", type number}, {"SOURCE", type text}, {"UNIT", type text}, {"COST", type number}}),
    ArrendondarCusto = Table.TransformColumns(AlterarTipo, {{"COST", each Number.Round(_, 2), type number}}),
   ```

3. The rows >= year 2000 were filtered:

   ```powerquery
    SelecionarLinhas = Table.SelectRows(ArrendondarCusto, each [YEAR] >= StartYear),
   ```

4. After that, the current time columns available were used to add Quarter, QuarterYear and YearMonth Columns:

   ```powerquery
   AdicionarQuarter = Table.AddColumn(SelecionarLinhas, "QUARTER", each if [MONTH] <= 3 then "Q1" else if [MONTH] <= 6 then "Q2" else if [MONTH] >= 9 then "Q3" else "Q4", type text),
    AdicionarYearQuarter = Table.AddColumn(AdicionarQuarter, "YearQuarter", each Text.From([YEAR]) & [QUARTER], type text),
    AdicionarYearMonth = Table.AddColumn(AdicionarYearQuarter, "YearMonth", each Text.From([YEAR]) & Text.From([MONTH]), type text),
   ```

5. Finally, the columns were reordered:

   ```powerquery
   ReordenarColunas = Table.ReorderColumns(AdicionarYearMonth, {"YEAR", "MONTH", "YearMonth", "YearQuarter", "SOURCE", "UNIT", "COST"})
   ```


After these tables, the Dim tables were built from the categorical data contained in the fact tables described above. All Dim tables were organized in the Dimensions folder in Power Query Editor.



### Dim_Date (Imported to model):


1. Dim_Date M Code is widely available online and in the Power Query M Manual uploaded in Manuals a full date table in m is available, so just a quick explanation. Simply define a initial and end date, define the culture in which the dates and names are to be presented and counted, use Duration.Days to count the different dates in the range. Then utilize List.Dates to use that initial date and counting of dates to produce a list of dates corresponding to that initial date and subsequent dates contained in the counting. From that list, create a table, define the type date for the column date and from there add transformations to define months, quarters, names for days, months and so on. Finally, define the YearMonth and YearQuarter, in this case. 

   ```powerquery
   let
    DataInicial = #date(2000,1,1),
    DataFinal = #date(2030,1,1),
    Cultura = "en-EN",
    ContagemDias = Duration.Days(DataFinal-DataInicial)+1,
    ListaDias = List.Dates(DataInicial , ContagemDias, #duration(1,0,0,0)),
    TabelaDias = Table.FromList(ListaDias, Splitter.SplitByNothing(),{"Date"}),
    DefinirTipo = Table.TransformColumnTypes(TabelaDias, {{"Date", type date}}),
    AdicionarAno = Table.AddColumn(DefinirTipo, "Year", each Date.Year([Date]), Int64.Type),
    AdicionarTrimestre = Table.AddColumn(AdicionarAno, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    AdicionarDia = Table.AddColumn(AdicionarTrimestre, "Day Number", each Date.Day([Date]), Int64.Type),
    AdicionarMes = Table.AddColumn(AdicionarDia, "Month Number", each Date.Month([Date]), Int64.Type),
    AdicionarNomeMes = Table.AddColumn(AdicionarMes, "Month Name", each Date.MonthName([Date]), type text ),
    AdicionarNomeDia = Table.AddColumn(AdicionarNomeMes, "Day Name", each Date.DayOfWeekName([Date]), type text),
    AdicionarYearMonth = Table.AddColumn(AdicionarNomeDia, "YearMonth", each [Year]*100+[Month Number], Int64.Type),
    AdicionarYearQuarter = Table.AddColumn(AdicionarYearMonth, "YearQuarter", each Text.From([Year]) & [Quarter], type text),
    ReordenarColunas = Table.ReorderColumns(AdicionarYearQuarter, {"Date", "Day Number", "Day Name", "Month Number", "Month Name", "Quarter", "Year", "YearMonth", "YearQuarter"})
    in
    ReordenarColunas
   ```


### Dim_Species (Imported to the model)

1. Initially, the code to define the dim_species table was solely based on the Catch_Estimates imported table, using a dynamic selection of column names, select those columns with Table.SelectColumns, Table.Distinct to select only the distinct dimensional values, adding an index column  and finally reordering the columns.

2. However, when reached the visualization phase, there was a blank value appearing that should no exist. The measures used were related to the Fishing_Effort fact table, indicating that were values present in this table that were not present in the Catch_Estimates table. Therefore, after some head-banging, the solution found was to do the process described above and combine the resulting tables. This was possible because the same columns were present in the fact tables. Tried nested joins but the performance was worse than combining tables, which was achived more quicly.

3. First, selected the columns from Catch_Estimate, dinamically:

   ```powerquery
   ColunasPretendidas = {"Species Category Code", "Species Category", "Species Code", "Species Name"},
    VerificarColunas = List.Select(ColunasPretendidas, each List.Contains(Table.ColumnNames(Catch_Estimates),_)),
    SelecionarColunas = Table.SelectColumns(Catch_Estimates, VerificarColunas),
   ```

5. Secondly, select the same columns from the Fishing_Effort Table:

   ```powerquery
   ColunasPretendidasEffort= {"SPECIES_CATEGORY_CODE", "SPECIES_CATEGORY", "SPECIES_CODE", "SPECIES"},
    VerificarColunasEffort= List.Select(ColunasPretendidasEffort, each List.Contains(Table.ColumnNames(Fishing_Effort),_)),
    SelecionarColunasEffort = Table.SelectColumns(Fishing_Effort, VerificarColunasEffort),
   ```

6. As the columns titles were writen differently, required normalization:
  
   ```powerquery
   UniformizarColunas = Table.RenameColumns(SelecionarColunasEffort, {{"SPECIES_CATEGORY_CODE","Species Category Code"}, {"SPECIES_CATEGORY", "Species Category"}, {"SPECIES_CODE", "Species Code"}, {"SPECIES", "Species Name"}}),
   ```

7. Then the two resulting tables were combined:

   ```powerquery
   CombinarTabelas = Table.Combine({SelecionarColunas, UniformizarColunas}),
   ```

8. To make it a dimension table, distinct values were filtered:

   ```powerquery
    SelecionarDistintos = Table.Distinct(CombinarTabelas, {"Species Code"}),
   ```

9. As the previous tables, an index column is added:

   ```powerquery
   Index= Table.AddIndexColumn(SelecionarDistintos, "Index", 1, 1, Int64.Type),
   ```

11. Lastly, the columns were reordered:

    ```powerquery
    ReordenarColunas = Table.ReorderColumns(Index, {"Index", "Species Code", "Species Name", "Species Category Code", "Species Category"})
    ```



### Dim_Fisheries&Gear (Imported to the model)


1. The same problem that arose during the report phase for the Dim_Species arose here as well. Applied the same initial procedure and, as the source columns were the same, applied the same solution when the blanks appeared in visualizations for metrics measuring Gear and Fisheries.

2. The only difference was, after analyzing the resulting table, the values with NA string that came from the fishing effort table needed to be normalized to "Not specified".

3. The complete code is:

   ```powerquery
   let
    
    ColunasPretendidas = {"Fishery Type Code", "Fishery Type", "Fishery Group Code", "Fishery Group", "Fishery Code", "Fishery", "Gear FAO Code", "Gear"},
    VerificarColunas = List.Select(ColunasPretendidas, each List.Contains(Table.ColumnNames(Catch_Estimates), _)),
    SelecionarColunas = Table.SelectColumns(Catch_Estimates, VerificarColunas),
    ColunasPretendidasEffort = {"GEAR_CODE", "GEAR", "FISHERY", "FISHERY_CODE", "FISHERY_GROUP", "FISHERY_GROUP_CODE", "FISHERY_TYPE", "FISHERY_TYPE_CODE"},
    VerificarColunasEffort = List.Select(ColunasPretendidasEffort, each List.Contains(Table.ColumnNames(Fishing_Effort), _)),
    SelecionarColunasEffort = Table.SelectColumns(Fishing_Effort,VerificarColunasEffort),
    UniformizarNomesColunas = Table.RenameColumns(SelecionarColunasEffort, {{"GEAR_CODE", "Gear FAO Code"},{"GEAR", "Gear"},{"FISHERY", "Fishery"},{"FISHERY_CODE", "Fishery Code"},{"FISHERY_GROUP", "Fishery Group"},{"FISHERY_GROUP_CODE", "Fishery Group Code"},{"FISHERY_TYPE", "Fishery Type"},{"FISHERY_TYPE_CODE", "Fishery Type Code"}}),
    CombinarCatcheEffort  = Table.Combine({ SelecionarColunas, UniformizarNomesColunas}),
    SelecionarDistintos = Table.Distinct(CombinarCatcheEffort, {"Gear FAO Code"}),
    Index = Table.AddIndexColumn(SelecionarDistintos, "Index", 1, 1, Int64.Type),
    ReordenarColunas = Table.ReorderColumns(Index, {"Index", "Gear", "Gear FAO Code", "Fishery", "Fishery Code", "Fishery Group", "Fishery Group Code", "Fishery Type", "Fishery Type Code"}),
    TratarNulls = Table.ReplaceValue(ReordenarColunas, "NA", "Not specified", Replacer.ReplaceValue, {"Gear", "Gear FAO Code"})

    in
    TratarNulls
   ```



### Dim_Fleet (Imported to the model)


1. As with the previous dim tables, the initial transformations resulted in mismatches in the report phase. When analysing the distinct values in the fleet column in the catch and fleet statistics fact tables, there was a difference of 7 values (54 vs 47). Therefore, the same process was applied here, with the difference that a nested join was used to bring to the nested table column the values not present in the left table. In this case, the Fleet_Statistics table:

   ```powerquery
   Catch_Estimates = Table.Buffer(Catch_Estimates),
    SelecionarColunas = Table.SelectColumns(Catch_Estimates, {"Fleet Code", "Fleet", "SubFleet", "EU Fleet"}),
    SelecionarDistintos = Table.Distinct(SelecionarColunas, {"Fleet Code"}),
    UniraFleetStatistics = Table.NestedJoin(SelecionarDistintos, {"Fleet Code"}, Fleet_Statistics, {"Fleet Code"}, "Fleet_Joined", JoinKind.LeftOuter),
    ExpandirColuna = Table.ExpandTableColumn(UniraFleetStatistics, "Fleet_Joined", {"Class Type Code", "Class Type"}),
   ```

2. After expanding the nested columns, null required handling in non-matched values:

   ```powerquery
   TratarNulls = Table.ReplaceValue(ExpandirColuna, null, "Not specified", Replacer.ReplaceValue, {"Class Type", "Class Type Code"}),
   ```

3. Next, only distinct values were kept:

   ```powerquery
   Deduplicar = Table.Distinct(TratarNulls, {"Fleet Code"}),
   ```

4. Finally, an index was added and the columns were reordered:

   ```powerquery
   Index = Table.AddIndexColumn(Deduplicar, "Index", 1, 1, Int64.Type),
    ReordenarColunas = Table.ReorderColumns(Index, {"Index", "Fleet", "SubFleet", "Fleet Code", "Class Type", "Class Type Code", "EU Fleet"})
   ```



### Dim_Fate (Imported to the model)


1. The creation of this table stemmed from my experience as an observer, as caught fish can have multiple destinations. Could be retained for landing and sale, could be discarded due to absence of quota, because a moratorium is applied to that species, it has depredation and is not in sale condition or could be used for crew consumption, for example. Although currently only holds one distinct value, future analysis could benefit greatly for an expanded collection on fate data.

2. This is the simplest dimension table. Just applied a dynamic selection of columns, filtered distinct values, added an index and reordered columns.

   ```powerquery
   let
    ColunasPretendidas = {"Fate Code", "Fate"},
    VerificarColunas = List.Select(ColunasPretendidas, each List.Contains(Table.ColumnNames(Catch_Estimates),_)),
    SelecionarColunas = Table.SelectColumns(Catch_Estimates, VerificarColunas),
    AplicarDistinct = Table.Distinct(SelecionarColunas, {"Fate Code"}),
    Index = Table.AddIndexColumn(AplicarDistinct, "Index", 1, 1, Int64.Type),
    ReordenarColunas =  Table.ReorderColumns(Index, {"Index", "Fate", "Fate Code"})
    in
    ReordenarColunas

   ```


As mentioned before, after the creation of the Dimension Tables, the fact tables that were uploaded to the model were defined. This step consisted on excluding now redundant columns before upload.




### Catch_Estimates_Model (uploaded to the model)


1. First, the staging table was buffered:

    ```powerquery
   Fonte = Table.Buffer(Catch_Estimates),
   ```

2. Secondly, the now redundant columns were removed:

   ```powerquery
   RemoverColunas = Table.RemoveColumns(Fonte, {"Fishery Type", "Fishery Group", "Fishery", "Gear", "Fleet", "SubFleet", "Fate","Fate Type",  "Species Name", "Species Scientific Name", "Species Category"})
   ```



### Fleet_Statistics_Model  & Fishing_Effort_Model (uploaded to the model):


1. The same transformations were applied to the other fact tables used in the model. The table Fleet_Statistics_Model was defined as:

   ```powerquery
   let
    Fonte = Table.Buffer(Fleet_Statistics),
    RemoverColunas = Table.RemoveColumns(Fonte, {"Fishery Type", "Fleet", "SubFleet", "Class Type", "Gear", "Gear Code"})
    in
    RemoverColunas
   ```

2. The Fishing Effort Model was defined in the same way:

   ```powerquery
   let
    Fonte = Table.Buffer(Fishing_Effort),
    RemoverColunas = Table.RemoveColumns(Fonte, {"FLEET", "EU Fleet", "FISHERY_TYPE", "FISHERY_GROUP", "FISHERY", "GEAR", "SPECIES_CATEGORY", "SPECIES", "FATE", "FATE_TYPE"})
    in
    RemoverColunas
   ```



This ends the Data Preparation phase.




## Data Model

The resulting data model is a multi-fact star-schema with multiple one to many (1 : *) relationships from dimensional tables to the fact tables. Bridge tables were calculated in DAX, which will be described below. 

### Data Model Diagram:
![IOTC Model](<assets/- IOTC Model.png>)









   

   
   




   




   

  

    


     




