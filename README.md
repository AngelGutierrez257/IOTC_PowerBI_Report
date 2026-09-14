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
  - Connection type: Csv with parameter
    ``` Fonte = Csv.Document(File.Contents(#"RC-SCI_1950-2024_PATH"))´´´ with parameter RC-SCI_1950-2024_PATH.

### Yearly fleet statistics:
  - Original Title: Fishing_Craft_Statistics_20250702.zip
  - Model Title: Fleet_Statistics
  - Origin: https://iotc.org/data/datasets/Annual Number of Vessels by Fishing Fleet, Gear, Architecture, Mechanisation type, Size class, and Fish Preservation Method (Fishing Craft Statistics)
  - Connector type: Csv with parameter
     ``` Fonte = Csv.Document(File.Contents(#"Conexao Pasta" & "\IOTC-DATASETS-2026-02-23-CE-1952-2024.csv")) ´´´ with parameter Conexao Pasta as Folder Connector.
   
### Monthly Fishing Effort
  - Original Title: IOTC-DATASETS-2026-02-23-CE-1952-2024.csv
  - Model Title: Fishing_Effort
  - Origin: https://iotc.org/data/datasets/latest/CE/All
  - Connector type: Csv with parameter
     ``` Fonte = Csv.Document(File.Contents(#"Conexao Pasta" & "\IOTC-DATASETS-2026-02-23-CE-1952-2024.csv")),´´´ with parameter parameter Conexao Pasta as Folder Connector.

### Tuna Import Prices
  - Original Title: FFA_import_price_tuna_time_series.xlsx
  - Model Title: Tuna_Import_Prices
  - Origin: https://iotc.org/data/datasets/latest/SD/TUNAS
  - Connector type: Csv with parameter
    ```
    Fonte = Excel.Workbook(File.Contents(#"Fish_Prices_Excel_Path")),
    ´´´ with parameter Fish_Prices_Excel_Path.

### Crude Prices
  - Original Title: FFA_crude_oil_price_time_series.xlsx
  - Model Title: Crude_Prices
  - Origin: https://iotc.org/data/datasets/latest/SD/FUEL
  -  Connector type: Csv with parameter
      ```
      Fonte = Excel.Workbook(File.Contents(Crude_Oil_Prices_Excel_Path)),
      ´´´ with parameter Crude_Oil_Prices_Excel_Path.

### IOTC Main Areas (Geographical data)
  - Original Title: IOTC_MAIN_AREAS_10.0.0.csv
  - Model Title: IOTC_Major_Geo_Areas
  - Origin: https://data.iotc.org/reference/latest/domain/admin/#geospatialData
     ```
     Fonte = Csv.Document(File.Contents(Geographical_Data_Connection)),
     ´´´ with parameter Geographical_Data_Connection.
 

The use of parameters is preferred has it handles changes in source locations more gracefully and easily and also allow the use of deployment pipelines.


## 2 Data Preparation - Power Query:

  Mai tasks: Data normalization, null handling, prepare data for modelling, filter and define data types, rows and columns relevant for the model.

  ### Power Query Parameter: 
  Besides the connectors, a parameter was defined to uniformly select rows >= the year 2000. 
  Parameter Name: StartYear


  ## Fact Tables:

  ### Catch_Estimates (not imported to the model)
  
  -  Basic transformations: Promote headers, Remove Spaces, alter data types, column title uniformization:
    
      ``` 
      Fonte = Csv.Document(File.Contents(#"RC-SCI_1950-2024_PATH")),
    PromoverNomes = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),
    RemoverEspaços = Table.TransformColumns(PromoverNomes, {}, Text.Trim),
    AlterarTipos =  Table.TransformColumnTypes(RemoverEspaços, {{"YEAR", type number}, {"FISHING_GROUND_CODE", type text}, {"FISHING_GROUND", type text}, {"FLEET_CODE", type text}, {"FLEET", type text},   {"FISHERY_TYPE_CODE", type text}, {"FISHERY_TYPE", type text}, {"FISHERY_GROUP_CODE", type text},{"FISHERY_GROUP", type text}, {"FISHERY_CODE", type text}, {"FISHERY", type text}, {"GEAR_CODE", type text}, {"GEAR", type text}, {"SPECIES_CATEGORY_CODE", type text}, {"SPECIES_CATEGORY", type text}, {"SPECIES_CODE", type text}, {"SPECIES", type text}, {"SPECIES_SCIENTIFIC", type text}, {"FATE_TYPE_CODE", type text}, {"FATE_TYPE", type text}, {"FATE_CODE", type text}, {"FATE", type text}, {"CATCH", type number}}),
    UniformizaçãoTitulos = Table.RenameColumns(AlterarTipos, {{"YEAR", "Year"}, {"FISHING_GROUND_CODE", "Fishing Ground Code"}, {"FISHING_GROUND", "Fishing Ground"}, {"FLEET_CODE", "Fleet Code"}, {"FLEET", "Fleet"}, {"FISHERY_TYPE_CODE", "Fishery Type Code"}, {"FISHERY_TYPE", "Fishery Type"}, {"FISHERY_GROUP_CODE", "Fishery Group Code"},{"FISHERY_GROUP", "Fishery Group"}, {"FISHERY_CODE", "Fishery Code"}, {"FISHERY", "Fishery"}, {"GEAR_CODE", "Gear FAO Code"}, {"GEAR", "Gear"}, {"SPECIES_CATEGORY_CODE", "Species Category Code"}, {"SPECIES_CATEGORY", "Species Category"}, {"SPECIES_CODE", "Species Code"}, {"SPECIES", "Species Name"}, {"SPECIES_SCIENTIFIC", "Species Scientific Name"}, {"FATE_TYPE_CODE", "Fate Type Code"}, {"FATE_TYPE", "Fate Type"}, {"FATE_CODE", "Fate Code"}, {"FATE", "Fate"}, {"CATCH", "Catch Weight"}}),

      ´´´

  - Selecting only records >= year 2000:
    ```
    SelecionarLinhas = Table.SelectRows(UniformizaçãoTitulos, each [Year]>=StartYear),
    
    ´´´

  - Adding index to allow for running totals and previous-rows calculations, ordered operations and increase table relationship performance:
    ```
    Index = Table.AddIndexColumn(SelecionarLinhas,"Index",1,1, Int64.Type),
    ´´´

  - To normalize the country(fleet) data,  as for exemple could be presented as EU.Portugal, the following was done:
    ```
     NormalizarTaiwan = Table.TransformColumns( Index, {{"Fleet", each if Text.Contains(Text.Trim(_),"Taiwan,China") then Text.BeforeDelimiter(_,"," ) else _, type text}}),
    NormalizarEU = Table.ReplaceValue(Table.TransformColumns(NormalizarTaiwan, {{"Fleet", each if Text.Contains(Text.Trim(_), "EU") then Text.AfterDelimiter(_,"(") else _, type text}}), ")","", Replacer.ReplaceText, {"Fleet"}),
    ´´´

  - Some fleets referenced specific regions that constituted different fleets, as for example French Polynesia, and to properly analyze those, specifically, a SubFleet column was added:
    ```
     AdicionarSubFleet = Table.AddColumn(NormalizarEU, "SubFleet", each if Text.Contains(Text.Trim([Fleet]), ",") then Text.AfterDelimiter([Fleet], ",") else "", type text),
    ´´´
    
  - Handling nulls in SubFleet:
    ```
     SubstituirNullsSubFleet = Table.ReplaceValue(AdicionarSubFleet, "", "Not applicable", Replacer.ReplaceValue, {"SubFleet"}),
    ´´´

  - Normalize fleet column after extracting subregions:
    ```
    NormalizarFleet = Table.TransformColumns(SubstituirNullsSubFleet,{{"Fleet",each  if Text.Contains(_,",") then Text.BeforeDelimiter(_,",") else _, type text}}),
    ´´´

  - Normalize NEI values in the fleet column (NEI respects to data of extinct or not known enttities):
    ```
    NormalizarNEI =  Table.ReplaceValue(NormalizarFleet, "(","", Replacer.ReplaceText, {"Fleet"}),
    ´´´

  - Adding EU indicator: If country is part of the EU, then Yes, else No (This is relevant because the EU manages fisheries has a block through the EU's Common Fisheries Policy):
    ```
    AdicionarIdentificadorEU = Table.AddColumn(NormalizarNEI, "EU Fleet", each if Text.Contains([Fleet Code], "EU") then "Yes" else "No", type text),
    ´´´

  - Round catch to normalize number of digits:
    ```
    ArrendondarCatch = Table.TransformColumns(AdicionarIdentificadorEU,{{"Catch Weight", each Number.Round(_,2), type number}}),
    ´´´

  - Reorder columns:
    ```
    ReordenarColunas = Table.ReorderColumns(ArrendondarCatch, {"Index","Year","Fishing Ground Code","Fishing Ground","Fleet Code","Fleet","SubFleet","EU Fleet","Fishery Type Code","Fishery Type","Fishery Group Code","Fishery Group","Fishery Code","Fishery","Gear FAO Code", "Gear","Species Category Code","Species Category","Species Code", "Species Name", "Species Scientific Name",  "Fate Type Code", "Fate Type", "Fate Code", "Fate", "Catch Weight"})
    ´´´

  At this stage, the Catch_Estimates fact table has 26 columns. This number will be lower after the creation of the Dim Tables.


  ### Fleet_Statistics (not imported to the model):

  -  Basic transformations: Promote headers, Remove Spaces, alter data types, column title uniformization:
    
    ```Fonte = Csv.Document(File.Contents(#"Fishing_Craft_Statistics_Path")),
    PromoverNomes = Table.PromoteHeaders(Fonte, [PromoteAllScalars = true]),
    RemoverEspaços = Table.TransformColumns(PromoverNomes, {}, Text.Trim),
    RemoverColunas = Table.RemoveColumns(RemoverEspaços, {"FlSort","Flotte", "TypePêcherie", "Engin", "GrGroupe", "GrMult"}),
    AlterarTitulos = Table.RenameColumns(RemoverColunas, {{"FlCde", "Fleet Code"}, {"Year/An", "Year"}, {"TFCde", "Fishery Type Code"}, {"TypeFishery", "Fishery Type"}, {"GrSort", "Gear Code"},{"GrCde", "Gear FAO Code"}, {"Gear", "Gear"}, {"GrGroup", "Gear Group"}, {"ClassLow", "Class Lower Length"}, {"ClassHigh", "Class Upper Length"}, {"ClassTypeCode", "Class Type Code"}, {"CLass type", "Class Type"}, {"noBoats_nBateaux", "Number of Vessels"}, {"Fgrounds", "Fishing Grounds"}}),
    AlterarTipos = Table.TransformColumnTypes(AlterarTitulos, { {"Fleet Code", type text}, {"Fleet", type text}, {"Year", type number}, {"Fishery Type Code", type text}, {"Fishery Type", type text}, {"Gear Code", type number}, {"Gear FAO Code", type text}, {"Gear", type text}, {"Gear Group", type text}, {"Class Lower Length", type number}, {"Class Upper Length", type number}, {"Class Type Code", type text},{"Class Type", type text}, {"Number of Vessels", type number}, {"Fishing Grounds", type text}}),
    ´´´ 

  - Selecting only records >= year 2000:
    ```
    SelecionarLinhas = Table.SelectRows(AlterarTipos, each [Year]>=StartYear),
    ´´´
  - Adding index: 
    ```
    Index = Table.AddIndexColumn(SelecionarLinhas, "Index",1,1,Int64.Type),
    ´´´

  - Normalize capital letters in the fleet column:
    ```
     ProperFleet = Table.TransformColumns(Index, {{"Fleet", each Text.Proper(_)}}),
    ´´´

  - As the catch estimates fact table, the fleet data needed to be normalized to extract only the actual fleet:
    ```
    NormalizarTaiwan = Table.TransformColumns( ProperFleet, {{"Fleet", each if Text.Contains(Text.Trim(_),"Taiwan,China") then Text.BeforeDelimiter(_,"," ) else _, type text}}),
    NormalizarFleet = Table.ReplaceValue(NormalizarTaiwan, each [Fleet], each if Text.Contains(Text.Trim([Fleet]), ".") then   Text.Replace([Fleet], ".", " ") else [Fleet], Replacer.ReplaceValue, {"Fleet"}),
    NormalizarEU = Table.TransformColumns(NormalizarFleet, {{"Fleet", each if Text.Contains(Text.Trim(_), "Eu ") then Text.AfterDelimiter(_," ") else _, type text}}),
    // Quando a operação incide só sobre uma coluna sem referir valores noutra, usar TransformColumn
    ´´´
  Here the use of transform columns is detailed, as the use of replace value gave an error and it was due because TransformColumns is used to iterate the values of one column.

  - As well, the sub regions present in the fleet column needed to be extracted to normalize all fleet values:
    ```
 AdicionarSubFleet = Table.AddColumn(NormalizarEU, "SubFleet", each if Text.Contains(Text.Trim([Fleet]), " ") and Text.Contains(Text.Trim([Fleet Code]), "EU") or Text.Contains(Text.Trim([Fleet]), "Uk")   then Text.AfterDelimiter([Fleet], " ") else "", type text),
    SubstituirNullsSubFleet = Table.ReplaceValue(AdicionarSubFleet, "", "Not applicable", Replacer.ReplaceValue, {"SubFleet"}),
    ´´´

 - The EU fleet required a second step after the first one to normalized correctly:
   ```
    NormalizarEU2 = Table.ReplaceValue(SubstituirNullsSubFleet, each [Fleet], each if Text.Contains(Text.Trim([Fleet Code]), "EU") then Text.BeforeDelimiter(_," ") else [Fleet], Replacer.ReplaceValue, {"Fleet"}),
   ´´´

 - It also required a new correction to the data type:
   ```
   CorrigirTipoFleet = Table.TransformColumnTypes(NormalizarEU2, {{"Fleet", type text}}),
   ´´´

 - Add EU Fleet identifier:
   ```
   AdicionarEUFleet = Table.AddColumn(CorrigirTipoFleet, "EU Fleet", each if Text.Contains(Text.Trim([Fleet Code]), "EU") then "Yes" else "No", type text),
   ´´´
 - Finally, reorder columns on the Fleet_Statistics fact table. It contains 18 columns for now, as the dim description columns will be retired after the creation of the dim tables:
   ```
   ReordenarColunas = Table.ReorderColumns(AdicionarEUFleet, {"Index","Year","Fleet Code", "Fleet","SubFleet", "EU Fleet","Fishery Type Code", "Fishery Type", "Gear Code", "Gear FAO Code", "Gear", "Gear Group", "Class Lower Length",  "Class Upper Length", "Class Type Code", "Class Type", "Number of Vessels", "Fishing Grounds"})
   ´´´
   

  

    


     




