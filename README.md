# IOTC Catch and Effort Analysis - Power BI Semantic Model and Report

## Synopsis:

This report analyses economic, effort and catch data from IOTC to assess fleet efficiency and quota sustainability by fleet (flag), commercial species and gear type for the period from 2000 to 2024. The analytical framework and metrics draw directly from my experience as a scientific fisheries observer in the IOTC Regulatory Area. 

The report and semantic model were built as an applied project for the PL-300 certification, using Power Query for data preparation and modelling, and DAX for time intelligence, ranking, what-if scenarios and statistical analysis. See Data Preparation and Dax Measures & visualizations sections for the complete code, explained. 

For self-study, I asked Claude to prepare manuals for Dax and Power Query based on the Feynman technique and organize them based on the old Skillpipe manuals, in order to prepare for the exam. The visualization and administration skills were studied trough the Microsoft Learning Path (https://learn.microsoft.com/en-us/credentials/certifications/data-analyst-associate/?practice-assessment-type=certification) and the book PL-300 Exam Ref, of author Daniil Maslyuk.

## Highlights:

- Multifact star schema with 5 fact tables, 8 dimension tables and 2 DAX-calculated bridge tables.
- Bridge tables solve temporal granularity mismatches across sources.
- Time intelligence analysis: rolling 3 year averages, Year to Date (YTD) and Year over Year (YoY) calculations.
- What-if scenario modelling of revenue impact under quota reduction scenarios.
- Stock health signals derived from CPUE trend analysis.

## Snapshots:

### Model:
![IOTC Model](<assets/- IOTC Model .png>)


### Dashboard:
![IOTC Model](<assets/- IOTC Report Dashboard.png>)

## Repository structure:

- IOTC/ contains the semantic model and report json files.
- Manuals contains the Power Query and Dax manuals used to support development of this project and study for PL-300.
- assets/ contains illustrative images of each page of the report and of the semantic model.
- README.md contains the project presentation.
- Technical report.md report contains full technical documentation, with explained code:
   [Technical report.md](<Technical report.md>)


   




   

  

    


     




