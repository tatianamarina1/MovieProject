# Movie Dashboard
## Table of Content
- [Problem Statement](#problem-statement)
- [Data Source](#data-source)
- [Tools](#tools)
- [Data Cleaning](#data-cleaning)
- [Dashboard](#dashboard)
- [M code](#m-code)
- [Recommendations](#recommendations)

### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres.
Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:
- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions

The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.

### Data Source
Movie Data : The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors and genres. You can download the original datasource here: [Movie Dataset Excel file](https://github.com/user-attachments/files/34049964/Movies_Data_Homework.xlsx)

### Tools
1. Power Query - I used Power Query for Data Cleaning and Data Transformation
2. Excel - I used Excel for Data Analysis
3. Pivot Tables - for Creating the dashboard and Visualizations

### Data Cleaning
- Data loading and inspection.
- Combine data from different sources and tabs
- Data cleaning and formatting. The excel file after the data cleaning & preparation process can be downloaded here - [Movies Dashboard](https://github.com/user-attachments/files/34050753/Dashboard_Movies_Data_Homework.xlsx)

### Dashboard
<img width="776" height="625" alt="Dashboard_excel" src="https://github.com/user-attachments/assets/e8272a18-1919-4e60-856c-43d89f6f0347" />

### M code
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\suri-\Documents\Data_analyst_course\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Removed Columns" = Table.RemoveColumns(#"Merged Queries",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Removed Columns", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Genres",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries1" = Table.NestedJoin(#"Reordered Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre.1"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Genres1",{{"Genres.Genre.1", "Genre_Second"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns1",{{"Genres.Genre", "Genre"}}),
    #"Merged Queries2" = Table.NestedJoin(#"Renamed Columns1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Expanded Directors",{{"Directors.Director", "Director"}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries3" = Table.NestedJoin(#"Reordered Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actor"}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries4" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Merged Queries4", "Actors", {"Actor"}, {"Actor.1"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Expanded Actors1",{{"Actor.1", "Actor_2"}}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Renamed Columns3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries5" = Table.NestedJoin(#"Reordered Columns4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Merged Queries5", "Actors", {"Actor"}, {"Actor.1"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Expanded Actors2",{{"Actor.1", "Actor_3"}}),
    #"Reordered Columns5" = Table.ReorderColumns(#"Renamed Columns4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Actor_3", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries6" = Table.NestedJoin(#"Reordered Columns5", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Merged Queries6", "Actors", {"Actor"}, {"Actor.1"}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Expanded Actors3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Actor_3", "Cast_Fourth_ID", "Actor.1", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns5" = Table.RenameColumns(#"Reordered Columns6",{{"Actor.1", "Actor_4"}}),
    #"Merged Queries7" = Table.NestedJoin(#"Renamed Columns5", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors4" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actor.1"}),
    #"Renamed Columns6" = Table.RenameColumns(#"Expanded Actors4",{{"Actor.1", "Actor_5"}}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Renamed Columns6",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Actor_3", "Cast_Fourth_ID", "Actor_4", "Cast_Fifth_ID", "Actor_5", "Budget ($)", "Box Office Revenue ($)"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns7", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}})
in
    #"Changed Type1"
```
### Recommendations
Top 5 genres are Action, Comedy, Drama, Sci-Fi and Adventure. I would recommend Netflix to produce a movie with one of these genres as they brought in more in box office revenue based on the data from 2012 to 2016

<img width="643" height="246" alt="image" src="https://github.com/user-attachments/assets/7ce0264d-8bd7-44c1-b334-9faf55b911e9" />

Horror is the most successful genre in term of ROI (return on investment).
<img width="772" height="252" alt="image" src="https://github.com/user-attachments/assets/cc80abdf-950d-4a85-be97-6396114e2837" />




