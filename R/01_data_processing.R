library(dplyr)
library(openxlsx2)
library(here)

# Work with output files ####
Bridge_Facility <- arrow::read_parquet(here(
    "data/output/Py_Bridge_Facility.parquet"
))
Dim_Work_Unit <- arrow::read_parquet(here(
    "data/output/Py_Dim_Work_Unit.parquet"
))
Employee_Telework_Table <- arrow::read_parquet(here(
    "data/output/Py_Employee_Telework_Table.parquet"
))
Facility_Table <- arrow::read_parquet(here(
    "data/output/Py_Facility_Table.parquet"
))
HQ_Telework_Table <- arrow::read_parquet(here(
    "data/output/Py_HQ_Telework_Table.parquet"
))

addressTable <- Facility_Table |>
    select(
        Address = AddressEdit,
        AddressLink,
        City,
        Building_Number,
        Total_Rentable_Area,
        UOM,
        Latitude,
        Longitude
    ) |>
    distinct() |>
    mutate(
        BuildingFlag = case_when(
            startsWith(as.character(Building_Number), "N") ~ FALSE,
            startsWith(as.character(Building_Number), "B") ~ TRUE,
            .default = FALSE
        ),
        LandFlag = case_when(
            startsWith(as.character(Building_Number), "N") ~ TRUE,
            startsWith(as.character(Building_Number), "B") ~ FALSE,
            .default = FALSE
        ),
        .after = AddressLink
    ) |>
    filter(!UOM %in% c("M2", "HA")) |>
    group_by(Address) |>
    summarise(
        AddressLink = first(AddressLink),
        BuildingCount = sum(BuildingFlag),
        LandCount = sum(LandFlag),
        City = first(City),
        Total_Rentable_Area = sum(Total_Rentable_Area),
        Latitude = first(Latitude),
        Longitude = first(Longitude)
    )
