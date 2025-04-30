library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)
library(lubridate)

options(scipen = 999)
# R_Facility_Table
R_Facility_Table <- read.csv(here("PBI/data/R_Facility_Table.csv"))

# Province Wide Occupancy ####
pwor <- read_xlsx(
  here(
    "data/ho-rplm-report-pmr3345-province-wide-occupancy-2025-04-03152308.604.xlsx"
  ),
  start_row = 3
) |>
  mutate(
    Identifier = `Building/Land #`,
    Contract = case_when(
      Tenure == "LEASED" &
        `Contract / Building / Property Number` == "N/A" ~
        NA,
      Tenure == "LEASED" ~ `Contract / Building / Property Number`,
      .default = NA
    ),
    .before = everything()
  ) |>
  mutate(
    TotalRentableArea = as.numeric(gsub(
      "[^0-9.-]",
      "",
      `Total Rentable Area Leased/Owned by RPD`
    )),
    InventoryAllocatedRentableArea = as.numeric(gsub(
      "[^0-9.-]",
      "",
      `Inventory Allocated Rentable Area`
    )),
    CustomerAllocationPercentage = as.numeric(
      `Customer Allocation Percentage of Location`
    ),
    .after = `Primary Use`
  ) |>
  select(
    -c(
      `Total Rentable Area Leased/Owned by RPD`,
      `Inventory Allocated Rentable Area`,
      `Customer Allocation Percentage of Location`,
      `Building/Land #`,
      `Contract / Building / Property Number`
    )
  ) |>
  rename(
    LeaseExpiryDate = `Lease Expiry Date`,
    PrimaryUse = `Primary Use`,
    ContractCode = `Contract Code`,
    AgreementType = `Agreement Type`,
    CustomerCategory = `Customer Category`
  ) |>
  filter(!UOM %in% c("M2", "HA")) |>
  mutate(PrimaryUse = toTitleCase(tolower(PrimaryUse))) |>
  distinct() |>
  filter(
    Tenure == "OWNED" |
      Address %in%
        c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Avenue") # Strategic Leases
  ) |>
  select(-c(City, Address)) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  mutate(
    PrimaryUse = case_when(
      is.na(PrimaryUse.x) ~ PrimaryUse.y,
      .default = PrimaryUse.x
    ),
    .after = FacilityType
  ) |>
  select(-c(PrimaryUse.x, PrimaryUse.y)) |>
  relocate(
    Identifier,
    Address,
    City,
    .before = everything()
  ) |>
  select(
    Identifier,
    Name,
    Address,
    City,
    TotalRentableArea,
    InventoryAllocatedRentableArea
  ) |>
  group_by(Identifier) |>
  summarise(
    Name = first(Name),
    Address = first(Address),
    City = first(City),
    TotalRentableArea = first(TotalRentableArea),
    AllocatedRentableArea = sum(InventoryAllocatedRentableArea, na.rm = TRUE)
  ) |>
  ungroup() |>
  mutate(
    AllocatedRentableArea = case_when(
      AllocatedRentableArea > TotalRentableArea ~ TotalRentableArea,
      .default = AllocatedRentableArea
    )
  ) |>
  mutate(
    AllocationPercentage = round(
      AllocatedRentableArea / TotalRentableArea,
      digits = 2
    )
  )

write.csv(
  pwor,
  here("PBI/data/R_ProvWideOccupancy_Table.csv"),
  row.names = FALSE
)

# Customer Agreement Report ####
car <- read_xlsx(
  here(
    "data/csr0001-2025-03-11113053.016.xlsx"
  )
) |>
  rename(
    Identifier = Location,
    ParticipatesInPAM = `Participates In PAM`,
    UploadCharges = `Upload Charges`,
    AgreementNumber = `Agreement #`,
    AgreementStatus = `Agr Status`,
    CAR_FiscalYear = `Fiscal Year`,
    AgreementType = `Agreement Type`,
    LeaseStart = `Lease Start`,
    LeaseExpiry = `Lease Expiry`,
    AgreementDurationEndDate = `Agreement Duration End Date`,
    BuildingMetricRentableArea = `Rentable Area - Building metric`,
    BuildingMetricAppropriatedArea = `Appropriated Area - Building metric`,
    BuildingMetricBillableArea = `Billable Area - Building metric`,
    LandMetricArea = `Area - Land metric`,
    LandMetricAppropriatedArea = `Appropriated Area - Land metric`,
    LandMetricBillableArea = `Billable Area - Land metric`,
    ParkingTotalStalls = `Total Parking Stalls`,
    ParkingAppropriatedArea = `Appropriated Area - Parking`,
    ParkingBillableStalls = `Billable Parking Stalls`,
    BaseRent = `Base rent`,
    OMTotal = `O&M Total`,
    OM = `O&M`,
    LLOM = `LL O&M`,
    OMAdmin = `O&M Admin`,
    PrkCost = `Prk Cost`,
    LLAdmin = `LL Admin`,
    UtilAdmin = `Util Admin`,
    TaxAdmin = `Tax Admin`,
    TotalAdmin = `Total Admin`,
    AdminFee = `Admin Fee`,
    AnnualCharge = `Annual Charge`
  ) |>
  select(
    -c(
      ParticipatesInPAM,
      UploadCharges,
      Description,
      AgreementDurationEndDate
    )
  ) |>
  mutate(across(
    BuildingMetricRentableArea:AnnualCharge,
    ~ as.numeric(gsub(
      "[^0-9.-]",
      "",
      .x
    ))
  )) |>
  rename(City.x = City, Address.x = Address) |>
  left_join(
    R_Facility_Table,
    by = join_by(
      Identifier
    )
  ) |>
  # 4000 & 4001 Seymour and 1011 4th Ave aren't filtered or don't exist in table
  filter(!grepl("^L", Lease)) |>
  filter(AgreementStatus == "Active") |>
  filter(AgreementType != "Non Prop") |>
  filter(!is.na(Address)) |>
  select(-c(Address.x, City.x)) |>
  rename(
    FiscalYear = CAR_FiscalYear
  ) |>
  relocate(
    Identifier,
    Name,
    Address,
    City,
    .before = everything()
  ) |>
  select(-c(PropertyCode:GeoPrecision))

write.csv(
  car,
  here("PBI/Data/R_CustomerAgreementReport.csv"),
  row.names = FALSE
)

# Fiscal table ####
fiscal <- read_xlsx(
  here(
    #"data/ho-rplm-contract-by-fiscal-by-component-2025-03-25092941.451.xlsx"
    "data/contract-by-fiscal-by-component-2025-04-30.xlsx"
  ),
  start_row = 3
) |>
  rename_with(~ gsub(" ", "", .x)) |>
  select(-c(starts_with("2324"), City)) |>
  rename_with(~ gsub("2425Year", "FY2425", .x)) |>
  rename_with(~ gsub("&", "", .x)) |>
  rename_with(~ gsub("%", "", .x)) |>
  rename(Identifier = PrimaryLocation) |>
  mutate(across(
    FY2425RentableArea:Variance,
    ~ as.numeric(gsub(
      "[^0-9.-]",
      "",
      .x
    ))
  )) |>
  # test <- fiscal |> filter(PrimaryLocation != Identifier)
  # test <- fiscal |> select(Identifier) |> distinct() |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  filter(!is.na(Address)) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(PropertyCode:GeoPrecision))

write.csv(
  fiscal,
  here("PBI/Data/R_FiscalTable.csv"),
  row.names = FALSE
)

# VFA Data ####
vfa <- read_xlsx(
  here(
    "data/Asset List.xlsx"
  )
) |>
  rename_with(str_to_title, everything()) |>
  filter(!grepl("Archive", Regionname)) |> # remove values marked as archived.
  select(-c(Region_is_archived)) |>
  filter(!grepl("^x", Assetnumber)) |> # some archived values have x or xx for Assetnumber
  mutate(
    Assetnumber = case_when(
      Assetnumber == "N2000527 / N2000497" ~ "N2000527",
      .default = Assetnumber
    ),
    Assetname = case_when(
      Assetname == "St Ann's Site" ~ "St Ann's Site N2000527 and N2000497",
      .default = Assetname
    )
  ) |>
  left_join(R_Facility_Table, by = join_by(Assetnumber == Identifier)) |>
  filter(!is.na(Address)) |>
  select(
    Identifier = Assetnumber,
    Name,
    Address,
    City,
    Regioneid,
    Regionname,
    Campuseid,
    Campusname,
    Asseteid,
    Assetname,
    Assetsize,
    Replacementvalue,
    Use,
    YearConstructed = `Year Constructed`,
    Age,
    Fci,
    FciCost = Fcicost,
    Currency,
    AssetUnit = Assetunit,
    AssetType = Assettype,
    Ri,
    RiCost = Ricost,
    AssetCostUnit = Asset_costunit
  )

write.csv(
  vfa,
  here("PBI/Data/R_VFATable.csv"),
  row.names = FALSE
)

# Project Costs Data ####
ProjectCostsCBRE <- read_xlsx(
  here("data/ProjectReportApril152025.xlsx"),
  sheet = "CBRE"
) |>
  mutate(Source = "CBRE", .before = everything())

ProjectCostsWDS <- read_xlsx(
  here("data/ProjectReportApril152025.xlsx"),
  sheet = "WDS"
) |>
  mutate(Source = "WBS", .before = everything())

ProjectCosts <- rbind(ProjectCostsCBRE, ProjectCostsWDS) |>
  rename_with(~ gsub(" ", "", .), everything()) |>
  rename_with(~ gsub("/[\\n]?", "", perl = TRUE, .), everything()) |>
  rename_with(~ gsub("[/()#%]", "", .), everything()) |>
  rename(Identifier = BuildingID) |>
  group_by(Identifier) |>
  summarise(
    Region = first(Region),
    N_Projects = n(),
    Total_Paid = sum(Paid, na.rm = TRUE)
  ) |>
  filter(
    !is.na(Identifier) & !Identifier %in% c("9999100", "9999200", "9999300")
  ) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  filter(!is.na(Address)) |>
  select(Identifier, Name, Address, City, Region, N_Projects, Total_Paid) |>
  mutate(Total_Paid = round(Total_Paid, digits = 2))

write.csv(
  ProjectCosts,
  here("PBI/Data/R_ProjectCosts_Table.csv"),
  row.names = FALSE
)

# zoning ####
# check for vacant land ####

vacant_land <- read_xlsx(
  here("data/PMR3678-Active-Land-No-Active-Buildings.xlsx"),
  start_row = 3
) |>
  fill(`City ID`, .direction = "down") |>
  mutate(
    City = toTitleCase(tolower(`City ID`)),
    .keep = "unused",
    .after = Address
  ) |>
  rename_with(~ gsub(" ", "", .), cols = everything()) |>
  group_by(City) |>
  fill(Address, .direction = "down") |>
  ungroup() |>
  filter(!is.na(Address)) |>
  filter(!is.na(LandName)) |>
  filter(Tenure == "OWNED") |>
  filter(OccupancyStatus == "VACANT")

# Zoning ####

zoning_buildings <- read_xlsx(here("data/RPD_Buildings_Zoning.xlsx")) |>
  filter(
    Tenure == "Owned" |
      Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  ) |>
  select(
    Identifier = Building_Number,
    ZoneCode = ZONE_CODE,
    ZoneClass = ZONE_CLASS,
    ParcelName = PARCEL_NAME,
    ParcelStatus = PARCEL_STATUS,
    PlanNumber = PLAN_NUMBER,
    PID,
    PIN
  ) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(PropertyCode:GeoPrecision))

zoning_lands <- read_xlsx(here("data/RPD_Land_Zoning.xlsx")) |>
  filter(
    Tenure == "Owned" |
      Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  ) |>
  select(
    Identifier = Land_Number,
    ZoneCode = ZONE_CODE,
    ZoneClass = ZONE_CLASS,
    ParcelName = PARCEL_NAME,
    ParcelStatus = PARCEL_STATUS,
    PlanNumber = PLAN_NUMBER,
    PID,
    PIN
  ) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(PropertyCode:GeoPrecision))

zoning <- rbind(zoning_buildings, zoning_lands) |> distinct()

write.csv(
  zoning,
  here("PBI/Data/R_Zoning_Table.csv"),
  row.names = FALSE
)

# Work Orders ####

workorders <- read.csv(here("data/WO Detail List_Full Data_data.csv")) |>
  select(
    Identifier = `Property.ID`,
    PriorityCode = `Priority.Code`,
    SLACompletionStatus = `SLA.Completion.Status`,
    WorkCategory = `Category.Code.Desc`,
    CompletionDate = `Actual.Completed.Ts`
  ) |>
  mutate(CompletionDate = gsub("( [0-9]+:.*)", "", CompletionDate)) |>
  mutate(CompletionDate = as.Date(CompletionDate, format = "%m/%d/%Y")) |>
  filter(CompletionDate >= as.Date("2024-03-31", format = ("%Y-%m-%d")))
