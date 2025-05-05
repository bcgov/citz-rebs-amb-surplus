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
    "data/pmr3345-province-wide-occupancy-2025-05-01.xlsx"
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
    InventoryAllocatedRentableArea,
    BuildingRentableArea,
    UsableArea,
    LandArea
  ) |>
  group_by(Identifier) |>
  summarise(
    Name = first(Name),
    Address = first(Address),
    City = first(City),
    TotalRentableArea = first(TotalRentableArea, na_rm = TRUE),
    AllocatedRentableArea = sum(InventoryAllocatedRentableArea, na.rm = TRUE),
    BuildingRentableArea = first(BuildingRentableArea),
    UsableArea = first(UsableArea),
    LandArea = first(LandArea)
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
car24_26 <- read_xlsx(
  here(
    "data/csr0001-FY2425FY2526.xlsx"
  ),
  start_row = 3
)

car22_24 <- read_xlsx(
  here(
    "data/csr0001-FY2223FY2324.xlsx"
  ),
  start_row = 3
)
car20_22 <- read_xlsx(
  here(
    "data/csr0001-FY2021FY2122.xlsx"
  ),
  start_row = 3
)

car <- rbind(car20_22, car22_24, car24_26) |>
  select(
    Identifier = Location,
    AgreementNumber = `Agreement #`,
    AgreementStatus = `Agr Status`,
    FiscalYear = `Fiscal Year`,
    AgreementType = `Agreement Type`,
    Lease,
    LeaseStart = `Lease Start`,
    LeaseExpiry = `Lease Expiry`,
    AgreementDurationEndDate = `Agreement Duration End Date`,
    BuildingRentableArea = `Rentable Area - Building metric`,
    BuildingAppropriatedArea = `Appropriated Area - Building metric`,
    BuildingBillableArea = `Billable Area - Building metric`,
    LandArea = `Area - Land metric`,
    LandAppropriatedArea = `Appropriated Area - Land metric`,
    LandBillableArea = `Billable Area - Land metric`,
    ParkingTotalStalls = `Total Parking Stalls`,
    ParkingAppropriatedArea = `Appropriated Area - Parking`,
    ParkingBillableStalls = `Billable Parking Stalls`,
    BaseRent = `Base rent`,
    OM = `O&M`,
    Utilities,
    OMTotal = `O&M Total`,
    LLOM = `LL O&M`,
    Tax,
    Amort,
    PrkCost = `Prk Cost`,
    LLAdmin = `LL Admin`,
    OMAdmin = `O&M Admin`,
    UtilAdmin = `Util Admin`,
    TaxAdmin = `Tax Admin`,
    TotalAdmin = `Total Admin`,
    AdminFee = `Admin Fee`,
    AnnualCharge = `Annual Charge`
  ) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(PropertyCode:GeoPrecision)) |>
  rename(
    LandArea = LandArea.x,
    BuildingRentableArea = BuildingRentableArea.x
  ) |>
  # Address in R_Facility_Table
  filter(!is.na(Address)) |>
  # Remove any leases that aren't strategic
  # This only seems to remove leases associated with buildings in the owned portfolio.
  # Usually parking only but sometimes building as well.
  # filter(
  #   !grepl("^L", Lease) |
  #     Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  # ) |>
  filter(AgreementStatus == "Active") |>
  filter(AgreementType != "Non Prop") |>
  mutate(across(
    BuildingRentableArea:AnnualCharge,
    ~ as.numeric(gsub(
      "[^0-9.-]",
      "",
      .x
    ))
  )) |>
  group_by(Identifier, Name, Address, City, FiscalYear) |>
  summarise(across(BuildingRentableArea:AnnualCharge, ~ sum(., na.rm = TRUE)))

write.csv(
  car,
  here("PBI/Data/R_CustomerAgreementReport.csv"),
  row.names = FALSE
)

# Fiscal table ####
# Only most current year appears to be available. Not sure reliability or what is in 2526 year
fiscal <- read_xlsx(
  here(
    #"data/ho-rplm-contract-by-fiscal-by-component-2025-03-25092941.451.xlsx"
    # "data/contract-by-fiscal-by-component-2025-04-30.xlsx"
    "data/contract-by-fiscal-by-component-23-25-2025-05-05.xlsx"
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
  # test <- fiscal |> filter(ContractName != Identifier)
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
    CreationDate = `Creation.Date`,
    CompletionDate = `Actual.Completed.Ts`
  ) |>
  mutate(CompletionDate = gsub("( [0-9]+:.*)", "", CompletionDate)) |>
  mutate(
    CompletionDate = as.Date(CompletionDate, format = "%m/%d/%Y"),
    CreationDate = as.Date(CreationDate, format = "%m/%d/%Y")
  ) |>
  mutate(
    FYCreation = case_when(
      CreationDate |> timetk::between_time('2025-04-01', '2026-03-31') ~
        "FY2526",
      CreationDate |> timetk::between_time('2024-04-01', '2025-03-31') ~
        "FY2425",
      CreationDate |> timetk::between_time('2023-04-01', '2024-03-31') ~
        "FY2324",
      CreationDate |> timetk::between_time('2022-04-01', '2023-03-31') ~
        "FY2223",
      CreationDate |> timetk::between_time('2021-04-01', '2022-03-31') ~
        "FY2122",
      CreationDate |> timetk::between_time('2020-04-01', '2021-03-31') ~
        "FY2021",
    )
  ) |>
  group_by(Identifier, FYCreation) |>
  summarise(
    WorkOrderCount = n()
  ) |>
  pivot_wider(names_from = FYCreation, values_from = WorkOrderCount) |>
  mutate(across(starts_with("FY"), ~ replace_na(., 0))) |>
  pivot_longer(
    starts_with("FY"),
    names_to = "FYCreation",
    values_to = "WorkOrderCount"
  ) |>
  ungroup() |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(PropertyCode:GeoPrecision)) |>
  filter(!is.na(Address))

write.csv(
  workorders,
  here("PBI/Data/R_WorkOrders_Table.csv"),
  row.names = FALSE
)

# Climate Risk ####
climate <- read_xlsx(
  here("data/RPD Building Climate Risk Scores.xlsx"),
  start_row = 2
) |>
  rename_with(~ gsub("_", "", .), .cols = everything()) |>
  rename_with(~ gsub(" ", "", .), .cols = everything()) |>
  rename_with(~ gsub("-", "", .), .cols = everything()) |>
  select(-c(OBJECTID, Address, City, PostalCode)) |>
  rename(
    Location = LOCATION,
    LastUpdate = lastupdate,
    Identifier = BuildingNumber,
    SiteCode = ComplexNumber,
    PropertyCode = LandNumber
  ) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  filter(!is.na(Address)) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(PropertyCode.y:GeoPrecision)) |>
  rename_with(~ gsub("\\.x", "", .), .cols = everything())

write.csv(
  climate,
  here("PBI/Data/R_ClimateRisk_Table.csv"),
  row.names = FALSE
)

# @Real Data ####

tripaymentlineitem <- read_xlsx(here(
  "data/T_TRIPAYMENTLINEITEM Owned Property PLI's May 5 2025.xlsx"
))

tripay <- tripaymentlineitem |>
  select(
    -c(
      TRIIDTX,
      AREARESNUMBERNU,
      ARENALEASEPAYMENTLI,
      TRISTATUSCL,
      TRISTATUSCLOBJID,
      T1_1091_OBJID,
      T1_SPEC_ID,
      T1_SYS_GUIID,
      TRICURRENCYUO,
      SYS_TYPE1,
      AREPAYMENTSCHEDULETYPE
    )
  ) |>
  rename_with(~ toTitleCase(tolower(.)), .cols = everything()) |>
  filter(Trinametx != "*NONPROP") |>
  mutate(
    PaidDate = as.Date(Areactualduedate, format = "%d-%b-%Y"),
    .after = Trinametx
  ) |>
  filter(PaidDate > "2020-03-31") |>
  rename(
    Identifier = Trinametx,
    PaymentType = Tripaymenttypecl,
    AmountPaid = Triactualamountnu,
    OrgPaid = Triremittoorganization
  ) |>
  mutate(
    FiscalYear = case_when(
      PaidDate |> timetk::between_time('2024-04-01', '2025-03-31') ~ "FY2425",
      PaidDate |> timetk::between_time('2023-04-01', '2024-03-31') ~ "FY2324",
      PaidDate |> timetk::between_time('2022-04-01', '2023-03-31') ~ "FY2223",
      PaidDate |> timetk::between_time('2021-04-01', '2022-03-31') ~ "FY2122",
      PaidDate |> timetk::between_time('2020-04-01', '2021-03-31') ~ "FY2021"
    ),
    .after = PaidDate
  ) |>
  group_by(Identifier, FiscalYear, PaymentType) |>
  summarise(TotalPaid = sum(AmountPaid)) |>
  ungroup() |>
  filter(PaymentType %in% c("O&M Total", "Utilities")) |>
  left_join(R_Facility_Table, by = join_by(Identifier)) |>
  relocate(Name, Address, City, FacilityType, .after = Identifier) |>
  select(-c(PropertyCode:GeoPrecision)) |>
  filter(!is.na(Address))

write.csv(
  tripay,
  here("PBI/Data/R_OMnUtilities_Table.csv"),
  row.names = FALSE
)
