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
        c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave") # Strategic Leases
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
  mutate(
    TotalRentableLandArea = case_when(
      UOM == "M2" ~ NA,
      .default = TotalRentableArea
    ),
    TotalRentableBuildingArea = case_when(
      UOM == "HA" ~ NA,
      .default = TotalRentableArea
    ),
    InventoryAllocatedRentableLandArea = case_when(
      UOM == "M2" ~ NA,
      .default = InventoryAllocatedRentableArea
    ),
    InventoryAllocatedRentableBuildingArea = case_when(
      UOM == "HA" ~ NA,
      .default = InventoryAllocatedRentableArea
    ),
    .before = TotalRentableArea
  ) |>
  select(
    Identifier,
    Name,
    Address,
    City,
    CustomerCategory,
    Division,
    Department,
    UOM,
    TotalRentableLandArea,
    InventoryAllocatedRentableLandArea,
    TotalRentableBuildingArea,
    InventoryAllocatedRentableBuildingArea,
    UsableArea,
    LandArea
  ) |>
  mutate(
    CustomerCategory = case_when(
      Department %in% c("RPD VACANT SPACE", "Vacancy") ~ "Vacant",
      .default = CustomerCategory
    )
  ) |>
  filter(!CustomerCategory == "Vacant") |>
  group_by(Identifier) |>
  summarise(
    Name = first(Name),
    Address = first(Address),
    City = first(City),
    TotalRentableLandArea = first(TotalRentableLandArea, na_rm = TRUE),
    ,
    InventoryAllocatedRentableLandArea = sum(
      InventoryAllocatedRentableLandArea,
      na.rm = TRUE
    ),
    TotalRentableBuildingArea = first(TotalRentableBuildingArea, na_rm = TRUE),
    ,
    InventoryAllocatedRentableBuildingArea = sum(
      InventoryAllocatedRentableBuildingArea,
      na.rm = TRUE
    )
  ) |>
  ungroup() |>
  mutate(
    InventoryAllocatedRentableLandArea = case_when(
      InventoryAllocatedRentableLandArea > TotalRentableLandArea ~
        TotalRentableLandArea,
      .default = InventoryAllocatedRentableLandArea
    ),
    InventoryAllocatedRentableBuildingArea = case_when(
      InventoryAllocatedRentableBuildingArea > TotalRentableBuildingArea ~
        TotalRentableBuildingArea,
      .default = InventoryAllocatedRentableBuildingArea
    )
  ) |>
  mutate(
    AllocationPercentageLand = round(
      InventoryAllocatedRentableLandArea / TotalRentableLandArea,
      digits = 2
    ),
    AllocationPercentageBuilding = round(
      InventoryAllocatedRentableBuildingArea / TotalRentableBuildingArea,
      digits = 2
    )
  )

write.csv(
  pwor,
  here("PBI/data/R_ProvWideOccupancy_Table.csv"),
  row.names = FALSE
)

# test <- pwor |>
#   select(Identifier) |>
#   mutate(pwor = "Yes") |>
#   right_join(R_Facility_Table, by = join_by(Identifier)) |>
#   filter(is.na(pwor))

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
    Client,
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
  # group_by(Identifier, Name, Address, City, FiscalYear) |>
  # summarise(across(BuildingRentableArea:AnnualCharge, ~ sum(., na.rm = TRUE)))
  select(
    Identifier,
    Name,
    Address,
    City,
    FiscalYear,
    Client,
    AgreementType,
    AgreementNumber,
    BuildingRentableArea:AnnualCharge
  )

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
  select(-c(PropertyCode:GeoPrecision)) |>
  select(Identifier:City, FY2425OM:FY2425LLOM, FY2425TotalCost)

write.csv(
  fiscal,
  here("PBI/Data/R_FiscalTable.csv"),
  row.names = FALSE
)

# VFA Data ####
vfa <- read_xlsx(
  here(
    # "data/Asset List.xlsx"
    "data/VFA FCI June.xlsx"
  )
) |>
  select(
    AssetNumber = `Asset - Number`,
    City = `Asset - City`,
    FCI = `Asset - FCI`,
    Address = `Asset - Address 1`
  ) |>
  # rename_with(str_to_title, everything()) |>
  # filter(!grepl("Archive", Regionname)) |> # remove values marked as archived.
  # select(-c(Region_is_archived)) |>
  # filter(!grepl("^x", Assetnumber)) |> # some archived values have x or xx for Assetnumber
  mutate(
    AssetNumber = case_when(
      AssetNumber == "N2000527 / N2000497" ~ "N2000527",
      .default = AssetNumber
    )
    # ,
    # Assetname = case_when(
    #   Assetname == "St Ann's Site" ~ "St Ann's Site N2000527 and N2000497",
    #   .default = Assetname
    # )
  ) |>
  left_join(R_Facility_Table, by = join_by(AssetNumber == Identifier)) |>
  select(
    Identifier = AssetNumber,
    Name,
    Address = Address.y,
    City = City.y,
    FCI
  ) |>
  filter(!is.na(Address))

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
  select(
    Identifier,
    Name,
    Address,
    City,
    Region,
    N_Projects,
    Total_Paid,
    BuildingRentableArea,
    LandArea
  ) |>
  mutate(Total_Paid = round(Total_Paid, digits = 2)) |>
  mutate(
    CostPerSqM = case_when(
      startsWith(Identifier, "B") ~ Total_Paid / BuildingRentableArea,
      .default = NA
    ),
    CostPerHA = case_when(
      startsWith(Identifier, "N") ~ Total_Paid / LandArea,
      .default = NA
    )
  )

write.csv(
  ProjectCosts,
  here("PBI/Data/R_ProjectCosts_Table.csv"),
  row.names = FALSE
)

# check for vacant land ####
vacant_land <- read_xlsx(
  # here("data/PMR3678-Active-Land-No-Active-Buildings.xlsx"),
  here("data/pmr3678-active-lands-no-buildings-2025-05-09.xlsx"),
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
  filter(OccupancyStatus == "VACANT") |>
  left_join(R_Facility_Table, by = join_by(LandName == Identifier))

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
    PIN,
    RegionalDistrict = REGIONAL_DISTRICT
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
    PIN,
    RegionalDistrict = REGIONAL_DISTRICT
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
  ) #|>
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
  filter(!is.na(Address)) |>
  pivot_wider(names_from = PaymentType, values_from = TotalPaid)

write.csv(
  tripay,
  here("PBI/Data/R_OMnUtilities_Table.csv"),
  row.names = FALSE
)

# Contaminated Sites ####

csites <- read_xlsx(
  here("data/Contaminated_sites_inventory_April_2025.xlsx"),
  sheet = "Apr 2025"
)

contaminated_building <- csites |>
  rename_with(~ gsub(" ", "", .), .cols = everything()) |>
  filter(
    BuildingTenure == "Owned" |
      LandTenure == "Owned" |
      BuildingAddress %in%
        c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  ) |>
  select(
    -c(
      City,
      BuildingRentableArea,
      BuildingDescription,
      BuildingAddress,
      LandArea,
      LandDescription,
      LandAddress
    )
  ) |>
  mutate(
    Building = case_when(
      nchar(Building) == 7 ~ paste0("B", Building),
      nchar(Building) == 6 ~ paste0("B0", Building),
      nchar(Building) == 5 ~ paste0("B00", Building)
    )
  ) |>
  left_join(R_Facility_Table, by = join_by(Building == Identifier)) |>
  filter(!is.na(Name)) |>
  select(-c(Land, LandTenure, BuildingTenure, `Ministry/Non-Ministry`)) |>
  rename(
    Identifier = Building,
    SiteInvestigated = `SiteInvestigated?-ReportAvailable`,
    ContaminationStatus = `IsSiteContaminated?`,
    StatusOfRemediation = StatusofRemediation
  ) |>
  relocate(Name, Address, City, .after = Identifier) |>
  select(-c(AddressEdit:GeoPrecision)) |>
  mutate(
    SiteInvestigated = case_when(
      SiteInvestigated == "no" ~ "No",
      SiteInvestigated == "sample" ~ "Sample",
      .default = SiteInvestigated
    ),
    ContaminationStatus = case_when(
      ContaminationStatus == "no" ~ "No",
      ContaminationStatus == "yes" ~ "Yes",
      ContaminationStatus == "unknown" ~ "Unknown",
      .default = ContaminationStatus
    )
  )

write.csv(
  contaminated_building,
  here("PBI/data/R_Contaminated_Buildings_Table.csv")
)


# Make Ratings Table ####
R_HQ_Telework_Table <- read.csv(here("PBI/data/R_HQ_Telework_Table.csv"))

vfa <- read.csv(here("PBI/Data/R_VFATable.csv"))

fiscal <- read.csv(here("PBI/Data/R_FiscalTable.csv"))

utilization <- R_HQ_Telework_Table |>
  filter(DaysInOfficePresumed != "" & FTECount > 0) |>
  group_by(Address) |>
  summarise(
    FTECountPerAddress = sum(FTECount),
    DaysInOfficePresumed = sum(DaysInOfficePresumed)
  ) |>
  ungroup() |>
  mutate(AvgDaysInOffice = DaysInOfficePresumed / 5) |>
  mutate(InPersonPct = AvgDaysInOffice / FTECountPerAddress)

pwor_building <- read.csv(here("PBI/data/R_ProvWideOccupancy_Table.csv")) |>
  select(Identifier, TotalRentableBuildingArea, AllocationPercentageBuilding)

ratings <- R_Facility_Table |>
  select(
    Identifier,
    Name,
    Address,
    City,
    PrimaryUse,
    FacilityType,
    FacilityBuildingRentableArea = BuildingRentableArea
  ) |>
  filter(startsWith(Identifier, "B")) |>
  left_join(vfa, join_by(Identifier, Name, Address, City)) |>
  left_join(
    fiscal,
    join_by(Identifier, Name, Address, City)
  ) |>
  select(
    Identifier,
    Name,
    Address,
    City,
    PrimaryUse,
    FacilityBuildingRentableArea,
    FCI,
    FY2425TotalCost
  ) |>
  mutate(
    BuildingCategory = case_when(
      PrimaryUse %in% c("Office", "Trailer Office", "Yard Office") ~ "Office",
      PrimaryUse %in% c("Courthouse") ~ "Courthouse",
      PrimaryUse %in% c("Jail") ~ "Jail",
      PrimaryUse %in%
        c("Storage Open", "Storage Small", "Storage Vehicle", "Warehouse") ~
        "Storage",
      PrimaryUse %in% c("Ambulance", "Health Unit", "Laboratory") ~ "Health",
      PrimaryUse %in%
        c(
          "Residential Detached",
          "Residential Multi",
          "Dormitory",
          "Mobile Home"
        ) ~
        "Residential",
      .default = "Other"
    ),
    .after = PrimaryUse
  ) |>
  left_join(pwor_building, by = join_by(Identifier)) |>
  mutate(
    BuildingRentableArea = case_when(
      is.na(TotalRentableBuildingArea) ~ FacilityBuildingRentableArea,
      .default = TotalRentableBuildingArea
    ),
    AllocationPercentageBuilding = case_when(
      is.na(AllocationPercentageBuilding) ~ 0,
      .default = AllocationPercentageBuilding
    )
  ) |>
  select(-c(TotalRentableBuildingArea, FacilityBuildingRentableArea)) |>
  relocate(
    BuildingRentableArea,
    AllocationPercentageBuilding,
    .after = BuildingCategory
  ) |>
  mutate(
    ConditionRating = case_when(
      FCI <= 0.2 ~ "Excellent",
      FCI <= 0.4 ~ "Good",
      FCI <= 0.6 ~ "Fair",
      FCI <= 0.8 ~ "Poor",
      FCI > 0.8 ~ "Dire",
      is.na(FCI) ~ "Fair"
    ),
    .after = FCI
  ) |>
  mutate(
    OccupancyAdjustedCostPerSqM = (FY2425TotalCost / BuildingRentableArea) /
      AllocationPercentageBuilding,
    .after = FY2425TotalCost
  ) |>
  mutate(
    Occupancy = case_when(
      AllocationPercentageBuilding == 0 ~ FALSE,
      .default = TRUE
    ),
    .after = AllocationPercentageBuilding
  ) |>
  group_by(BuildingCategory, Occupancy) |>
  mutate(
    percentile = percent_rank(OccupancyAdjustedCostPerSqM),
    FinancialRating = case_when(
      percentile <= 0.20 ~ "Excellent",
      percentile <= 0.40 ~ "Good",
      percentile <= 0.60 ~ "Fair",
      percentile <= 0.80 ~ "Poor",
      percentile > 0.80 ~ "Dire",
      is.na(percentile) ~ "Fair"
    ),
    .after = OccupancyAdjustedCostPerSqM
  ) |>
  ungroup() |>
  mutate(
    FinancialRating = case_when(
      AllocationPercentageBuilding == 0 ~ "Dire",
      .default = FinancialRating
    )
  ) |>
  left_join(utilization, by = join_by(Address)) |>
  mutate(
    OccupancyAdjustedUtilizationRating = InPersonPct *
      AllocationPercentageBuilding
  ) |>
  mutate(
    UtilizationRating = case_when(
      OccupancyAdjustedUtilizationRating <= 0.20 ~ "Dire",
      OccupancyAdjustedUtilizationRating <= 0.40 ~ "Poor",
      OccupancyAdjustedUtilizationRating <= 0.60 ~ "Fair",
      OccupancyAdjustedUtilizationRating <= 0.80 ~ "Good",
      OccupancyAdjustedUtilizationRating > 0.80 ~ "Excellent",
      is.na(OccupancyAdjustedUtilizationRating) ~ "Fair"
    ),
    .after = OccupancyAdjustedUtilizationRating
  ) |>
  mutate(
    ConditionScore = case_when(
      ConditionRating == "Dire" ~ 1,
      ConditionRating == "Poor" ~ 2,
      ConditionRating == "Fair" ~ 3,
      ConditionRating == "Good" ~ 4,
      ConditionRating == "Excellent" ~ 5,
      is.na(ConditionRating) ~ 3,
      TRUE ~ 3
    ),
    FinancialScore = case_when(
      FinancialRating == "Dire" ~ 1,
      FinancialRating == "Poor" ~ 2,
      FinancialRating == "Fair" ~ 3,
      FinancialRating == "Good" ~ 4,
      FinancialRating == "Excellent" ~ 5,
      is.na(FinancialRating) ~ 3,
      TRUE ~ 3
    ),
    UtilizationScore = case_when(
      UtilizationRating == "Dire" ~ 1,
      UtilizationRating == "Poor" ~ 2,
      UtilizationRating == "Fair" ~ 3,
      UtilizationRating == "Good" ~ 4,
      UtilizationRating == "Excellent" ~ 5,
      is.na(UtilizationRating) ~ 3,
      TRUE ~ 3
    )
  ) |>
  rowwise() |>
  mutate(
    OverallScore = sum(ConditionScore, FinancialScore, UtilizationScore)
  ) |>
  ungroup() |>
  group_by(Address) |>
  mutate(
    TotalPropertyRentableArea = sum(BuildingRentableArea, na.rm = TRUE),
    .after = BuildingRentableArea
  ) |>
  ungroup() |>
  mutate(
    WeightedAssetScore = OverallScore *
      (BuildingRentableArea / TotalPropertyRentableArea)
  ) |>
  group_by(Address) |>
  mutate(PropertyWeightedScore = sum(WeightedAssetScore)) |>
  ungroup() |>
  mutate(
    OverallRating = case_when(
      OverallScore <= 5.4 ~ "Dire",
      OverallScore <= 7.8 ~ "Poor",
      OverallScore <= 10.2 ~ "Fair",
      OverallScore <= 12.6 ~ "Good",
      OverallScore <= 15 ~ "Excellent"
    ),
    .after = OverallScore
  )

rankset <- ratings |>
  select(Address, PropertyWeightedScore) |>
  group_by(Address) |>
  slice(1) |>
  ungroup() |>
  arrange(PropertyWeightedScore) |>
  mutate(Rank = rank(PropertyWeightedScore, ties.method = "min"))

ratingsFinal <- ratings |>
  left_join(rankset, by = join_by(Address, PropertyWeightedScore))

write.csv(ratingsFinal, here("PBI/data/R_Ratings_Table.csv"), row.names = FALSE)
