library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)


# Prep datasets ####
## Province Wide Occupancy ####
PWOR <- read_xlsx(
  here(
    "data/ho-rplm-report-pmr3345-province-wide-occupancy-2025-04-03152308.604.xlsx"
  ),
  start_row = 3
)

cleanPWOR <- PWOR |>
  mutate(
    IdentifierPWOR = `Building/Land #`,
    IdentifierType = case_when(
      startsWith(as.character(`Building/Land #`), "N") ~ "Land",
      startsWith(as.character(`Building/Land #`), "B") ~ "Building",
      .default = NA
    ),
    ContractPWOR = case_when(
      Tenure == "LEASED" & `Contract / Building / Property Number` == "N/A" ~
        NA,
      Tenure == "LEASED" ~ `Contract / Building / Property Number`,
      .default = NA
    ),
    .before = everything()
  ) |>
  mutate(
    TotalRentableAreaPWOR = as.numeric(gsub(
      "[^0-9.-]",
      "",
      `Total Rentable Area Leased/Owned by RPD`
    )),
    InventoryAllocatedRentableAreaPWOR = as.numeric(gsub(
      "[^0-9.-]",
      "",
      `Inventory Allocated Rentable Area`
    )),
    CustomerAllocationPercentagePWOR = as.numeric(
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
  distinct()

write.csv(
  cleanPWOR,
  here("data/output/CleanedProvinceWideOccupancy.csv"),
  row.names = FALSE
)

# Minor issue with this is some of the places are off the road network
# e.g. Summit Of Stagleap or very vague addresses like Hwy 3 Sparwood
# However this only applies to 142 rows of which many are the same place just different building.
# lat/lon assignment is not far off.

## Customer Agreement Report ####
CAR <- read_xlsx(
  here(
    "data/csr0001-2025-03-11113053.016.xlsx"
  )
)

cleanCAR <- CAR |>
  rename(
    IdentifierCAR = Location,
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
    -c(ParticipatesInPAM, UploadCharges, Description, AgreementDurationEndDate)
  ) |>
  mutate(across(
    BuildingMetricRentableArea:AnnualCharge,
    ~ as.numeric(gsub(
      "[^0-9.-]",
      "",
      .x
    ))
  ))

write.csv(
  cleanCAR,
  here("data/output/CleanedCustomerAgreementReport.csv"),
  row.names = FALSE
)

## Fiscal Component ####
Fiscal <- read_xlsx(
  here(
    "data/ho-rplm-contract-by-fiscal-by-component-2025-03-25092941.451.xlsx"
  ),
  start_row = 3
)

cleanFiscal <- Fiscal |>
  rename_with(~ gsub(" ", "", .x)) |>
  select(-c(starts_with("2324"))) |>
  rename_with(~ gsub("2425Year", "FY2425", .x)) |>
  rename_with(~ gsub("&", "", .x)) |>
  rename_with(~ gsub("%", "", .x)) |>
  mutate(across(
    FY2425RentableArea:Variance,
    ~ as.numeric(gsub(
      "[^0-9.-]",
      "",
      .x
    ))
  ))

write.csv(
  cleanFiscal,
  here("data/output/CleanedFiscalReport.csv"),
  row.names = FALSE
)

## VFA Data ####
VFAData <- read_xlsx(
  here(
    "data/Asset List.xlsx"
  )
) |>
  rename_with(str_to_title, everything())

cleanVFA <- VFAData |>
  filter(!grepl("Archive", Regionname)) |> # remove values marked as archived.
  select(-c(Region_is_archived)) |>
  filter(!grepl("^x", Assetnumber)) # some archived values have x or xx for Assetnumber

write.csv(
  cleanVFA,
  here("data/output/CleanedVFAAssetList.csv"),
  row.names = FALSE
)

## Employee Data ####
Employee_GAL_BCGov = read_xlsx(here("data/2024-12-19 GAL Directory.xlsx")) |>
  select(
    Name,
    DisplayName = `Display Name`,
    FirstName = `First Name`,
    LastName = `Last Name`,
    City,
    Street,
    EmailAddress = `Email Address`, # check for empty or blank rows
    EmployeeID = `Employee ID`,
    JobTitle = `Job Title`,
    Company,
    Division,
    Department
  ) |>
  mutate(EmployeeID = as.numeric(EmployeeID))

GAL_set <- Employee_GAL_BCGov |>
  select(
    FirstName,
    LastName,
    EmailAddress, # check for empty or blank rows
    EmployeeID
  ) |>
  mutate(EmployeeID = as.numeric(EmployeeID)) |>
  distinct()

## Establishment Report ####
Employee_EstablishmentReportBCGov = read_xlsx(
  here("data/Establishment Report - All BCPS - 2025-02-24.xlsx"),
  sheet = "Establishment 2025-02-24"
) |>
  # Remove anyone without employment status, transferred out, or unapproved positions
  filter(`Empl Status` != "NULL") |>
  filter(!Type %in% c("Int Out", "Ext Out")) |>
  filter(`Position Status` == "Approved" | is.na(`Position Status`)) |>
  mutate(EmployeeID = as.numeric(`Empl ID`)) |>
  select(
    EmployeeID,
    City,
    Company = Organization,
    Division = `Program Division`,
    Department = `Program Branch`,
    DepartmentID = `DeptID`,
    BusinessUnit = `Business Unit`,
    Title,
    Job_Code = `Job Code`
  ) |>
  distinct()

Establishment_set <- Employee_EstablishmentReportBCGov |>
  select(
    EmployeeID,
    DepartmentID,
    BusinessUnit
  )

## Employee Headcount ####
Employee_HeadcountCITZ = read_xlsx(
  here("data/Headcount by Classification - All BCPS - 2025-02-24.xlsx"),
  sheet = "Headcount 2025-02-24"
) |>
  filter(empl_status == "Active" | is.na(empl_status)) |>
  filter(!is.na(emplid)) |>
  mutate(HIRE_DT = as.Date(HIRE_DT)) |>
  group_by(emplid) |>
  filter(HIRE_DT == max(HIRE_DT)) |>
  ungroup() |>
  mutate(EmployeeID = as.numeric(emplid)) |>
  distinct() |>
  select(
    EmployeeID,
    Name = name,
    EmailAddress = EMAILID,
    IDIR,
    BusinessUnit = business_unit,
    DepartmentID = deptid,
  ) |>
  mutate(EmailAddress = na_if(EmailAddress, "NULL"), IDIR = na_if(IDIR, "NULL"))

Headcount_set <- Employee_HeadcountCITZ |>
  select(EmployeeID, EmailAddress, IDIR, BusinessUnit, DepartmentID)

## Employee Telework ####
Employee_TeleworkBCGovPSA = read_xlsx(
  here(
    "data/Staff Data For Telework Data Collection - All BCPS - 2025-03-06.xlsx"
  ),
  sheet = "Telework All Orgs 2025-03-06"
) |>
  filter(Empl_Status == "Active") |>
  mutate(Street = paste(`Work Address1`, `Work Address2`)) |>
  mutate(EMAILID = na_if(EMAILID, "NULL")) |>
  distinct() |>
  select(
    FirstName = FIRST_NAME,
    LastName = LAST_NAME,
    EmployeeID = EMPLID,
    EmailAddress = EMAILID,
    IDIR,
    DepartmentID = DEPTID,
    Street,
    City = `Work City`,
    DaysInOffice = `Days In Office`,
    DaysTeleworking = `Days Teleworking`,
    Sunday,
    Monday,
    Tuesday,
    Wednesday,
    Thursday,
    Friday,
    Saturday,
    Company = Program,
    Division = `Program Division`,
    Branch = `Program Branch`
  ) |>
  mutate(
    EmployeeID = as.numeric(EmployeeID),
    EmailAddress = na_if(EmailAddress, "NULL"),
    IDIR = na_if(IDIR, "NULL")
  ) |>
  mutate(across(DaysInOffice:Saturday, ~ na_if(., "NULL")))

Telework_set <- Employee_TeleworkBCGovPSA |>
  select(EmployeeID, EmailAddress, DepartmentID)

## Build EmployeeList ####
EmployeeList <- Headcount_set |>
  full_join(
    Telework_set,
    by = join_by(EmployeeID, EmailAddress, DepartmentID)
  ) |>
  group_by(EmployeeID) |>
  summarise(
    EmailAddress = max(EmailAddress),
    IDIR = max(IDIR),
    BusinessUnit = max(BusinessUnit),
    DepartmentID = max(DepartmentID)
  ) |>
  ungroup() |>
  full_join(GAL_set, join_by(EmployeeID, EmailAddress)) |>
  group_by(EmployeeID) |>
  summarise(
    EmailAddress = max(EmailAddress),
    IDIR = max(IDIR),
    BusinessUnit = max(BusinessUnit),
    DepartmentID = max(DepartmentID),
    FirstName = max(FirstName),
    LastName = max(LastName)
  ) |>
  ungroup() |>
  full_join(
    Establishment_set,
    by = join_by(EmployeeID, DepartmentID, BusinessUnit)
  ) |>
  group_by(EmployeeID) |>
  summarise(
    EmailAddress = max(EmailAddress),
    IDIR = max(IDIR),
    BusinessUnit = max(BusinessUnit),
    DepartmentID = max(DepartmentID),
    FirstName = max(FirstName),
    LastName = max(LastName)
  ) |>
  ungroup() |>
  left_join(
    Employee_TeleworkBCGovPSA,
    by = join_by(
      EmployeeID,
      FirstName,
      LastName,
      EmailAddress,
      IDIR,
      DepartmentID
    )
  ) |>
  group_by(EmployeeID) |>
  fill(DaysInOffice:Saturday, .direction = "downup") |>
  distinct() |>
  left_join(
    Employee_EstablishmentReportBCGov,
    by = join_by(EmployeeID, BusinessUnit, DepartmentID)
  ) |>
  distinct() |>
  left_join(
    Employee_GAL_BCGov,
    by = join_by(EmployeeID, EmailAddress, FirstName, LastName)
  ) |>
  select(-c(Division, JobTitle, Job_Code, Name, Title)) |>
  distinct()

Address <- EmployeeList |>
  select(EmployeeID, City, City.x, City.y, Street.x, Street.y) |>
  pivot_longer(
    cols = contains("City"),
    names_to = "Column_Set",
    values_to = "City_Set"
  ) |>
  group_by(EmployeeID, City_Set) |>
  mutate(rowcount = n()) |>
  group_by(EmployeeID) |>
  slice(which.max(rowcount)) |>
  select(-c(rowcount, Column_Set)) |>
  mutate(Street = case_when(is.na(Street.y) ~ Street.x, .default = Street.y)) |>
  select(-c(Street.x, Street.y))

EmployeeList2 <- EmployeeList |>
  select(-c(City, City.x, City.y, Street.x, Street.y)) |>
  left_join(Address, by = join_by(EmployeeID)) |>
  mutate(
    Ministry = case_when(is.na(Company) ~ Company.y, .default = Company),
    Department = case_when(
      is.na(Department.y) ~ Department.x,
      .default = Department.y
    )
  ) |>
  select(
    -c(
      Company.x,
      Company,
      Company.y,
      Division.x,
      Division.y,
      Branch,
      Department.x,
      Department.y
    )
  )

## Geocode Address ####
Address <- EmployeeList2 |>
  select(EmployeeID, InputStreet = Street) |>
  filter(InputStreet != "0 – Confidential Location") |> # Can't reveal their secrets
  mutate(InputStreet = gsub("–", "-", InputStreet)) |> # weird character replacement
  mutate(InputStreet = gsub(",", "", InputStreet)) |>

  # Geocode Address
  #|>
  mutate(InputStreet = Street, InputCity = City) |>
  filter(InputStreet != "0 – Confidential Location") |> # Can't reveal their secrets
  mutate(InputStreet = gsub("–", "-", InputStreet)) |> # weird character replacement
  mutate(InputStreet = gsub(",", "", InputStreet)) |>
  # test_semicolon <- Employee_GAL_BCGov |>
  #   filter(grepl(";", InputStreet)) |>
  #   mutate(InputStreet = gsub(";.*", "", InputStreet))
  # mutate(InputStreet = gsub(";.*", "", InputStreet)) |>
  mutate(InputStreet = gsub(";", " ", InputStreet)) |>
  # test_flr <- Employee_GAL_BCGov |>
  #   filter(grepl("[Ff][l][r] [0-9]+ - ", InputStreet))
  mutate(InputStreet = gsub("[Ff][l][r] +[0-9]+ - ", "", InputStreet)) |>
  # test_floor <- Employee_GAL_BCGov |>
  #   filter(grepl("[0-9]+[snrt][tdh] [Ff][l][o][o][r] - ", InputStreet))
  mutate(
    InputStreet = gsub(
      "[0-9]+[snrt][tdh] [Ff][l][o][o][r][ ]?[-]?[ ]?",
      "",
      InputStreet
    )
  ) |>
  mutate(
    InputStreet = gsub("[0-9]+[snrt][tdh] [Ff][l][r]", "", InputStreet)
  ) |>
  mutate(
    InputStreet = gsub(
      "[:]?[-]?[ ]?[P][.]?[O][.]?[ ][B][Oo][Xx].*",
      "",
      InputStreet
    )
  ) |>
  mutate(InputStreet = gsub("[Ff][l][r][ ]+[0-9][NESW]", "", InputStreet)) |>
  mutate(InputStreet = gsub("#", "", InputStreet)) |>
  mutate(InputStreet = gsub("[.]", "", InputStreet)) |>
  mutate(
    InputStreet = gsub("Mailing Address[:]?", "", InputStreet)
  ) |>
  # Box 820 2250 West Trans Canada Highway
  # 150 - 4600 Jacombs Road Richmond BC Canada V6V 3B1
  mutate(
    InputStreet = na_if(InputStreet, ""),
    InputStreet = na_if(InputStreet, " "),
    InputCity = na_if(InputCity, ""),
    InputCity = na_if(InputCity, " ")
  )
# check for duplicates??

EmployeeAddressList <- Employee_GAL_BCGov |>
  select(Street, City, InputStreet, InputCity) |>
  filter(!is.na(InputStreet)) |>
  filter(!is.na(InputCity)) |>
  distinct() |>
  mutate(GeoStreet = "", GeoType = "", Score = "", Precision = "")
# test <- EmployeeAddressList |>
#   filter(is.na(InputStreet))
# test <- EmployeeAddressList |>
#   filter(grepl("P.O.", InputStreet))

EmployeeAddressList <- EmployeeAddressList |>
  filter(Score >= 90 | Precision == 100 & Score > 80)

write.csv(
  EmployeeAddressList,
  here("PBI/Data/GalEmployeeAddress.csv"),
  row.names = FALSE
)

EmployeeAddressList <- read.csv(here("PBI/Data/GalEmployeeAddress.csv"))

Employee_GAL_BCGov <- Employee_GAL_BCGov |>
  left_join(
    EmployeeAddressList,
    by = join_by(Street, City, InputStreet, InputCity)
  ) |>
  mutate(EmployeeID = as.numeric(EmployeeID))


Employee_Telework_Table <- arrow::read_parquet(here(
  "data/output/Py_Employee_Telework_Table.parquet"
))

match_list <- Employee_Telework_Table |>
  select(Start) |>
  mutate(
    GeoStreet = "",
    GeoType = "",
    Score = "",
    Precision = ""
  ) |>
  distinct()

# Use geocoder to improve addresses
API_KEY = read.csv("C:/Projects/credentials/bc_geocoder_api_key.csv") |> pull()
query_url = 'https://geocoder.api.gov.bc.ca/addresses.geojson?addressString='

for (ii in 1:nrow(match_list)) {
  location <- paste0(
    str_replace_all(match_list[ii, "Start"], " ", "%20")
  )
  req <- request(paste0(query_url, location)) |>
    req_headers(API_KEY = API_KEY) |>
    req_perform()
  resp <- req |> resp_body_json()
  match_list$GeoStreet[ii] <- resp$features[[
    1
  ]]$properties$fullAddress
  match_list$GeoType[ii] <- resp$features[[
    1
  ]]$properties$matchPrecision
  match_list$Precision[ii] <- resp$features[[
    1
  ]]$properties$precisionPoints
  match_list$Score[ii] <- resp$features[[1]]$properties$score
}

Employee_Telework_geo <- Employee_Telework_Table |>
  left_join(match_list, by = join_by(Start))

HQ_Telework_Table <- arrow::read_parquet(here(
  "data/output/Py_HQ_Telework_Table.parquet"
)) |>
  mutate(
    GeoStreet = "",
    GeoType = "",
    Score = "",
    Precision = "",
    .after = AddressEdit
  )

for (ii in 1:nrow(HQ_Telework_Table)) {
  location <- paste0(
    str_replace_all(HQ_Telework_Table[ii, "Start"], " ", "%20")
  )
  req <- request(paste0(query_url, location)) |>
    req_headers(API_KEY = API_KEY) |>
    req_perform()
  resp <- req |> resp_body_json()
  HQ_Telework_Table$GeoStreet[ii] <- resp$features[[1]]$properties$fullAddress
  HQ_Telework_Table$GeoType[ii] <- resp$features[[
    1
  ]]$properties$matchPrecision
  HQ_Telework_Table$Precision[ii] <- resp$features[[
    1
  ]]$properties$precisionPoints
  HQ_Telework_Table$Score[ii] <- resp$features[[1]]$properties$score
}
