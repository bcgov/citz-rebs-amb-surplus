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
    EmailAddress = `Email Address`,
    EmployeeID = `Employee ID`,
    JobTitle = `Job Title`,
    Company,
    Division,
    Department
  ) |>
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
# Use geocoder to improve addresses
API_KEY = read.csv("C:/Projects/credentials/bc_geocoder_api_key.csv") |> pull()
query_url = 'https://geocoder.api.gov.bc.ca/addresses.geojson?addressString='

for (ii in 1:nrow(EmployeeAddressList)) {
  location <- paste0(
    str_replace_all(EmployeeAddressList[ii, "InputStreet"], " ", "%20"),
    "%20",
    str_replace_all(EmployeeAddressList[ii, "InputCity"], " ", "%20")
  )
  req <- request(paste0(query_url, location)) |>
    req_headers(API_KEY = API_KEY) |>
    req_perform()
  resp <- req |> resp_body_json()
  EmployeeAddressList$GeoStreet[ii] <- resp$features[[1]]$properties$fullAddress
  EmployeeAddressList$GeoType[ii] <- resp$features[[
    1
  ]]$properties$matchPrecision
  EmployeeAddressList$Precision[ii] <- resp$features[[
    1
  ]]$properties$precisionPoints
  EmployeeAddressList$Score[ii] <- resp$features[[1]]$properties$score
}

EmployeeAddressList <- EmployeeAddressList |>
  filter(Score >= 90 | Precision == 100 & Score > 80)

write.csv(
  EmployeeAddressList,
  here("PBI/Data/GalEmployeeAddress.csv"),
  row.names = FALSE
)

Employee_EstablishmentReportBCGov = read_xlsx(
  here("data/Establishment Report - All BCPS - 2025-02-24.xlsx"),
  sheet = "Establishment 2025-02-24"
) |>
  filter(`Empl Status` != "NULL") |>
  select(
    EmployeeID = `Empl ID`,
    Company = Organization,
    Division = `Program Division`,
    Department = `Program Branch`,
    DepartmentID = `DeptID`
  )

Employee_HeadcountCITZ = read_xlsx(
  here("data/Headcount by Classification - All BCPS - 2025-02-24.xlsx"),
  sheet = "Headcount 2025-02-24"
)
Employee_TeleworkBCGovPSA = read_xlsx(
  here(
    "data/Staff Data For Telework Data Collection - All BCPS - 2025-03-06.xlsx"
  ),
  sheet = "Telework All Orgs 2025-03-06"
)
