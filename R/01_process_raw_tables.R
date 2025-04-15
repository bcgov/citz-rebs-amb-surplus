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
