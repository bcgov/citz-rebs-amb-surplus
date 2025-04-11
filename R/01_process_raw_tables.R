library(dplyr)
library(openxlsx2)
library(here)

# Province Wide Occupancy ####
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
  )

# Customer Agreement Report ####
CAR <- read_xlsx(
  here(
    "data/csr0001-2025-03-11113053.016.xlsx"
  )
)

cleanCAR <- CAR |>
  rename(
    ParticipatesInPAM = `Participates In PAM`,
    UploadCharges = `Upload Charges`,
    AgreementNumber = `Agreement #`,
    AgreementStatus = `Agr Status`,
    FiscalYear = `Fiscal Year`,
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
  )
