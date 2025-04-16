library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)

options(scipen = 999)

# Goal of this is to take the cleaned scripts and build a well linked address table that links to every row
# in each of the detail tables (note I haven't fully achieved this yet but really close.)

pworTable <- read.csv(here("data/output/cleanedProvinceWideOccupancy.csv"))
carTable <- read.csv(here("data/output/cleanedCustomerAgreementReport.csv"))
vfaTable <- read.csv(here("data/output/cleanedVFAAssetList.csv"))
fiscalTable <- read.csv(here("data/output/cleanedFiscalReport.csv"))

# Build Tables ####

## AddressList ####
# Initial address list to cover the full scope of addresses prior to geocoding
AddressList <- pworTable |>
  select(
    Address,
    City,
    BuildingNumber = IdentifierPWOR,
    Contract = ContractPWOR,
    Tenure,
    AgreementNumber = ContractCode,
    AgreementType
  ) |>
  full_join(
    carTable,
    by = join_by(
      BuildingNumber == IdentifierCAR,
      Address == Address,
      AgreementNumber == AgreementNumber,
      AgreementType == AgreementType,
      City == City
    )
  ) |>
  # Remove non prop agreements
  # which list no city etc.
  filter(!is.na(Address)) |>
  select(
    Address,
    City,
    BuildingNumber,
    Contract,
    Tenure,
    AgreementNumber,
    AgreementType,
    Lease
  ) |>
  full_join(vfaTable, by = join_by(BuildingNumber == Assetnumber)) |>
  # Try and fill in missing address info with assetname as it often contains address details
  mutate(Address = case_when(is.na(Address) ~ Assetname, .default = Address)) |>
  select(
    Address,
    City,
    BuildingNumber,
    Contract,
    Tenure,
    AgreementNumber,
    AgreementType,
    Lease,
    Assetname,
    Campusname
  ) |>
  full_join(
    fiscalTable,
    by = join_by(BuildingNumber == ContractName, City == City)
  ) |>
  select(
    Address,
    City,
    BuildingNumber,
    Contract,
    Tenure,
    AgreementNumber,
    AgreementType,
    Lease,
    Assetname,
    Campusname
  ) |>
  filter(!is.na(Address)) |>
  # test <- AddressList |>
  #   filter(is.na(Address)) |>
  #   left_join(fiscalTable, by = join_by(BuildingNumber == ContractName)) |>
  #   filter(grepl("^B", BuildingNumber))
  # write.csv(test, here("data/output/missing_fiscal_buildings.csv"), row.names = FALSE)
  select(
    Address,
    City,
    BuildingNumber,
    Contract,
    Tenure,
    AgreementNumber,
    AgreementType,
    Lease,
    Assetname,
    Campusname
  ) |>
  select(Address, City, BuildingNumber) |>
  mutate(
    geo_name = "",
    address_type = "",
    score = "",
    precision = "",
    lat = "",
    lon = "",
    .after = City
  ) |>
  # should maybe do this before geocoding in future
  mutate(
    InputAddress = gsub("_", " ", Address),
    InputCity = toTitleCase(tolower(City)),
    .after = City
  ) |>
  mutate(
    InputCity = case_when(
      InputCity == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = InputCity
    )
  ) |>
  relocate(BuildingNumber, .before = everything()) |>
  distinct()

# Use geocoder to improve addresses
API_KEY = read.csv("C:/Projects/credentials/bc_geocoder_api_key.csv") |> pull()
query_url = 'https://geocoder.api.gov.bc.ca/addresses.geojson?addressString='

for (ii in 1:nrow(AddressList)) {
  location <- paste0(
    str_replace_all(AddressList[ii, "InputAddress"], " ", "%20"),
    "%20",
    str_replace_all(AddressList[ii, "InputCity"], " ", "%20")
  )
  req <- request(paste0(query_url, location)) |>
    req_headers(API_KEY = API_KEY) |>
    req_perform()
  resp <- req |> resp_body_json()
  AddressList$lon[ii] <- resp$features[[1]]$geometry$coordinates[[1]]
  AddressList$lat[ii] <- resp$features[[1]]$geometry$coordinates[[2]]
  AddressList$geo_name[ii] <- resp$features[[1]]$properties$fullAddress
  AddressList$address_type[ii] <- resp$features[[
    1
  ]]$properties$matchPrecision
  AddressList$precision[ii] <- resp$features[[
    1
  ]]$properties$precisionPoints
  AddressList$score[ii] <- resp$features[[1]]$properties$score
}

# touch up results to get best address and city results
AddressTable <- AddressList |>
  separate_wider_delim(
    geo_name,
    delim = ",",
    names = c("geo_Street", "geo_City", "Province"),
    too_few = "align_start"
  ) |>
  mutate(
    geo_Street = trimws(geo_Street),
    geo_City = trimws(geo_City),
    Province = trimws(Province)
  ) |>
  mutate(
    geo_Street = gsub("--", "-", geo_Street),
  ) |>
  mutate(
    BestAddress = case_when(score >= 90 ~ geo_Street, .default = InputAddress),
    BestCity = case_when(
      score >= 90 ~ geo_City,
      .default = InputCity
    ),
    .after = BuildingNumber
  ) |>
  # test <- AddressTable |>
  #   filter(score < 90)
  select(
    BuildingNumber,
    BestAddress,
    BestCity,
    LinkAddress = Address,
    LinkCity = City,
    InputAddress,
    InputCity,
    Score = score
  ) |>
  # Remove some duplicate addresses that make it through
  # 205 Ind Rd is a CAR duplicate of 205 Industrial Rd
  # 550 2nd Ave NE seems to be the wrong address for 101 - 6th Ave NE
  # Granted all are leases so don't really matter for this.
  filter(!LinkAddress %in% c("205 Ind Rd", "550 2nd Ave NE"))

## R_PWOR_Table ####
# building and land identifiers
R_PWOR_Table <- pworTable |>
  # filter for owned or strategic leases
  filter(
    Tenure == "OWNED" |
      Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Avenue") # Strategic Leases
  ) |>
  distinct() |>
  left_join(AddressTable, by = join_by(IdentifierPWOR == BuildingNumber)) |>
  select(
    BuildingNumber = IdentifierPWOR,
    Address = BestAddress,
    City = BestCity,
    LinkAddress,
    LinkCity,
    Tenure,
    PrimaryUse,
    UOM,
    TotalRentableArea = TotalRentableAreaPWOR,
    InventoryAllocatedRentableArea = InventoryAllocatedRentableAreaPWOR,
    CustomerAllocationPercentage = CustomerAllocationPercentagePWOR,
    ContractCode,
    AgreementType,
    CustomerCategory,
    Division,
    Department,
    InputAddress,
    InputCity,
    Score
  )

# test <- R_PWOR_Table |>
#   filter(LinkAddress != Address)
write.csv(
  R_PWOR_Table,
  here("PBI/Data/R_ProvWideOccupancyTable.csv"),
  row.names = FALSE
)

## VFA Table ####
R_VFA_Table <- vfaTable |>
  left_join(AddressTable, by = join_by(Assetnumber == BuildingNumber)) |>
  select(
    BuildingNumber = Assetnumber,
    Address = BestAddress,
    City = BestCity,
    LinkAddress,
    LinkCity,
    Regioneid,
    Regionname,
    Campuseid,
    Campusname,
    Asseteid,
    Assetname,
    Assetsize,
    Replacementvalue,
    Use,
    YearConstructed = Year.Constructed,
    Age,
    Fci,
    FciCost = Fcicost,
    Currency,
    AssetUnit = Assetunit,
    AssetType = Assettype,
    Ri,
    RiCost = Ricost,
    AssetCostUnit = Asset_costunit,
    InputAddress,
    InputCity,
    Score
  ) |>
  distinct()

# test <- R_VFA_Table |>
#   group_by(BuildingNumber) |>
#   mutate(rowcount = n()) |>
#   filter(rowcount >= 2)

write.csv(
  R_VFA_Table,
  here("PBI/Data/R_VFATable.csv"),
  row.names = FALSE
)

## CAR Table ####
R_CAR_Table <- carTable |>
  left_join(
    AddressTable,
    by = join_by(
      IdentifierCAR == BuildingNumber
    )
  ) |>
  # 4000 & 4001 Seymour and 1011 4th Ave aren't filtered or don't exist in table
  filter(!grepl("^L", Lease)) |>
  rename(
    BuildingNumber = IdentifierCAR,
    FiscalYear = CAR_FiscalYear
  ) |>
  relocate(
    BuildingNumber,
    BestAddress,
    BestCity,
    LinkAddress,
    LinkCity,
    .before = everything()
  ) |>
  select(-c(Address, City)) |>
  distinct() |>
  filter(!is.na(BestAddress)) |>
  filter(!AgreementStatus == "Draft") |>
  rename(Address = BestAddress, City = BestCity)

write.csv(
  R_CAR_Table,
  here("PBI/Data/R_CustomerAgreementReport.csv"),
  row.names = FALSE
)

# Fiscal table ####
R_Fiscal_Table <- fiscalTable |>
  left_join(AddressTable, by = join_by(ContractName == BuildingNumber)) |>
  filter(!is.na(BestAddress)) |>
  select(-c(City)) |>
  rename(
    BuildingNumber = ContractName,
    Address = BestAddress,
    City = BestCity
  ) |>
  relocate(Address, City, LinkAddress, LinkCity, .after = BuildingNumber)

write.csv(
  R_Fiscal_Table,
  here("PBI/Data/R_FiscalTable.csv"),
  row.names = FALSE
)

# Use the above created tables to create a trimmed address list after geocoding
R_PWOR_Table <- read.csv(here("PBI/Data/R_ProvWideOccupancyTable.csv"))
R_VFA_Table <- read.csv(here("PBI/Data/R_VFATable.csv"))
R_CAR_Table <- read.csv(here("PBI/Data/R_CustomerAgreementReport.csv"))
R_Fiscal_Table <- read.csv(here("PBI/Data/R_FiscalTable.csv"))

# Trim AddressList ####
AddressTableTrim <- R_PWOR_Table |>
  select(
    BuildingNumber,
    Address,
    City,
    LinkAddress,
    LinkCity,
    InputAddress,
    InputCity,
    Score
  ) |>
  full_join(
    R_Fiscal_Table,
    by = join_by(
      BuildingNumber,
      Address,
      City,
      LinkAddress,
      LinkCity,
      InputAddress,
      InputCity,
      Score
    )
  ) |>
  full_join(
    R_CAR_Table,
    by = join_by(
      BuildingNumber,
      Address,
      City,
      LinkAddress,
      LinkCity,
      InputAddress,
      InputCity,
      Score
    )
  ) |>
  full_join(
    R_VFA_Table,
    by = join_by(
      BuildingNumber,
      Address,
      City,
      LinkAddress,
      LinkCity,
      InputAddress,
      InputCity,
      Score
    )
  ) |>
  select(
    BuildingNumber,
    Address,
    City,
    LinkAddress,
    LinkCity,
    InputAddress,
    InputCity,
    Score
  ) |>
  distinct()

write.csv(
  AddressTableTrim,
  here("PBI/Data/R_AddressTable.csv"),
  row.names = FALSE
)
