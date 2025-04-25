library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)

options(scipen = 999)
# Portfolio Reports Processing ####

buildings_report <- read_xlsx(
  here("data/portfolio-report-builders-2025-04-24-building-data.xlsx"),
  start_row = 3
) |> # Clean up column names
  rename_with(~ gsub(" ", "", .), cols = everything()) |>
  rename(
    Address = Address1
  ) |>
  rename_with(~ gsub("[\\.]?[?]?[-]?", "", .), cols = everything()) |>
  rename_with(~ gsub("ft²", "", .), .cols = everything()) |>
  # Remove leased buildings other than strategic leases
  filter(
    Tenure == "OWNED" |
      Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  ) |>
  # modify column values
  mutate(
    City = toTitleCase(tolower(CityCode)),
    Region = toTitleCase(tolower(RegionCode)),
    Tenure = toTitleCase(tolower(Tenure)),
    .after = Address,
    .keep = "unused"
  ) |>
  mutate(
    City = case_when(
      City == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = City
    ),
    FacilityType = na_if(FacilityType, " "),
    StrategicClassification = na_if(StrategicClassification, " ")
  ) |>
  # Setup to join with lands
  select(
    Identifier = BuildingCode,
    Name = BuildingName,
    PropertyCode,
    SiteCode,
    Address,
    City,
    FacilityType,
    PrimaryUse,
    StrategicClassification,
    BuildingRentableArea = RentableArea,
    UsableArea,
    Latitude,
    Longitude
  ) |>
  mutate(LandArea = NA, .before = BuildingRentableArea)

lands_report <- read_xlsx(
  here("data/portfolio-report-builders-2025-04-24-land-data.xlsx"),
  start_row = 3
) |>
  # Clean up column names
  rename_with(~ gsub(" ", "", .), cols = everything()) |>
  rename(
    Address = Address1
  ) |>
  rename_with(~ gsub("[\\.]?[?]?[-]?", "", .), cols = everything()) |>
  rename_with(~ gsub("ft²", "", .), .cols = everything()) |>
  filter(PropertyStatus == "OWNED") |>
  mutate(
    City = toTitleCase(tolower(CityName)),
    Region = toTitleCase(tolower(RegionCode)),
    Tenure = toTitleCase(tolower(PropertyStatus)),
    .keep = "unused"
  ) |>
  mutate(
    City = case_when(
      City == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = City
    )
  ) |>
  mutate(
    Identifier = PropertyCode,
    PrimaryUse = na_if(PrimaryUse, " "),
    StrategicClassification = na_if(StrategicClassification, " ")
  ) |>
  mutate(FacilityType = "Land", UsableArea = NA) |>
  select(
    Identifier,
    Name = PropertyName,
    PropertyCode,
    SiteCode,
    Address,
    City,
    FacilityType,
    PrimaryUse,
    StrategicClassification,
    LandArea,
    BuildingRentableArea = AreaBldgRentable,
    UsableArea,
    Latitude,
    Longitude
  )

Facility_Table <- rbind(buildings_report, lands_report)

AddressList <- Facility_Table |>
  select(
    LinkAddress = Address,
    City
  ) |>
  #mutate(AddressInput = gsub("[-]", " ", AddressStart)) |>
  mutate(AddressEdit = gsub("[\\.]", " ", LinkAddress)) |>
  mutate(AddressEdit = gsub("[_]", " ", AddressEdit)) |>
  distinct() |>
  mutate(
    geo_name = "",
    address_type = "",
    score = "",
    precision = "",
    lat = "",
    lon = ""
  ) |>
  distinct()

# Use geocoder to improve addresses
API_KEY = read.csv("C:/Projects/credentials/bc_geocoder_api_key.csv") |> pull()
query_url = 'https://geocoder.api.gov.bc.ca/addresses.geojson?addressString='

for (ii in 1:nrow(AddressList)) {
  location <- paste0(
    str_replace_all(AddressList[ii, "AddressEdit"], " ", "%20"),
    "%20",
    str_replace_all(AddressList[ii, "City"], " ", "%20")
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

AddressListFinal <- AddressList |>
  separate_wider_delim(
    geo_name,
    delim = ",",
    names = c("geo_Street", "geo_City", "Province"),
    too_few = "align_start"
  ) |>
  mutate(
    geo_Street = trimws(geo_Street),
    geo_City = trimws(geo_City),
    Province = trimws(Province),
    score = as.numeric(score),
    precision = as.numeric(precision)
  ) |>
  mutate(
    geo_Street = gsub("--", "-", geo_Street),
  ) |>
  mutate(
    Address = case_when(
      score >= 85 & precision >= 99 ~ geo_Street,
      .default = AddressEdit
    ),
    City = case_when(
      score >= 85 & precision >= 99 ~ geo_City,
      .default = City
    )
  ) |>
  select(
    Address,
    City,
    AddressEdit,
    LinkAddress,
    GeoScore = score,
    GeoPrecision = precision,
    GeoLat = lat,
    GeoLon = lon
  )

# Clean up duplicate land values between buildings and properties tables
R_Facility_table <- Facility_Table |>
  left_join(AddressListFinal, by = join_by(Address == LinkAddress, City)) |>
  rename(LinkAddress = Address, Address = Address.y) |>
  relocate(Address, AddressEdit, LinkAddress, .after = SiteCode) |>
  group_by(Identifier) |>
  summarise(
    Name = first(Name),
    PropertyCode = first(PropertyCode, na_rm = TRUE),
    SiteCode = first(SiteCode, na_rm = TRUE),
    Address = first(Address, na_rm = TRUE),
    AddressEdit = first(AddressEdit, na_rm = TRUE),
    LinkAddress = first(LinkAddress, na_rm = TRUE),
    City = first(City, na_rm = TRUE),
    FacilityType = first(FacilityType, na_rm = TRUE),
    PrimaryUse = first(PrimaryUse, na_rm = TRUE),
    StrategicClassification = first(StrategicClassification, na_rm = TRUE),
    LandArea = first(LandArea, na_rm = TRUE),
    BuildingRentableArea = sum(BuildingRentableArea, na_rm = TRUE),
    UsableArea = sum(UsableArea, na_rm = TRUE),
    Latitude = first(Latitude, na_rm = TRUE),
    Longitude = first(Longitude, na_rm = TRUE),
    GeoScore = first(GeoScore, na_rm = TRUE),
    GeoPrecision = first(GeoPrecision, na_rm = TRUE)
  )

# check for vacant land

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

pwor <- read_xlsx(
  here(
    "data/ho-rplm-report-pmr3345-province-wide-occupancy-2025-04-03152308.604.xlsx"
  ),
  start_row = 3
) |>
  mutate(
    Identifier = `Building/Land #`,
    IdentifierType = case_when(
      startsWith(as.character(`Building/Land #`), "N") ~ "Land",
      startsWith(as.character(`Building/Land #`), "B") ~ "Building",
      .default = NA
    ),
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
  group_by(Identifier) |>
  mutate(rowcount = n()) |>
  ungroup() |>
  # fix for 4 cases where AgreementType is NA but its a single plot of land
  mutate(
    AgreementType = case_when(
      is.na(AgreementType) & startsWith(Identifier, "N") & rowcount == 1 ~
        "Land Only",
      .default = AgreementType
    )
  ) |>
  select(-c(rowcount)) |>
  # find vacant plots in that they only have Land Only Agreements
  group_by(Identifier) |>
  filter(if_all(AgreementType, ~ startsWith(AgreementType, "Land Only"))) |>
  filter(Tenure == "OWNED")

test <- pwor |>
  group_by(Identifier) |>
  filter(if_any(AgreementType, is.na)) |>
  filter(Tenure == "OWNED")
