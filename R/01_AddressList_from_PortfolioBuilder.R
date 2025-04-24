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
  # remove fully na or otherwise unwanted columns
  select(
    -c(
      Address2,
      LocationType,
      BuildingGraphic,
      RealPropertyUniqueIdentifier,
      BuildingNumber,
      ConditionIndex,
      RegionName,
      ProvincialRegion,
      BuildingUse
    )
  ) |>
  # remove junk
  filter(!is.na(Address) | Address != "Dummy Building") |>
  relocate(
    BuildingCode,
    Address,
    CityCode,
    CityName,
    RegionCode,
    PostalCode,
    .before = BuildingName
  ) |>
  # Remove leased buildings other than strategic leases
  filter(
    Tenure == "OWNED" |
      Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  ) |>
  # modify column values
  mutate(
    City = toTitleCase(tolower(CityCode)),
    Region = toTitleCase(tolower(RegionCode)),
    .after = Address,
  ) |>
  mutate(
    City = case_when(
      City == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = City
    )
  ) |>
  mutate(
    Tenure = toTitleCase(tolower(Tenure)),
    OccupancyStatus = toTitleCase(tolower(OccupancyStatus))
  ) |>
  # Remove leftover columns from cleaning or those with mainly missing data/unhelpful data
  select(
    -c(
      CityCode,
      CityName,
      RegionCode,
      LeaseStatus,
      DateBuilt,
      OccupancyStatus,
      HistoricBuilding,
      BuildingOccupancy,
      CostperArea,
      CosttoReplace,
      MaxBldgOccupancy,
      PostalCode,
      POBCStatus,
      ExtGrossArea,
      IntGrossArea,
      ValueMarket,
      DateMarketValueAssessed
    )
  ) |>
  relocate(
    StrategicClassification,
    FacilityType,
    RentableArea,
    UsableArea,
    .after = Tenure
  ) |>
  mutate(
    Identifier = BuildingCode,
    Name = BuildingName,
    .before = everything(),
    .keep = "unused"
  ) |>
  mutate(LandArea = NA)

lands_report <- read_xlsx(
  here("data/portfolio-report-builders-2025-04-24-land-data.xlsx"),
  start_row = 3
) |>
  filter(`Property Status` == "OWNED") |>
  # Clean up column names
  rename_with(~ gsub(" ", "", .), cols = everything()) |>
  rename(
    Address = Address1
  ) |>
  rename_with(~ gsub("[\\.]?[?]?[-]?", "", .), cols = everything()) |>
  rename_with(~ gsub("ft²", "", .), .cols = everything()) |>
  rename(
    Tenure = PropertyStatus,
    City = CityName,
    Region = RegionCode,
    RentableArea = AreaBldgRentable
  ) |>
  relocate(Address, City, Region, .after = PropertyCode) |>
  mutate(
    City = toTitleCase(tolower(City)),
    Region = toTitleCase(tolower(Region)),
    Tenure = toTitleCase(tolower(Tenure))
  ) |>
  mutate(
    City = case_when(
      City == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = City
    )
  ) |>
  mutate(Identifier = PropertyCode, Name = PropertyName) |>
  select(-c(POBCStatus, PrimaryUse, PropertyName, AreaLeaseNegotiated)) |>
  mutate(FacilityType = "Land", UsableArea = NA, PrimaryUse = NA)

Facility_Table <- rbind(buildings_report, lands_report) |>
  mutate(
    FacilityType = na_if(FacilityType, " "),
    StrategicClassification = na_if(StrategicClassification, " ")
  ) |>
  select(-c(Region)) |>
  relocate(PrimaryUse, LandArea, .after = FacilityType)

AddressList <- Facility_Table |>
  select(AddressStart = Address, City) |>
  distinct() |>
  mutate(AddressInput = gsub("[-]", " ", AddressStart)) |>
  mutate(AddressInput = gsub("[\\.]", " ", AddressInput)) |>
  mutate(AddressInput = gsub("[_]", " ", AddressInput)) |>
  mutate(
    geo_name = "",
    address_type = "",
    score = "",
    precision = "",
    lat = "",
    lon = ""
  )

# Use geocoder to improve addresses
API_KEY = read.csv("C:/Projects/credentials/bc_geocoder_api_key.csv") |> pull()
query_url = 'https://geocoder.api.gov.bc.ca/addresses.geojson?addressString='

for (ii in 1:nrow(AddressList)) {
  location <- paste0(
    str_replace_all(AddressList[ii, "AddressInput"], " ", "%20"),
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
      .default = AddressInput
    ),
    City = case_when(
      score >= 85 & precision >= 99 ~ geo_City,
      .default = City
    )
  ) |>
  mutate(LinkAddress = AddressStart) |>
  select(
    Address,
    City,
    LinkAddress,
    GeoScore = score,
    GeoPrecision = precision,
    GeoLat = lat,
    GeoLon = lon
  )

write.csv(
  AddressListFinal,
  here("Data/output/portfolio_address_list.csv"),
  row.names = FALSE
)

OG_Address <- read.csv(here("PBI/Data/R_AddressTable.csv")) |>
  select(Address) |>
  distinct()

# Produces a different amount of buildings from my merge efforts...
