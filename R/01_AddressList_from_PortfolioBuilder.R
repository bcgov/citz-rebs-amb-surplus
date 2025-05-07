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

# From Archibus, Report Builders, Portfolio Report Builders, grab the buildings tab
# Building Code & Name, Address 1, City/Site/Property Code, Tenure, Building Number, Facility type and Primary Use,
# Strategic classification, usable and rentable area (col says sqft, set profile to use sqM), lat and long.
buildings_report <- read_xlsx(
  here("data/portfolio-report-builders-2025-04-24-building-data.xlsx"),
  start_row = 3
) |> # Clean up column names to be more R friendly
  rename_with(~ gsub(" ", "", .), cols = everything()) |>
  rename(
    Address = Address1
  ) |>
  rename_with(~ gsub("[\\.]?[?]?[-]?", "", .), cols = everything()) |>
  rename_with(~ gsub("m²", "", .), .cols = everything()) |>
  # Remove leased buildings other than strategic leases
  filter(
    Tenure == "OWNED" |
      Address %in% c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Ave")
  ) |>
  # modify column values to not hurt my eyeballs.
  mutate(
    City = toTitleCase(tolower(CityCode)),
    Tenure = toTitleCase(tolower(Tenure)),
    .after = Address,
    .keep = "unused"
  ) |>
  mutate(
    # Update Name as it will be easier on geocoder and I think its known at this point
    City = case_when(
      City == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = City
    ),
    # clean up blank entries and replace with NA
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
    UsableArea,
    BuildingRentableArea = RentableArea,
    Latitude,
    Longitude
  ) |>
  # Remove properties, will collect with lands_report
  filter(
    !startsWith(Identifier, "N")
  ) |>
  # Add column to line up with lands_report
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
  rename_with(~ gsub("m²", "", .), .cols = everything()) |>
  # remove leased lands, strategic leases aren't in here
  filter(PropertyStatus == "OWNED") |>
  # make data prettier
  mutate(
    City = toTitleCase(tolower(CityName)),
    Tenure = toTitleCase(tolower(PropertyStatus)),
    .keep = "unused"
  ) |>
  # clean up name
  mutate(
    City = case_when(
      City == "Daajing Giids (Queen Charlotte)" ~ "Daajing Giids",
      .default = City
    )
  ) |>
  mutate(
    Identifier = PropertyCode,
    # replace blanks with NA values
    PrimaryUse = na_if(PrimaryUse, " "),
    StrategicClassification = na_if(StrategicClassification, " ")
  ) |>
  # Add columns to line up with building_report
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

# join both reports together
Facility_Table <- rbind(buildings_report, lands_report)

# setup for geocoding to get standardized addresses
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
R_Facility_Table <- Facility_Table |>
  left_join(AddressListFinal, by = join_by(Address == LinkAddress)) |>
  rename(LinkAddress = Address, Address = Address.y, City = City.y) |>
  relocate(Address, AddressEdit, LinkAddress, City, .after = SiteCode) |>
  select(-c(City.x)) |>
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
    BuildingRentableArea = sum(BuildingRentableArea, na.rm = TRUE),
    UsableArea = sum(UsableArea, na.rm = TRUE),
    Latitude = first(Latitude, na_rm = TRUE),
    Longitude = first(Longitude, na_rm = TRUE),
    GeoScore = first(GeoScore, na_rm = TRUE),
    GeoPrecision = first(GeoPrecision, na_rm = TRUE)
  )

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

zoning <- rbind(zoning_buildings, zoning_lands) |>
  select(Identifier, RegionalDistrict) |>
  distinct()

R_Facility_Table <- R_Facility_Table |>
  left_join(zoning, by = join_by(Identifier)) |>
  relocate(RegionalDistrict, .after = City)

write.csv(
  R_Facility_Table,
  here("PBI/data/R_Facility_Table.csv"),
  row.names = FALSE
)
