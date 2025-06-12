library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)
library(arrow)

connectivity_raw <- read_parquet(here("data/Network_Person_Counts.parquet")) |>
  select(AddressLink, street_address, approx_users, date, dotw, weekday_yn) |>
  filter(date >= "2025-01-01") |>
  filter(weekday_yn != "Weekend") |>
  group_by(AddressLink, street_address, dotw) |>
  summarize(
    average_daily_users = mean(approx_users)
  ) |>
  pivot_wider(names_from = dotw, values_from = average_daily_users) |>
  relocate(
    Monday,
    Tuesday,
    Wednesday,
    Thursday,
    Friday,
    .after = street_address
  ) |>
  rowwise() |>
  mutate(WeeklyAverage = mean(c_across(c(Monday:Friday)), na.rm = TRUE))

write.csv(
  connectivity_raw,
  here("PBI/data/R_Connectivity_Table.csv"),
  row.names = FALSE
)


# list <- connectivity_raw |> select(AddressLink) |> arrange() |> distinct()

# Attempt Geocoding ####
addresses <- connectivity_raw |>
  select(street_address) |>
  distinct() |>
  # mutate(split_city = str_extract(street_address, "([^,]*$)")) |>
  mutate(clean_address = gsub(",", " ", street_address)) |>
  mutate(geocoder_address = "", geocoder_score = "", geocoder_precision = "") |>
  # add address cleanup steps
  mutate(test = gsub("#", "", street_address))

write_xlsx(addresses, here("data/address_first_pass.xlsx"))

# Use geocoder to improve addresses
API_KEY = read.csv("C:/Projects/credentials/bc_geocoder_api_key.csv") |> pull()
query_url = 'https://geocoder.api.gov.bc.ca/addresses.geojson?addressString='

for (ii in 1:nrow(addresses)) {
  location <- paste0(
    str_replace_all(addresses[ii, "clean_address"], " ", "%20")
  )
  req <- request(paste0(query_url, location)) |>
    req_headers(API_KEY = API_KEY) |>
    req_perform()

  resp <- req |> resp_body_json()

  addresses$geocoder_address[ii] <- resp$features[[1]]$properties$fullAddress
  addresses$geocoder_precision[ii] <- resp$features[[
    1
  ]]$properties$precisionPoints
  addresses$geocoder_score[ii] <- resp$features[[1]]$properties$score
}

AddressTable <- addresses |>
  separate_wider_delim(
    geocoder_address,
    delim = ",",
    names = c("Address", "City", "Province"),
    too_few = "align_start"
  ) |>
  mutate(
    Address = trimws(Address),
    City = trimws(City),
    Province = trimws(Province),
    geocoder_score = as.numeric(geocoder_score),
    geocoder_precision = as.numeric(geocoder_precision)
  ) #|>
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

ii = 390

location <- "13%20TELEGRAPH%20CREEK%20RD%20DEASE%20LAKE"
location <- paste0(
  str_replace_all(addresses[ii, "clean_address"], " ", "%20")
)
req <- request(paste0(query_url, location)) |>
  req_headers(API_KEY = API_KEY) |>
  req_perform()

resp <- req |> resp_body_json()
