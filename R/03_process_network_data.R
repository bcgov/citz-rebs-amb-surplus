library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)
library(arrow)

connectivity_raw <- read_parquet(here("data/Network_Person_Counts.parquet"))

addresses <- connectivity_raw |>
  select(street_address) |>
  distinct() |>
  # mutate(split_city = str_extract(street_address, "([^,]*$)")) |>
  mutate(clean_address = gsub(",", " ", street_address)) |>
  mutate(geocoder_address = "", geocoder_score = "", geocoder_precision = "")


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

ii = 390

location <- "13%20TELEGRAPH%20CREEK%20RD%20DEASE%20LAKE"
location <- paste0(
  str_replace_all(addresses[ii, "clean_address"], " ", "%20")
)
req <- request(paste0(query_url, location)) |>
  req_headers(API_KEY = API_KEY) |>
  req_perform()

resp <- req |> resp_body_json()
