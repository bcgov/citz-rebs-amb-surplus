library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)

Employee_Telework_Table <- arrow::read_parquet(here(
  "data/output/Py_Employee_Telework_Table.parquet"
))
HQ_Telework_Table <- arrow::read_parquet(here(
  "data/output/Py_HQ_Telework_Table.parquet"
))

AddressList <- Employee_Telework_Table |>
  select(Start, Match) |>
  filter(Match != "No Match") |>
  distinct() |>
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
    str_replace_all(AddressList[ii, "Match"], " ", "%20")
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
    BestAddress = case_when(
      score >= 85 & precision >= 99 ~ geo_Street,
      .default = Match
    ),
    BestCity = case_when(
      score >= 85 & precision >= 99 ~ geo_City,
      .default = NA
    )
  )

Employee_Telework_Table <- Employee_Telework_Table |>
  left_join(AddressListFinal, by = join_by(Match, Start)) |>
  relocate(Start, Match, BestAddress, BestCity, .after = ImportDate) |>
  select(-c(Province, address_type)) |>
  mutate(
    BestAddress = case_when(is.na(BestAddress) ~ Match, .default = BestAddress)
  ) |>
  mutate(Address = BestAddress, AddressEdit = Match, .after = EmailAddress) |>
  select(-c(X, BestAddress, BestCity, geo_Street, geo_City, lat, lon))

AddressListFinal <- AddressListFinal |>
  select(-c(Start)) |>
  distinct()

HQ_Telework_Table <- HQ_Telework_Table |>
  left_join(AddressListFinal, by = join_by(AddressEdit == Match)) |>
  mutate(
    Address = case_when(
      score >= 85 & precision >= 99 ~ geo_Street,
      .default = AddressEdit
    ),
    .before = AddressEdit
  ) |>
  select(
    -c(
      geo_Street,
      geo_City,
      Province,
      address_type,
      lat,
      lon,
      BestAddress,
      BestCity
    )
  )

write.csv(
  Employee_Telework_Table,
  here("PBI/data/R_Employee_Telework_Table.csv"),
  row.names = FALSE
)
write.csv(
  HQ_Telework_Table,
  here("PBI/data/R_HQ_Telework_Table.csv"),
  row.names = FALSE
)

R_Bridge_Address_Table <- Employee_Telework_Table |>
  select(Address) |>
  distinct() |>
  full_join(
    HQ_Telework_Table |> select(Address) |> distinct(),
    by = join_by(Address)
  )

write.csv(
  R_Bridge_Address_Table,
  here("PBI/data/R_Bridge_Address_Table.csv"),
  row.names = FALSE
)
