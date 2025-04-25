library(dplyr)
library(openxlsx2)
library(here)
library(httr2)
library(jsonlite)
library(stringr)
library(tidyr)
library(tools)

options(scipen = 999)
# Pre-process Core Data ####

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
            Tenure == "LEASED" &
                `Contract / Building / Property Number` == "N/A" ~
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
    mutate(PrimaryUse = toTitleCase(tolower(PrimaryUse))) |>
    distinct()

write.csv(
    cleanPWOR,
    here("data/output/CleanedProvinceWideOccupancy.csv"),
    row.names = FALSE
)

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
        -c(
            ParticipatesInPAM,
            UploadCharges,
            Description,
            AgreementDurationEndDate
        )
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

## Apple Report ####
# Apple <- read_xlsx(here("data/output/Apple_Table_Evan.xlsx"))

# cleanApple <- Apple |>
#     select(
#         Address = a_Address,
#         City = a_City,
#         BuildingNumber = a_Building_Number,
#         LeaseNumber = a_Lease_Number,
#         LandNumber = a_Land_Number,
#         VacantLand_HA = `a_Vacant Land (HA)`
#     )
# Build Tables ####

## AddressList ####
# Initial address list to cover the full scope of addresses prior to geocoding
AddressList <- cleanPWOR |>
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
        cleanCAR,
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
    full_join(cleanVFA, by = join_by(BuildingNumber == Assetnumber)) |>
    # Try and fill in missing address info with assetname as it often contains address details
    mutate(
        Address = case_when(is.na(Address) ~ Assetname, .default = Address)
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
    full_join(
        cleanFiscal,
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

test <- AddressList |>
    select(Address, City, InputAddress, InputCity) |>
    distinct()
test2 <- test |> select(InputAddress) |> distinct()
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
        BestAddress = case_when(
            score >= 90 ~ geo_Street,
            .default = InputAddress
        ),
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
R_PWOR_Table <- cleanPWOR |>
    # filter for owned or strategic leases
    filter(
        Tenure == "OWNED" |
            Address %in%
                c("4000 Seymour Pl", "4001 Seymour Pl", "1011 4th Avenue") # Strategic Leases
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
R_VFA_Table <- cleanVFA |>
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

write.csv(
    R_VFA_Table,
    here("PBI/Data/R_VFATable.csv"),
    row.names = FALSE
)

## CAR Table ####
R_CAR_Table <- cleanCAR |>
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

## Fiscal table ####
R_Fiscal_Table <- cleanFiscal |>
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
## Trim AddressList ####
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

## Project Costs Data ####
# Think this is the wrong section for this
R_Facility_Table <- read.csv(here("PBI/data/R_AddressTable.csv"))

ProjectCostsCBRE <- read_xlsx(
    here("data/ProjectReportApril152025.xlsx"),
    sheet = "CBRE"
) |>
    mutate(Source = "CBRE", .before = everything())

ProjectCostsWDS <- read_xlsx(
    here("data/ProjectReportApril152025.xlsx"),
    sheet = "WDS"
) |>
    mutate(Source = "WBS", .before = everything())

ProjectCosts <- rbind(ProjectCostsCBRE, ProjectCostsWDS) |>
    rename_with(~ gsub(" ", "", .), everything()) |>
    rename_with(~ gsub("/[\\n]?", "", perl = TRUE, .), everything()) |>
    rename_with(~ gsub("[/()#%]", "", .), everything())

# test <- ProjectCosts |> filter(is.na(BuildingID))
# test <- ProjectCosts |> filter(BuildingID == "9999100")
# test <- ProjectCostsSummary |>
# filter(
#     is.na(BuildingID) | BuildingID %in% c("9999100", "9999200", "9999300")
# ) |>
#     summarise(N_Projects = sum(N_Projects), Total_Paid = sum(Total_Paid))

# test$Total_Paid[1]

ProjectCostsSummary <- ProjectCosts |>
    group_by(BuildingID) |>
    summarise(
        Region = first(Region),
        N_Projects = n(),
        Total_Paid = sum(Paid, na.rm = TRUE)
    ) |>
    filter(
        !is.na(BuildingID) & !BuildingID %in% c("9999100", "9999200", "9999300")
    )

# This could be joined to the facility table at this point, or a building data table.

# Employee Data ####

## Temporary Geocode with finished tables ####
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
        BestAddress = case_when(
            is.na(BestAddress) ~ Match,
            .default = BestAddress
        )
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
