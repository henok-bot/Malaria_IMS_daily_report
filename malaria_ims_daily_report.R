##############################################################################-
############## Malaria daily data filling template #################


#Necessary packages - load the Epi packages script
pacman::p_load(tidyverse, readxl, import, here, janitor, rio, magrittr, esquisse,
               stringdist, stringr, openxlsx, lubridate)

### Setting up Filling template ###################

#Importing master data base of your daily report(you can modify the location based on your file)
master_data <- import( "C:/Users/Henok/Desktop/Malaria IMS daily report script for regions/Daily report/Daily_update_DB_R.xlsx")
glimpse(master_data) #Just to take a look

# creating latest day vector for later, assuming you will fill yesterdays report today
latest_date <- Sys.Date() - days(2)
#if formatting needed 
# latest_date <- format(latest_date,"%d/%m/%y")

# Add new rows for today by copying yesterday’s rows and update the date
new_data <- master_data %>% filter(`Date (MM/DD/YYYY)` == max(`Date (MM/DD/YYYY)`)) %>% 
  mutate(`Date (MM/DD/YYYY)` = latest_date) %>% 
  mutate(across(-c(`Date (MM/DD/YYYY)`, ADMIN3_EN, Region, Zone, Cluster, Woreda,
                   Rworeda, Kebele), ~NA)) %>% #make every indicator blank
  rename(region = Region) # to make it similar with each day report file column 

glimpse(new_data) # just to take a look


#### Read through, clean and combine report #############

#make sure your columns have the same column name as in the example and they are in one folder)

report_folder_path <- "C:/Users/Henok/Documents/Malaria IMS daily report/Daily report/October 8"
file_list <- list.files(report_folder_path, pattern = "*xlsx", full.names = TRUE)

# Loop through the file list and print column names for each file
for (file in file_list) {
  cat("\nColumn names for file:", file, "\n")
  
  # Read the first sheet of the file
  region_data <- read.xlsx(file, sheet = 1)
  
  # Print the column names
  print(colnames(region_data))
} # To help you look at the column names 

##extract data from each file
extract_region_data <- function(file){
  region_data <- read.xlsx(file,sheet = 1)
 ## cleaning the region data and making sure columns match
 region_data <- region_data %>% select(region = 'Region',
                                       woreda = 'Woreda',
                                       total_test  = 'Total.test(Fever.examined)',
                                       outpatient_case = 'Outpatient.cases',
                                       inpatient_case = 'Inpatient.cases',
                                       death = 'Malaria.Deathes',
                                       pf = 'Confirmed.Malaria.PF',
                                       pv = 'Confirmed.Malaria.PV',
                                       mixed = 'Confrimed.mixed',
                                       expected_facilites = 'Expected.reporting.facilites(number)',
                                       reported_facilities = 'Faciliites.which.reported(number)')
 return(region_data)
 } # This function stands to clean your column names and select the relevant indicators


#extract the files as data frame from the excels 
region_reports <- lapply(file_list,extract_region_data) #This gives you  list of the data frames
combined_report <- bind_rows(region_reports) %>% filter(!is.na(woreda)) %>% #change to one data frame and remove blank woreda rows
  rename(Rworeda = woreda ) #change the name of woreda to 

### Preliminary check on woreda name similarity ###########
# To do a preliminary check for matching variables 

#merge the data template(new_data) with combined_report of the day to see similarity of woredas(Rworeda)
merged_data <- new_data %>% left_join(combined_report,by = "Rworeda")
comparison_table <- merged_data %>% 
  group_by(region.x) %>% summarise(
    total_new_woredas = n(),
    matched_woredas = sum(!is.na(region.y)),
    unmatched_woredas = sum(is.na(region.y))
  ) # Take a look at the number of unmatched woredas:likely due to spelling variation or not reported 

print(comparison_table)
# Identify the mismatches to take measure 

mismatched_woredas <- merged_data %>% filter(is.na(region.y)) %>% select (Rworeda, region.x) 
 # filter(region.x %in% c("Oromia", "South West Ethiopia","Southern Ethiopi", "Central Ethiopia")) # not important when working with all region data


# This tries to create a data cleaning tunnel by checking woredas between new_data and combined_report
#The cleaning tunnel uses optimal string alignment distance (powerful approach to support your data cleaning )
#and add it also adds  a column close match(suggested woredas likely similar from the combined reports)
fuzzy_matches <- mismatched_woredas %>% mutate(
  close_match = sapply(Rworeda, function(Rworeda) {
    # Get the closest match in combined_data based on string distance
    combined_report$Rworeda[which.min(stringdist::stringdist(Rworeda, combined_report$Rworeda))]
  })
) %>% mutate(close_match = as.character(close_match)) # changing the list to vector

fuzzy_matches <- fuzzy_matches %>% filter(!Rworeda %in% c("zana", "soro"))
#If you accept all the fuzzy matching, update the new data with the suggested woreda names
new_data_corrected <- new_data %>% left_join(fuzzy_matches, by = "Rworeda") %>% 
  mutate(Rworeda = ifelse(!is.na(close_match),close_match, Rworeda)) %>%  
  select(-close_match, region.x) # remove close match and region.x(from the fuzzy match)


#Test again with the corrected woredas to see the matching scale

merged_data <- new_data_corrected %>% left_join(combined_report,by = "Rworeda")
comparison_table <- merged_data %>% 
  group_by(region.x) %>% summarise(
    total_new_woredas = n(),
    matched_woredas = sum(!is.na(region.y)),
    unmatched_woredas = sum(is.na(region.y))
  )
print(comparison_table)




# changing the column names to make them similar with the combined report for coalescing
colnames(new_data_corrected)
new_data_corrected <-  new_data_corrected %>% 
  rename(total_test  = `Total test`,
         outpatient_case = `Confirmed Malaria Case(Outpatient)`,
         inpatient_case = `Confirmed Malaria Case(Inpatient)`,
         death = `Confirmed Malaria death`,
         pf = PF,
         pv = PV,
         mixed = Mixed,
         expected_facilites = `Expected facilites`,
         reported_facilities = `Reporting facilites`)

#IF all is cool, you are done with the basic cleaning here

### Filling the data to the template #############
# Now fill the corrected data template by merging with the combined data

filled_data <- new_data_corrected %>% left_join(combined_report, by = "Rworeda")

colnames(master_data)


# Changing column names and class them to make them ready for merging with master data set
filled_data <- filled_data %>%  mutate(
  Region = as.character(region.x),
  'Total test' = as.numeric(coalesce(total_test.y,total_test.x)),
  'Confirmed Malaria Case(Inpatient)' = as.numeric(coalesce(inpatient_case.y,inpatient_case.x)),
 'Confirmed Malaria Case(Outpatient)' = as.numeric(coalesce(outpatient_case.y,outpatient_case.x)),
 'Confirmed Malaria death' = as.numeric(coalesce(death.y,death.x)),
 'Total Malaria case' = as.numeric(`Total Malaria case`),
 PF = as.numeric(coalesce(pf.y, pf.x)),
 PV = as.numeric(coalesce(pv.y,pv.x)),
 Mixed = as.numeric(coalesce(mixed.y, mixed.x)),
 'Expected facilites' = as.numeric(coalesce(expected_facilites.y,expected_facilites.x)),
 'Reporting facilites' = as.numeric(coalesce(reported_facilities.y, reported_facilities.x))
) %>% select(-ends_with(".x"), -ends_with(".y"))

# To add the total malaria case addition column
filled_data <- filled_data %>% rowwise() %>% mutate( 
  'Total Malaria case' = sum(`Confirmed Malaria Case(Inpatient)`,`Confirmed Malaria Case(Outpatient)`,na.rm = TRUE)) %>% 
  ungroup() 

filled_data <- filled_data %>% rowwise() %>% mutate(
  'Woreda reported' = if_else(!is.na(`Confirmed Malaria Case(Outpatient)`) & 
                                `Confirmed Malaria Case(Outpatient)` != "", "Yes", "No")) %>% 
  ungroup()

#Make the logic for the woreda reported
###--to be done


colnames(master_data)
glimpse(master_data)

# Doing some forgotten cleaning on the master data set imported
master_data <- master_data %>% 
  mutate(`Date (MM/DD/YYYY)` = as.Date(`Date (MM/DD/YYYY)`),
    Kebele = as.numeric(Kebele),
         `Total test` = as.numeric(`Total test`),
         `Confirmed Malaria Case(Inpatient)` = as.numeric(`Confirmed Malaria Case(Inpatient)`),
         `Confirmed Malaria Case(Outpatient)`= as.numeric(`Confirmed Malaria Case(Outpatient)`),
         `Confirmed Malaria death` = as.numeric(`Confirmed Malaria death`),
         `Total Malaria case` = as.numeric(`Total Malaria case`),
         PF = as.numeric(PF),
         PV = as.numeric(PV), 
         Mixed = as.numeric(Mixed),
         `Expected facilites`= as.numeric(`Expected facilites`),
         `Reporting facilites`= as.numeric(`Reporting facilites`)
)

### Merge and Export the final master data set###############
# Merge with the master data set and export it as xlsx 
updated_master_data <- bind_rows(master_data,filled_data)
#updated_master_data <- updated_master_data %>% 
 # mutate(`Date (MM/DD/YYYY)` = mdy(`Date (MM/DD/YYYY)`)) #To help excel in detecting the date


write.xlsx(updated_master_data, 
           file = "C:/Users/Henok/Desktop/Malaria IMS daily report script for regions/Daily report/Daily_update_DB_R.xlsx",
           sheetName = "National_report")


head(filled_data, 50)
