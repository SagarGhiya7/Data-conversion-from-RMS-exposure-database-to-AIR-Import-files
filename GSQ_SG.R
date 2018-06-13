# Loading packages
library(RODBC)
library(readr)
library(readxl)
library(stringr)

# Connecting to SQL Server
# Make sure SQL Server and R are on same desktop
# Enter your server name and database name  
dbhandle <- odbcDriverConnect('driver={SQL Server};server=frkcat-rms3sql1;database=AIG_2Q2016_GlobalProp_EDM17_ForMkts')

# Set Working Directory

setwd("C:/Users/u1112166/Desktop/GSQ")

# Check if all tables are properly imported
res <- sqlQuery(dbhandle, 'select * from information_schema.tables')

# Excel file to be made by user for importing. Contains information about portfolios
Portfolio <- data.frame(read_excel("AIG_2Q2016_GlobalProp_EDM17_ForMkts.xlsx"))


# SQL Script 03
# Below mentioned scripts are for reading scripts from SQL and running it. Relate to it whenever you come across such lines in program.

# Reading Script
my_query3 <- paste(readLines('03_Add_Tables_to_EDM_For_SiteQueries_AllCountries.sql', warn = FALSE),collapse = '\n')

# Running it.
results3 <- sqlQuery(dbhandle, my_query3)


# Check3 : Checking to see if results are correct
res3 <- sqlQuery(dbhandle, 'select * from CONSTMAP')

# for loop that runs through all the portfolios

for (i in Portfolio$Sr_No) {
  # Script 01
  my_query1 <- paste(readLines('01_Datalink_EDM_Tables.sql', warn = FALSE), collapse = '\n')
  
  # Making changes as per peril entered by user in excel file
  my_query1 <- str_replace(my_query1, '(?<=@PORTINFOID =\\s)\\w+', as.character(Portfolio[i, 3]))
  
  if (Portfolio[i, 5] == "EQ" | Portfolio[i,5] == "FF") {
    my_query1 <- str_replace(my_query1, '(?<=@POLICYTYPE =\\s)\\w+', "1")
    
    my_query1 <- str_replace(my_query1, '(?<=@PERIL =\\s)\\w+', "1")
  }
  if (Portfolio[i, 5] == "HU") {
    my_query1 <- str_replace(my_query1, '(?<=@POLICYTYPE =\\s)\\w+', "2")
    my_query1 <- str_replace(my_query1, '(?<=@PERIL =\\s)\\w+', "2")
  }
  if (Portfolio[i, 5] == "SCS") {
    my_query1 <- str_replace(my_query1, '(?<=@POLICYTYPE =\\s)\\w+', "3")
    my_query1 <- str_replace(my_query1, '(?<=@PERIL =\\s)\\w+', "3")
  }
  
  results1 <- sqlQuery(dbhandle, my_query1)
  
  # Check1
  #res1 <- sqlQuery(dbhandle, 'select * from  EDM_Site_Data ')
  
  #Q2
  
  my_query2 <- paste(readLines('02_Create_POLICY_FOR_AIRIMPORT_Table_Touchstone.sql', warn = FALSE),collapse = '\n')
  
  # Making changes as per peril entered by user in excel file
  if (Portfolio[i, 6] == "EQ" | Portfolio[i, 6] == "FF") {
    my_query2 <-  str_replace(my_query2, '(?<=@PERILCODE =\\s)\\W\\w+\\W', "'PEA'")
  }
  if (Portfolio[i, 6] == "Wind") {
    my_query2 <-  str_replace(my_query2, '(?<=@PERILCODE =\\s)\\W\\w+\\W', "'PWA'")
  }
  
  if (Portfolio[i, 6] == "SCS") {
    my_query2 <- str_replace(my_query2, '(?<=@PERILCODE =\\s)\\W\\w+\\W', "'CS'")
  }
  results2 <- sqlQuery(dbhandle, my_query2)
  
  # Check2
  #res2 <- sqlQuery(dbhandle, 'select * from POLICY_FOR_AIRIMPORT ')
  
  
  
  # Script 04
  
  my_query4 <-  paste(readLines('04_Check_for_Missing_Const_and_Occ_Mappings.sql', warn = FALSE), collapse = '\n')
  results4 <- data.frame(sqlQuery(dbhandle, my_query4))
  
  # Following is the code if mappings are not correct
  # If script 04 has rows, they will be printed and program will stop
  
  if (nrow(results4) > 0) {
    print(results4)
    next
  }
  
  
  # Script 06
  
  my_query6 <- paste(readLines('06_Create_SITE_FOR_AIRIMPORT_Table.sql', warn = FALSE), collapse = '\n')
  
  # Making changes as per peril entered by user in excel file
  
  if (Portfolio[i, 7] == "EQ" | Portfolio[i, 7] == "FF") {
    my_query6 <- str_replace(my_query6, '(?<=@PERIL =\\s)\\W\\w+\\W', "'EQ'")
  }
  
  if (Portfolio[i, 7] == "Wind") {
    my_query6 <- str_replace(my_query6, '(?<=@PERIL =\\s)\\W\\w+\\W', "'WD'")
  }
  
  if (Portfolio[i, 7] == "SCS") {
    my_query6 <- str_replace(my_query6, '(?<=@PERIL =\\s)\\W\\w+\\W', "'CS'")
  }
  
  if (Portfolio[i, 7] == "Flood") {
    my_query6 <- str_replace(my_query6, '(?<=@PERILCODE =\\s)\\W\\w+\\W', "'FL'")
  }
  
  results6 <- sqlQuery(dbhandle, my_query6)
  
  #Check 6
  
 # res6 <- sqlQuery(dbhandle, 'select * from SITE_FOR_AIRIMPORT')
  
  # Special conditions check : check whether to execute script 07, 08 , 09
  # Reading sql script for it
  
  query_extra <- paste(readLines('Special_Condition_Check.sql', warn = FALSE), collapse = '\n')
  
  # Making changes by changing portinfoid to which portfolio belongs
  
  query_extra <- str_replace(query_extra, '(?<=PA.PORTINFOID)\\s*\\W\\w+', paste("= ", as.character(Portfolio[i, 3])))
  # Running script
  results_extra <- data.frame(sqlQuery(dbhandle, query_extra))
  
  
  # Scripts 07, 08, 09
  # Executes below only if special conditions exist
  # Else directly to script 10
  
  if (nrow(results_extra) > 0) {
    my_query7 <- paste( readLines( '07_Create_Special_Conditions_Table_Delete_Conditions_NotNeeded.sql', warn = FALSE),collapse = '\n')
    # Making necessary changes
    my_query7 <- str_replace(my_query7, '(?<=@PORTINFOID =\\s)\\w+', as.character(Portfolio[i, 3]))
    
    if (Portfolio[i, 5] == "EQ" | Portfolio[i,5] == "FF") {
      my_query7 <- str_replace(my_query7, '(?<=@POLICYTYPE =\\s)\\w+', "1")
    }
    if (Portfolio[i, 5] == "HU") {
      my_query7 <- str_replace(my_query7, '(?<=@POLICYTYPE =\\s)\\w+', "2")
      
    }
    if (Portfolio[i, 5] == "SCS") {
      my_query7 <- str_replace(my_query7, '(?<=@POLICYTYPE =\\s)\\w+', "3")
      
    }
    results7 <- sqlQuery(dbhandle, my_query7)
    
    
    my_query8 <- paste( readLines( '08_Special_Conditions_Operations_for_Site_and_Pol_Files_TOUCHSTONE.sql', warn = FALSE), collapse = '\n')
    results8 <- sqlQuery(dbhandle, my_query8)
    
    my_query9 <- paste(readLines('09_Check_Sublimits_in_Pol_and_Site_Import_Files.sql', warn = FALSE),collapse = '\n')
    results9 <- sqlQuery(dbhandle, my_query9)
    
  }
  
  
  # Script 10
  
  my_query10 <- paste(readLines('10_Create_POLICY_FAC_FOR_AIRIMPORT_Table.sql', skipNul=TRUE, warn = FALSE),collapse = '\n')
  results10 <- sqlQuery(dbhandle, my_query10)
  
  #Check 10
  
  #res10 <- sqlQuery(dbhandle, 'select * from POLICY_FAC_FOR_AIRIMPORT')
  
  # Script 11
  
  my_query11 <- paste(readLines('11_Final_Cleanup_of_ImportFiles_TOUCHSTONE.sql'))
  
  # Making changes as per peril entered by user in excel file
  
  if (Portfolio[1, 7] == "EQ" | Portfolio[1, 7] == "FF") {
    my_query11 <- str_replace(my_query11, '(?<=@PERIL =\\s)\\W\\w+\\W', "'EQ'")
  }
  
  if (Portfolio[1, 7] == "Wind") {
    my_query11 <- str_replace(my_query11, '(?<=@PERIL =\\s)\\W\\w+\\W', "'WD'")
  }
  
  if (Portfolio[1, 7] == "SCS") {
    my_query11 <- str_replace(my_query11, '(?<=@PERIL =\\s)\\W\\w+\\W', "'CS'")
  }
  
  if (Portfolio[1, 7] == "Flood") {
    my_query11 <- str_replace(my_query11, '(?<=@PERILCODE =\\s)\\W\\w+\\W', "'FL'")
  }
  
  results11 <- sqlQuery(dbhandle, my_query11)
  
  #Q13
  
  my_query13 <- paste(readLines('13_PatchUp_Import_Files_Based_On_AIRImportErrors.sql', n = -1L, warn = FALSE))
  results13 <- sqlQuery(dbhandle, my_query13)
  
  # Exporting as CSV files for AIR Import
  write.csv( data.frame(sqlQuery( dbhandle, 'select * from POLICY_FOR_AIRIMPORT')), file = paste(Portfolio[i, 4], "_", Portfolio[i, 5], "_Account.csv"),row.names = FALSE)
  write.csv( data.frame(sqlQuery( dbhandle, 'select * from SITE_FOR_AIRIMPORT')), file = paste(Portfolio[i, 4], "_", Portfolio[i, 5], "_Location.csv"),row.names = FALSE)
  write.csv( data.frame(sqlQuery( dbhandle, 'select * from POLICY_FAC_FOR_AIRIMPORT')), file = paste(Portfolio[i, 4], "_", Portfolio[i, 5], "_Reinsurance.csv"), row.names = FALSE)
  
}

# End
