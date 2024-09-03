rm(list = ls())

library(data.table)
library(tidyverse)
library(readxl)
library(meta)

# Please refer to the google docs "GLP1_notes - 18.04.2024" for manual preprocessing steps on the excel file

# Specify the path to your Excel file
excel_file <- "/Users/germanja/Documents/glp1-weight-loss-main/data/GLP1_bodyweight_descriptive_analysis_results_20240418.xlsx"

# List all sheet names in the Excel file
sheet_names <- excel_sheets(excel_file)

# Initialize an empty list to store each sheet's data frame
sheets_list <- list()

# Keep only cohort with results (1-7)
sheet_names <- sheet_names[1:7]

# Loop through each sheet name
for (i in seq_along(sheet_names)) {
  if (sheet_names[i] == "BioMe") {
    # pivot_longer for BioMe data for ancestries
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 1)
    
    colnames(sheet_data) <- c("Descriptor",
                              "glp1_AMR", "glp1_AFR", "glp1_EAS", "glp1_EUR", "glp1_SAS",
                              "bs_AMR", "bs_AFR", "bs_EAS", "bs_EUR", "bs_SAS")
    
  
    sheet_data <- sheet_data %>% 
      pivot_longer(
        cols = all_of(starts_with("glp1")),
        names_to = "Ancestry_glp",
        values_to = "glp1",
        names_prefix = "glp1_") %>% 
      pivot_longer(
        cols = all_of(starts_with("bs")),
        names_to = "Ancestry_bs",
        values_to = "bs",
        names_prefix = "bs_") %>% 
      filter(Ancestry_glp == Ancestry_bs) %>%
      mutate(Study = sheet_names[i]) %>% 
      select(Study, Descriptor, Ancestry = Ancestry_glp, glp1, bs) %>% 
      mutate(glp1 = as.numeric(glp1),
             bs = as.numeric(bs)) %>% 
      filter(!(is.na(glp1) & is.na(bs)))
    
    } else if (sheet_names[i] == "AoU") {
      # pivot_longer for BioMe data for ancestries
      sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 1)
      
      colnames(sheet_data) <- c("Descriptor",
                                "glp1_ALL", "bs_ALL",
                                "glp1_EUR", "bs_EUR",
                                "glp1_AFR", "bs_AFR")
      
      sheet_data <- sheet_data %>% 
        pivot_longer(
          cols = all_of(starts_with("glp1")),
          names_to = "Ancestry_glp",
          values_to = "glp1",
          names_prefix = "glp1_") %>% 
        pivot_longer(
          cols = all_of(starts_with("bs")),
          names_to = "Ancestry_bs",
          values_to = "bs",
          names_prefix = "bs_") %>% 
        filter(Ancestry_glp == Ancestry_bs) %>%
        mutate(Study = sheet_names[i]) %>% 
        select(Study, Descriptor, Ancestry = Ancestry_glp, glp1, bs) %>% 
        mutate(glp1 = as.numeric(glp1),
               bs = as.numeric(bs))
    
    } else if (sheet_names[i] == "UCLA-ATLAS") {
    # pivot_longer for ATLAS data for ancestries
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 2)
    
    # Remove column 7 (marked for bariatric surgery - empty)
    sheet_data <- sheet_data %>% select(-7)
    
    colnames(sheet_data) <- c("Descriptor",
                              "glp1_ALL", "glp1_AFR", "glp1_AMR", "glp1_EAS", "glp1_EUR")
    
    sheet_data <- sheet_data %>% 
      pivot_longer(
        cols = all_of(starts_with("glp1")),
        names_to = "Ancestry_glp",
        values_to = "glp1",
        names_prefix = "glp1_") %>% 
      mutate(Study = sheet_names[i],
             bs = NA) %>% 
      select(Study, Descriptor, Ancestry = Ancestry_glp, glp1, bs) %>% 
      mutate(glp1 = as.numeric(glp1),
             bs = as.numeric(bs))
    } else {
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 1) %>%
      mutate(Study = sheet_names[i],
             Ancestry = ifelse(sheet_names[i] %in% c("FinnGen-HUS", "Estonian Biobank", "UKBB", "MGBB"), "EUR", "ALL")) %>% 
      select(Study, Descriptor, Ancestry, glp1 = "Value...2", bs = "Value...3") %>% 
      mutate(glp1 = as.numeric(glp1),
             bs = as.numeric(bs))
  }
  # Add the sheet data to the list
  sheets_list[[i]] <- sheet_data
}

# Bind all sheets together into one data frame
desc_stat <- bind_rows(sheets_list) %>% 
  mutate(Study = ifelse(Study == "FinnGen-HUS", "HUS", Study),
         Study = ifelse(Study == "Estonian Biobank", "ESTBB", Study),
         Ancestry = factor(Ancestry, levels = c("AFR", "ALL", "AMR", "EAS", "EUR", "SAS")))


# Table 1
# Columns: cohort, ancestry, N, mean age (SD), % females, mean weight at baseline (SD), % T2D, most common med (%), % on semaglutide

# Keep only the rows/descriptors we need
desc_stat_filt <- desc_stat %>% 
  filter(Descriptor %in% c("N individuals included in the analysis",
                           "Mean age in years",
                           "Mean age in years (at initiation/surgery)",
                           "SD age in years",
                           "% females",
                           "% women",
                           "T2D prevalence at initiation (in %)",
                           "T2D prevalence at initiation/surgery (in %)*",
                           "% Medication 1 (ATC= A10BJ01 & A10BX04)",                                                              
                           "% Medication 2 (ATC= A10BJ02 & A10BX07)",                                                              
                           "% Medication 3 (ATC= A10BJ03 & A10BX10)",                                                              
                           "% Medication 4 (ATC= A10BJ04 & A10BX13)",                                                              
                           "% Medication 5 (ATC= A10BJ05 & A10BX14)",                                                              
                           "% Medication 6 (ATC= A10BJ06)",                                                                        
                           "% Medication 7 (ATC= A10BJ07)",
                           "Body weight mean at initiation or surgery (W0)",
                           "Body weight SD at initiation or surgery (W0)")) %>% 
  mutate(Descriptor = case_when(startsWith(Descriptor, "N individuals") ~ "Number of individuals",
                                startsWith(Descriptor, "Mean age") ~ "Mean age at baseline",
                                startsWith(Descriptor, "SD age") ~ "SD age",
                                startsWith(Descriptor, "% females") ~ "Proportion of females",
                                startsWith(Descriptor, "% women") ~ "Proportion of females",
                                startsWith(Descriptor, "T2D prevalence") ~ "T2D prevalence",
                                startsWith(Descriptor, "% Medication 1") ~ "Proportion of exenatide",
                                startsWith(Descriptor, "% Medication 2") ~ "Proportion of liraglutide",
                                startsWith(Descriptor, "% Medication 3") ~ "Proportion of lixisenatide",
                                startsWith(Descriptor, "% Medication 4") ~ "Proportion of albiglutide",
                                startsWith(Descriptor, "% Medication 5") ~ "Proportion of dulaglutide",
                                startsWith(Descriptor, "% Medication 6") ~ "Proportion of semaglutide",
                                startsWith(Descriptor, "% Medication 7") ~ "Proportion of beinaglutide",
                                startsWith(Descriptor, "Body weight mean") ~ "Mean body weight at baseline",
                                startsWith(Descriptor, "Body weight SD") ~ "SD body weight"))

desc_stat_filt_wide <- desc_stat_filt %>% 
  pivot_wider(id_cols = c("Study","Ancestry"),
              names_from = "Descriptor",
              values_from = c("glp1", "bs")) %>% 
  mutate(`glp1_Mean age at baseline (SD)` = paste0(round(`glp1_Mean age at baseline`, 2), " (", round(`glp1_SD age`, 2), ")"),
         `bs_Mean age at baseline (SD)` = paste0(round(`bs_Mean age at baseline`, 2), " (", round(`bs_SD age`, 2), ")"),
         `glp1_Mean body weight at baseline (SD)` = paste0(round(`glp1_Mean body weight at baseline`, 2), " (", round(`glp1_SD body weight`, 2), ")"),
         `bs_Mean body weight at baseline (SD)` = paste0(round(`bs_Mean body weight at baseline`, 2), " (", round(`bs_SD body weight`, 2), ")"))


# Extract name of most used glp1 med in each cohort
most_used_glp1 <- c()
for (i in 1:nrow(desc_stat_filt_wide)){
  glp1_cols <- desc_stat_filt_wide[i, 8:14]
  glp1_med <- names(glp1_cols)[which(glp1_cols == max(glp1_cols, na.rm = T))]
  glp1_med <- gsub("glp1_Proportion of ", "", glp1_med)
  most_used_glp1 <- c(most_used_glp1, glp1_med)
}

desc_stat_filt_wide_glp1 <- desc_stat_filt_wide %>% 
  mutate(`glp1_Most common GLP1 medication` = most_used_glp1) %>% 
  select(Study, Ancestry,
         `glp1_Number of individuals`,
         `glp1_Proportion of females`,
         `glp1_Mean body weight at baseline (SD)`,
         `glp1_Mean age at baseline (SD)`,
         `glp1_T2D prevalence`,
         `glp1_Most common GLP1 medication`,
         `glp1_Proportion of semaglutide`)
colnames(desc_stat_filt_wide_glp1) <- gsub("glp1_", "", colnames(desc_stat_filt_wide_glp1))

desc_stat_filt_wide_bs <- desc_stat_filt_wide %>% 
  select(Study, Ancestry,
         `bs_Number of individuals`,
         `bs_Proportion of females`,
         `bs_Mean body weight at baseline (SD)`,
         `bs_Mean age at baseline (SD)`,
         `bs_T2D prevalence`) %>% 
  filter(!is.na(`bs_Number of individuals`))
colnames(desc_stat_filt_wide_bs) <- gsub("bs_", "", colnames(desc_stat_filt_wide_bs))

#fwrite(desc_stat_filt_wide_glp1, 'Projects/GLP1_bmi/figures_tables/Table1_GLP1.tsv', sep = "\t")
#fwrite(desc_stat_filt_wide_bs, 'Projects/GLP1_bmi/figures_tables/Table1_BS.tsv', sep = "\t")

mean(desc_stat_filt_wide$`glp1_Mean age at baseline`, na.rm = T)
mean(desc_stat_filt_wide$`bs_Mean age at baseline`, na.rm = T)

weighted.mean(desc_stat_filt_wide$`glp1_Mean age at baseline`, desc_stat_filt_wide$`glp1_Number of individuals`, na.rm = T)
weighted.mean(desc_stat_filt_wide$`bs_Mean age at baseline`, desc_stat_filt_wide$`bs_Number of individuals`, na.rm = T)

weighted.mean(desc_stat_filt_wide$`glp1_Proportion of females`, desc_stat_filt_wide$`glp1_Number of individuals`, na.rm = T)
weighted.mean(desc_stat_filt_wide$`bs_Proportion of females`, desc_stat_filt_wide$`bs_Number of individuals`, na.rm = T)

mean(desc_stat_filt_wide$`glp1_Proportion of females`, na.rm = T)
