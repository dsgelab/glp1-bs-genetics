rm(list = ls())

library(data.table)
library(tidyverse)
library(readxl)
library(meta)

# Please refer to the google docs "GLP1_notes - 18.04.2024" for manual preprocessing steps on the excel file

# Specify the path to your Excel file
excel_file <- "/Users/cordioli/Projects/GLP1_bmi/results/GLP1_bodyweight_descriptive_analysis_results_20240418.xlsx"

# List all sheet names in the Excel file
sheet_names <- excel_sheets(excel_file)

# Initialize an empty list to store each sheet's data frame
sheets_list <- list()

# Keep only cohort with results (1-7) and exclude HUS (1)
sheet_names <- sheet_names[2:7]

# Loop through each sheet name
for (i in seq_along(sheet_names)) {
  if (sheet_names[i] == "BioMe") {
    # pivot_longer for BioMe data for ancestries
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 2)
    
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
             bs = as.numeric(bs))
    
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
             Ancestry = ifelse(sheet_names[i] %in% c("FinnGen", "Estonian Biobank", "UKBB"), "EUR", "ALL")) %>% 
      select(Study, Descriptor, Ancestry, glp1 = "Value...2", bs = "Value...3") %>% 
      mutate(glp1 = as.numeric(glp1),
             bs = as.numeric(bs))
  }
  # Add the sheet data to the list
  sheets_list[[i]] <- sheet_data
}

# Bind all sheets together into one data frame
desc_stat <- bind_rows(sheets_list) %>% 
  mutate(Study = ifelse(Study == "FinnGen", "HUS Biobank", Study),
         Ancestry = factor(Ancestry, levels = c("AFR", "ALL", "AMR", "EAS", "EUR", "SAS")))

cbPalette <- c("#999999", "#E69F00", "#56B4E9", "#009E73", "#0072B2", "#D55E00", "#CC79A7")

# Plot sample size and break down by ancestries
desc_stat_N <- desc_stat %>%
  filter(Descriptor == "N individuals included in the analysis")

total_n_glp1 <- sum(desc_stat_N$glp1, na.rm = T)
total_n_bs <- sum(desc_stat_N$bs, na.rm = T)

ancestry_n_glp1 <- desc_stat_N %>% 
  group_by(Ancestry) %>% 
  summarise(n = sum(glp1, na.rm = T)) %>% 
  mutate(label_glp1 = paste0(Ancestry, " - ", n)) %>% 
  select(-n)

ancestry_n_bs <- desc_stat_N %>% 
  group_by(Ancestry) %>% 
  summarise(n = sum(bs, na.rm = T))  %>% 
  mutate(label_bs = paste0(Ancestry, " - ", n)) %>% 
  select(-n)

desc_stat_N <- desc_stat_N %>% 
  left_join(ancestry_n_glp1) %>% 
  left_join(ancestry_n_bs)

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_N_by_ancestry.png", width = 6, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_N %>% filter(!is.na(glp1)), aes(x = glp1, y = Study, fill = label_glp1)) + 
  geom_bar(stat = "identity") +
  labs(x = "Total N", y = "Study", fill = "Ancestry") +
  ggtitle(paste0("Total sample size: ", total_n_glp1)) +
  scale_fill_manual(values = cbPalette) +
  theme_minimal()
dev.off()

png("/Users/cordioli/Projects/GLP1_bmi/plots/bs_N_by_ancestry.png", width = 6, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_N %>% filter(!is.na(bs)), aes(x = bs, y = Study, fill = label_bs)) + 
  geom_bar(stat = "identity") +
  labs(x = "Total N", y = "Study", fill = "Ancestry") +
  ggtitle(paste0("Total sample size: ", total_n_bs)) +
  scale_fill_manual(values = cbPalette) +
  theme_minimal()
dev.off()

# Plot sample size and break down by T2D status
desc_stat_t2d <- desc_stat %>%
  filter(startsWith(Descriptor, "T2D prevalence")) %>% 
  mutate(glp1 = ifelse(glp1 > 1, glp1/100, glp1),
         bs = ifelse(bs > 1, bs/100, bs))

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_T2D_by_ancestry.png", width = 6, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_t2d %>% filter(!is.na(glp1)), aes(x = glp1, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  labs(x = "T2D prevalence", y = "Study", fill = "Ancestry") +
  ggtitle("GLP1") +
  scale_fill_manual(values = cbPalette) +
  scale_x_continuous(labels = scales::percent) +
  theme_minimal()
dev.off()

png("/Users/cordioli/Projects/GLP1_bmi/plots/bs_T2D_by_ancestry.png", width = 6, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_t2d %>% filter(!is.na(bs)), aes(x = bs, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  labs(x = "T2D prevalence", y = "Study", fill = "Ancestry") +
  ggtitle("Bariatric surgery") +
  scale_fill_manual(values = c("#999999", "#E69F00", "#56B4E9", "#0072B2")) +
  scale_x_continuous(labels = scales::percent) +
  theme_minimal()
dev.off()

# Age
desc_stat_age_glp1 <- desc_stat %>%
  filter(startsWith(Descriptor, "Mean age") | startsWith(Descriptor, "SD age")) %>%
  mutate(Descriptor = ifelse(startsWith(Descriptor, "Mean age"), "AGE", "SD")) %>% 
  select(Study, Descriptor, Ancestry, glp1) %>% 
  pivot_wider(
    names_from = Descriptor,
    values_from = glp1)

desc_stat_age_bs <- desc_stat %>%
  filter(startsWith(Descriptor, "Mean age") | startsWith(Descriptor, "SD age")) %>%
  mutate(Descriptor = ifelse(startsWith(Descriptor, "Mean age"), "AGE", "SD")) %>% 
  select(Study, Descriptor, Ancestry, bs) %>% 
  pivot_wider(
    names_from = Descriptor,
    values_from = bs)

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_age_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_age_glp1 %>% filter(!is.na(AGE)), aes(x = AGE, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  geom_errorbarh(aes(xmin = AGE - SD, xmax = AGE + SD,),
                 position = position_dodge2(preserve = "single")) +
  labs(x = "Mean age at initiation (SD)", y = "Study", fill = "Ancestry") +
  ggtitle("GLP1") +
  scale_fill_manual(values = cbPalette) +
  coord_cartesian(xlim = c(35,80)) +
  theme_minimal()
dev.off()


png("/Users/cordioli/Projects/GLP1_bmi/plots/bs_age_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_age_bs %>% filter(!is.na(AGE)), aes(x = AGE, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  geom_errorbarh(aes(xmin = AGE - SD, xmax = AGE + SD,),
                 position = position_dodge2(preserve = "single")) +
  labs(x = "Mean age at surgery (SD)", y = "Study", fill = "Ancestry") +
  ggtitle("GLP1") +
  scale_fill_manual(values = c("#999999", "#E69F00", "#56B4E9", "#0072B2")) +
  coord_cartesian(xlim = c(35,80)) +
  theme_minimal()
dev.off()


# Sex
# Plot sample size and break down by % females
desc_stat_sex <- desc_stat %>%
  filter(startsWith(Descriptor, "% women") | startsWith(Descriptor, "% females")) %>% 
  mutate(glp1 = ifelse(glp1 > 1, glp1/100, glp1),
         bs = ifelse(bs > 1, bs/100, bs))

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_sex_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_sex %>% filter(!is.na(glp1)), aes(x = glp1, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  labs(x = "% females", y = "Study", fill = "Ancestry") +
  ggtitle("GLP1") +
  scale_fill_manual(values = cbPalette) +
  scale_x_continuous(labels = scales::percent) +
  theme_minimal()
dev.off()

png("/Users/cordioli/Projects/GLP1_bmi/plots/bs_sex_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_sex %>% filter(!is.na(bs)), aes(x = bs, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  labs(x = "% females", y = "Study", fill = "Ancestry") +
  ggtitle("Bariatric surgery") +
  scale_fill_manual(values = c("#999999", "#E69F00", "#56B4E9", "#0072B2")) +
  scale_x_continuous(labels = scales::percent) +
  theme_minimal()
dev.off()

# Type of GLP1 (semaglutide?)
desc_stat_type_glp1 <- desc_stat %>%
  filter(startsWith(Descriptor, "% Medication"))

meds_map <- data.frame(Descriptor = sort(unique(desc_stat_type_glp1$Descriptor)),
                       GLP1_name = c("exenatide",
                                     "liraglutide",
                                     "lixisenatide",
                                     "albiglutide",
                                     "dulaglutide",
                                     "semaglutide",
                                     "beinaglutide"))

desc_stat_type_glp1 <- desc_stat %>%
  filter(startsWith(Descriptor, "% Medication")) %>%
  mutate(glp1 = ifelse(!Study %in% c("MGBB", "AoU"), glp1/100, glp1)) %>% 
  group_by(Study, Descriptor) %>% 
  summarise(mean_prop = mean(glp1, na.rm = T)) %>% 
  mutate(mean_prop = ifelse(mean_prop == "NaN", 0, mean_prop)) %>% 
  left_join(meds_map) %>% 
  filter(!Study %in% c("Geisinger", "Qatar BB", "MVP")) %>% 
  select(-Descriptor)

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_type_glp1.png", width = 8, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_type_glp1, aes(x = GLP1_name, y = mean_prop)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single"), fill = "steelblue4") +
  labs(x = "", y = "% GLP1 type") +
  scale_y_continuous(labels = scales::percent) +
  facet_grid(.~Study) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
dev.off()

# ggplot(desc_stat_type_glp1, aes(x = mean_prop, y = Study, fill = GLP1_name)) + 
#   geom_bar(stat = "identity") +
#   labs(x = "% GLP1 type", y = "") +
#   ggtitle("GLP1") +
#   scale_fill_manual(values = cbPalette) +
#   scale_x_continuous(labels = scales::percent) +
#   theme_minimal()

# Baseline mean weight
desc_stat_w0_glp1 <- desc_stat %>%
  filter(startsWith(Descriptor, "Body weight mean at initiation") | startsWith(Descriptor, "Body weight SD at initiation")) %>%
  mutate(Descriptor = ifelse(startsWith(Descriptor, "Body weight mean"), "MEAN_W0", "SD")) %>% 
  select(Study, Descriptor, Ancestry, glp1) %>% 
  pivot_wider(
    names_from = Descriptor,
    values_from = glp1)

desc_stat_w0_bs <- desc_stat %>%
  filter(startsWith(Descriptor, "Body weight mean at initiation") | startsWith(Descriptor, "Body weight SD at initiation")) %>%
  mutate(Descriptor = ifelse(startsWith(Descriptor, "Body weight mean"), "MEAN_W0", "SD")) %>% 
  select(Study, Descriptor, Ancestry, bs) %>% 
  pivot_wider(
    names_from = Descriptor,
    values_from = bs)

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_w0_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_w0_glp1 %>% filter(!is.na(MEAN_W0)), aes(x = MEAN_W0, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  geom_errorbarh(aes(xmin = MEAN_W0 - SD, xmax = MEAN_W0 + SD,),
                 position = position_dodge2(preserve = "single")) +
  labs(x = "Mean weight at initiation (SD) in kg", y = "Study", fill = "Ancestry") +
  ggtitle("GLP1") +
  scale_fill_manual(values = cbPalette) +
  coord_cartesian(xlim = c(60,140)) +
  theme_minimal()
dev.off()

png("/Users/cordioli/Projects/GLP1_bmi/plots/bs_w0_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_w0_bs %>% filter(!is.na(MEAN_W0)), aes(x = MEAN_W0, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  geom_errorbarh(aes(xmin = MEAN_W0 - SD, xmax = MEAN_W0 + SD,),
                 position = position_dodge2(preserve = "single")) +
  labs(x = "Mean weight at initiation (SD) in kg", y = "Study", fill = "Ancestry") +
  ggtitle("Bariatric surgery") +
  scale_fill_manual(values = c("#999999", "#E69F00", "#56B4E9", "#0072B2")) +
  coord_cartesian(xlim = c(80,160)) +
  theme_minimal()
dev.off()

# Mean reduction in weight
desc_stat_delta_glp1 <- desc_stat %>%
  filter(startsWith(Descriptor, "Average percentage change in body weight")) %>% 
  select(Study, Descriptor, Ancestry, glp1) %>% 
  mutate(glp1 = glp1/100)

desc_stat_delta_bs <- desc_stat %>%
  filter(startsWith(Descriptor, "Average percentage change in body weight")) %>% 
  select(Study, Descriptor, Ancestry, bs) %>% 
  mutate(bs = bs/100)

png("/Users/cordioli/Projects/GLP1_bmi/plots/glp1_delta_bw_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_delta_glp1 %>% filter(!is.na(glp1)), aes(x = glp1, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  labs(x = "Mean % body weigth change", y = "Study", fill = "Ancestry") +
  ggtitle("GLP1") +
  scale_fill_manual(values = cbPalette) +
  scale_x_continuous(labels = scales::percent) +
  theme_minimal()
dev.off()

png("/Users/cordioli/Projects/GLP1_bmi/plots/bs_delta_bw_by_ancestry.png", width = 4, height = 3.5, units = "in", res = 300)
ggplot(desc_stat_delta_bs %>% filter(!is.na(bs)), aes(x = bs, y = Study, fill = Ancestry)) + 
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  labs(x = "Mean % body weigth change", y = "Study", fill = "Ancestry") +
  ggtitle("Bariatric surgery") +
  scale_fill_manual(values = c("#999999", "#E69F00", "#56B4E9", "#0072B2")) +
  scale_x_continuous(labels = scales::percent) +
  theme_minimal()
dev.off()