rm(list = ls())

library(data.table)
library(tidyverse)
library(readxl)
library(meta)
library(cowplot)
library(ggtext)

# Please refer to the google docs "GLP1_notes - 18.04.2024" for manual preprocessing steps on the excel file

# Specify the path to your Excel file
excel_file <- "/Users/germanja/Documents/glp1-weight-loss-jg/data/GLP1_bodyweight_descriptive_analysis_results_20240418_jg.xlsx"

# List all sheet names in the Excel file
sheet_names <- excel_sheets(excel_file)

# Initialize an empty list to store each sheet's data frame
sheets_list <- list()

# Keep only cohort with results (1-8) # QATAR: 1-9 
sheet_names <- sheet_names[1:9] # QATAR: 9 

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
    
  } else if (sheet_names[i] == "MGBB") {
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
  } else if (sheet_names[i] %in% c("FinnGen-HUS", "Estonian Biobank", "UKBB", "BBSS")) { 
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 1) %>%
      mutate(Study = sheet_names[i],
             Ancestry = ifelse(sheet_names[i] %in% c("FinnGen-HUS", "Estonian Biobank", "UKBB", "BBSS"), "EUR", "ALL")) %>% 
             #Ancestry = ifelse(sheet_names[i] %in% c("QBB"), "MEA", "ALL")) %>% 
      select(Study, Descriptor, Ancestry, glp1 = "Value...2", bs = "Value...3") %>% 
      mutate(glp1 = as.numeric(glp1),
             bs = as.numeric(bs))
  } else {
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i], skip = 1) %>%
      mutate(Study = sheet_names[i],
             Ancestry = ifelse(sheet_names[i] %in% c("QBB"), "MEA", "ALL")) %>% 
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
         Ancestry = factor(Ancestry, levels = c("AFR", "ALL", "AMR", "EAS", "EUR", "MEA", "SAS"))) # QATAR: add "MEA" between "EUR" and "SAS"


# Figure 2
# average % weight change per each cohort/ancestry and average across studies

# Keep only the rows/descriptors we need
desc_stat_filt <- desc_stat %>% 
  filter(Descriptor %in% c("N individuals included in the analysis",
                           "Average percentage change in body weight between the two time points\r\n ∑( (mW1 - W0)/W0 * 100 ) / n",
                           "Average percentage change in body weight between the two time points\r\n∑( (mW1 - W0)/W0 * 100 ) / n",
                           "SD percentage change in body weight between the two time points\r\n∑( (mW1 - W0)/W0 * 100 ) / n")) %>% 
  mutate(Descriptor = case_when(startsWith(Descriptor, "N individuals") ~ "N",
                                startsWith(Descriptor, "Average percentage") ~ "mean_perc_change",
                                startsWith(Descriptor, "SD") ~ "SD_perc_change"))

# Mean reduction in weight
delta_glp1 <- desc_stat_filt %>%
  select(Study, Descriptor, Ancestry, glp1) %>% 
  pivot_wider(id_cols = c("Study", "Ancestry"),
              names_from = Descriptor,
              values_from = glp1) %>% 
  filter(!(Study %in% c("AoU", "MGBB", "UCLA-ATLAS") &  Ancestry == "ALL")) %>% 
  mutate(mean_perc_change = mean_perc_change/100,
         SD_perc_change = SD_perc_change/100,
         SE_perc_change = SD_perc_change/sqrt(N)) %>% 
  filter(Study != "BBSS")
  
delta_bs <- desc_stat_filt %>%
  select(Study, Descriptor, Ancestry, bs) %>% 
  pivot_wider(id_cols = c("Study", "Ancestry"),
              names_from = Descriptor,
              values_from = bs) %>% 
  filter(!(Study %in% c("AoU", "MGBB", "UCLA-ATLAS") &  Ancestry == "ALL"),
         !is.na(N)) %>% 
  mutate(mean_perc_change = mean_perc_change/100,
         SD_perc_change = SD_perc_change/100,
         SE_perc_change = SD_perc_change/sqrt(N))

# Calculate meta-analysis mean using meta pacckage
# GLP1
delta_glp1_meta <- metamean(data = delta_glp1, 
                            n = N,
                            mean = mean_perc_change,
                            sd = SD_perc_change)

delta_glp1_meta_mean <- delta_glp1_meta$TE.common
delta_glp1_meta_se <- delta_glp1_meta$seTE.common
delta_glp1_meta_lower <- delta_glp1_meta$lower.common
delta_glp1_meta_upper <- delta_glp1_meta$upper.common

delta_glp1_comb <- delta_glp1 %>% 
  mutate(lower = mean_perc_change - SE_perc_change,
         upper = mean_perc_change + SE_perc_change) %>% 
  bind_rows(data.frame(Study = "Combined",
                       Ancestry = "ALL",
                       N = sum(delta_glp1$N),
                       mean_perc_change = delta_glp1_meta$TE.common,
                       SD_perc_change = delta_glp1_meta$seTE.common*1.96,
                       lower = delta_glp1_meta$TE.common - delta_glp1_meta$seTE.common,
                       upper = delta_glp1_meta$TE.common + delta_glp1_meta$seTE.common)) %>% 
  mutate(Study = factor(Study, levels = rev(c("AoU", "BioMe", "ESTBB", "HUS", "MGBB", "UCLA-ATLAS", "UKBB", "QBB", "Combined")))) %>%  # QATAR: add QBB before combined
  mutate(type = "GLP1-RA") %>% 
  mutate(Ancestry = factor(ifelse(Ancestry == "ALL", "Combined", Ancestry), levels = c("AFR", "AMR", "EAS", "EUR", "MEA", "SAS", "Combined")),  # QATAR: add "MEA" between "EUR" and "SAS"
         label = paste0(Study, " (N: ", N, ")"))

# BS
delta_bs_meta <- metamean(data = delta_bs, 
                            n = N,
                            mean = mean_perc_change,
                            sd = SD_perc_change)

delta_bs_meta_mean <- delta_bs_meta$TE.common
delta_bs_meta_se <- delta_bs_meta$seTE.common
delta_bs_meta_lower <- delta_bs_meta$lower.common
delta_bs_meta_upper <- delta_bs_meta$upper.common

delta_bs_comb <- delta_bs %>% 
  mutate(lower = mean_perc_change - SE_perc_change,
         upper = mean_perc_change + SE_perc_change) %>% 
  bind_rows(data.frame(Study = "Combined",
                       Ancestry = "ALL",
                       N = sum(delta_bs$N),
                       mean_perc_change = delta_bs_meta$TE.common,
                       SD_perc_change = delta_bs_meta$seTE.common*1.96,
                       lower = delta_bs_meta$TE.common - delta_bs_meta$seTE.common,
                       upper = delta_bs_meta$TE.common + delta_bs_meta$seTE.common)) %>% 
  mutate(Study = factor(Study, levels = rev(c("AoU", "BBSS", "BioMe", "ESTBB", "HUS", "MGBB", "UCLA-ATLAS", "UKBB", "QBB", "Combined")))) %>%  # QATAR: add QBB before combined
  mutate(type = "Bariatric surgery") %>%
  mutate(Ancestry = factor(ifelse(Ancestry == "ALL", "Combined", Ancestry), levels = c("AFR", "AMR", "EAS", "EUR", "MEA", "SAS", "Combined")),  # QATAR: add "MEA" between "EUR" and "SAS"
         label = paste0(Study, " (N: ", N, ")"))


# # Combine
# delta_all <- delta_glp1_comb %>% mutate(type = "GLP1-RA") %>% 
#   bind_rows(delta_bs_comb %>% mutate(type = "Bariatric surgery")) %>% 
#   mutate(Ancestry = factor(ifelse(Ancestry == "ALL", "Combined", Ancestry), levels = c("AFR", "AMR", "EAS", "EUR", "SAS", "Combined")),
#          label = paste0(Study, " (N: ", N, ")"))
  

cbPalette <- c("#E69F00", "#56B4E9", "#009E73", "#0072B2", "#F0E442", "#CC79A7", "#999999")  # QATAR: add colorcode for MEA on right position "#D55E00", or "#F0E442",

p1 <- ggplot(delta_glp1_comb %>% filter(type == "GLP1-RA", !is.na(mean_perc_change)),
             aes(y = Study, x = mean_perc_change, fill = Ancestry)) +
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  geom_errorbar(aes(xmin = lower, xmax = upper), position = position_dodge2(preserve = "single")) +
  geom_vline(xintercept = 0, alpha = .8) +
  labs(x = "Mean percentage change in body weight", y = "Study", fill = "Ancestry") +
  scale_fill_manual(values = cbPalette,
                    labels = c("AFR", "AMR", "EAS", "EUR", "MEA", "SAS", "Combined")) +  # QATAR: add "MEA" between "EUR" and "SAS"
  scale_x_continuous(labels = scales::percent) +
  scale_y_discrete(labels = function(x) ifelse(x == "Combined", paste0("**", x, "**"), x)) +
  theme_minimal() +
  theme(axis.text.y = ggtext::element_markdown()) +
  ggtitle("GLP1-RA")

p1

# p1 +
#   scale_y_discrete(expand = expansion(mult = c(0, 0.75))) 

cbPaletteBS <- c("#E69F00", "#56B4E9", "#0072B2", "#F0E442", "#999999") # QATAR: add colorcode for MEA on right position "#D55E00", or "#F0E442",

p2 <- ggplot(delta_bs_comb %>% filter(type == "Bariatric surgery", !is.na(mean_perc_change)),
             aes(y = Study, x = mean_perc_change, fill = Ancestry)) +
  geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
  geom_errorbar(aes(xmin = lower, xmax = upper), position = position_dodge2(preserve = "single")) +
  geom_vline(xintercept = 0, alpha = .8) +
  labs(x = "Mean percentage change in body weight", y = "Study", fill = "Ancestry") +
  scale_fill_manual(values = cbPaletteBS,
                    labels = c("AFR", "AMR", "EUR", "MEA", "Combined")) + # QATAR: add "MEA" between "EUR" and "SAS"
  scale_x_continuous(labels = scales::percent) +
  scale_y_discrete(labels = function(x) ifelse(x == "Combined", paste0("**", x, "**"), x)) +
  theme_minimal() +
  theme(axis.text.y = ggtext::element_markdown()) +
  ggtitle("Bariatric surgery")
p2

png("/Users/germanja/Documents/glp1-weight-loss-jg/figures_tables/Figure2A_FOR_LEGENDwQ.png", width = 7, height = 7, units = "in", res = 300)
print(p1) + theme(legend.position = "bottom") + guides(fill = guide_legend(title = "Ancestry", nrow = 1, byrow = TRUE))
dev.off()

p1 <- p1 + theme(legend.position = "none")
p2 <- p2 + theme(legend.position = "none")

png("/Users/germanja/Documents/glp1-weight-loss-jg/figures_tables/Figure2SE_2wQ.png", width = 9, height = 4, units = "in", res = 300)
plot_grid(p1, p2)
dev.off()

# Combine legend and plot in biorender


# png("/Users/cordioli/Projects/GLP1_bmi/figures_tables/Figure2SE_facets.png", width = 7, height = 5, units = "in", res = 300)
# 
# ggplot(delta_all %>% filter(!is.na(mean_perc_change)), aes(y = Study, x = mean_perc_change, fill = Ancestry)) +
#   geom_bar(stat = "identity", position = position_dodge2(preserve = "single")) +
#   geom_errorbar(aes(xmin = lower, xmax = upper), position = position_dodge2(preserve = "single")) +
#   geom_vline(xintercept = 0, alpha = .8) +
#   labs(x = "Mean percentage change in body weight", y = "Study", fill = "Ancestry") +
#   scale_fill_manual(values = cbPalette) +
#   scale_x_continuous(labels = scales::percent) +
#   theme_minimal() +
#   facet_grid(.~factor(type, levels = c("GLP1-RA","Bariatric surgery")),
#              scales = "free_x")
# 
# dev.off()