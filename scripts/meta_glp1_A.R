rm(list = ls())

library(data.table)
library(tidyverse)
library(readxl)
library(meta)
library(ggforestplot)

snps_dict <- fread("/Users/germanja/Documents/glp1-weight-loss-jg/data/snps_dict.csv")
files <- system('ls /Users/germanja/Documents/glp1-weight-loss-jg/data/GLP1_bw*.xlsx', intern = T)

res <- NULL

# Read in results for GLP1 - model A
for (f in files){
  study <- strsplit(basename(f), split = "_")[[1]][3]
  
  # Skip BBSS since we have no GLP1 analysis for that one
  if (study == "BBSS") {
    next
  }
  
  ancestry <- gsub(".xlsx", "", strsplit(basename(f), split = "_")[[1]][5])
    
  df <- read_excel(f, sheet = 1) 
  
  # Rename "Beta_ baseline_body_weight_in_kg" if present
  if ("Beta_ baseline_body_weight_in_kg" %in% colnames(df)) {
    df <- df %>%
      rename(Beta_baseline_body_weight_in_kg = "Beta_ baseline_body_weight_in_kg",
             SE_baseline_body_weight_in_kg = "SE_ baseline_body_weight_in_kg")
  }
  
  df <- df %>% 
    mutate(Study = study,
           Ancestry = ancestry,
           across(Beta_exposure:SE_age_at_initiation_in_years, as.numeric)) %>%
    select(Study,
           Ancestry,
           -Reference, 
           Exposure = `Analysis name`,
           Effect_allele = `Effect allele`,
           Beta_exposure:SE_age_at_initiation_in_years)
  
  # In UKBB results, beta is ~396 for rs2295006. Set all value to NA
  if(study == "UKBB") {
    df <- df %>% 
      mutate(across(starts_with(c("Beta_", "SE_")), ~if_else(Exposure %in% c("rs2295006", "rs146868158"), NA, .)))
  }
  
  # Separate SNPs rows, check effect allele and change direction if swapped
  df_snps <- df %>%
    filter(!grepl("PGS", Exposure)) %>%
    left_join(snps_dict, by = c("Exposure" = "SNP")) %>%
    mutate(Beta_exposure = if_else(Effect_allele != Alt & Effect_allele == Ref, -Beta_exposure, Beta_exposure),
           Effect_allele = if_else(Effect_allele != Alt & Effect_allele == Ref, Ref, Effect_allele)) %>%
    select(-Ref, -Alt)
  
  df_pgs <- df %>%
    filter(grepl("PGS", Exposure))
  
  # Combine SNP and PGS rows back together
  df <- bind_rows(df_pgs, df_snps)
  res <- bind_rows(res, df) %>% 
    select(-Effect_allele)
}

# Calculate P values
calculate_p_value <- function(beta, se) {
  ifelse(is.na(beta) | is.na(se), NA, 2*pnorm(-abs(beta/se)) )
}

res <- res %>% 
  mutate(P_exposure = calculate_p_value(Beta_exposure, SE_exposure),
         P_baseline_body_weight_in_kg = calculate_p_value(Beta_baseline_body_weight_in_kg, SE_baseline_body_weight_in_kg),
         P_sex = calculate_p_value(Beta_sex, SE_sex),
         P_age_at_initiation_in_years = calculate_p_value(Beta_age_at_initiation_in_years, SE_age_at_initiation_in_years))


# # # Meta-analysis

# Perform ancestry specific meta-analyses
meta_results_anc <- NULL

for (a in unique(res$Ancestry)) {
  for (e in unique(res$Exposure)) {
    
    sub_res <- res %>%
      filter(Exposure == e,
             Ancestry == a)
    
    # Print the filtered data frame and its dimensions
    print(paste("Ancestry:", a, "Exposure:", e, "Number of studies:", nrow(sub_res)))
    
    if (nrow(sub_res) == 0) {
      next # Skip if no data
    }
    
    for (v in c("exposure", "baseline_body_weight_in_kg", "sex", "age_at_initiation_in_years")){
      
      if (!all(c(paste0("Beta_", v), paste0("SE_", v)) %in% names(sub_res))) {
        print(paste("Missing columns for variable:", v))
        next # Skip if the required columns are missing
      }
      
      m <- with(sub_res, metagen(
        TE = get(paste0("Beta_", v)), 
        seTE = get(paste0("SE_", v)), 
        data = sub_res,
        studlab = Study))
      
      r <- data.frame(
        Study = paste0("Meta_",a),
        Ancestry = a,
        Exposure = e,
        Variable = v,
        Beta = m$TE.common,
        SE = m$seTE.common,
        P = m$pval.common
      )
      
      meta_results_anc <- bind_rows(meta_results_anc, r)
      
    }
  }
}

# Pivot meta results in wide format
meta_results_anc_wide <- meta_results_anc %>%
  pivot_wider(
    names_from = Variable,
    values_from = c(Beta, SE, P),
    names_glue = "{.value}_{Variable}")


# Perform multi-ancestry meta-analyses
meta_results <- NULL
meta_het <- NULL

for (e in unique(meta_results_anc_wide$Exposure)) {
  
  sub_meta <- meta_results_anc_wide %>%
    filter(Exposure == e)
  
  for (v in c("exposure", "baseline_body_weight_in_kg", "sex", "age_at_initiation_in_years")){
    
    m <- with(sub_meta, metagen(
      TE = get(paste0("Beta_", v)), 
      seTE = get(paste0("SE_", v)), 
      data = sub_meta,
      studlab = Study))
    
    r <- data.frame(
      Study = "Meta_analysis",
      Ancestry = "ALL",
      Exposure = e,
      Variable = v,
      Beta = m$TE.common,
      SE = m$seTE.common,
      P = m$pval.common
    )
    
    meta_results <- bind_rows(meta_results, r)
    
    if (v == "exposure"){
      h <- data.frame(
        Exposure = e,
        Q_het = m$Q,
        P_Q_het = m$pval.Q,
        I2_het = m$I2
      )
      meta_het <- bind_rows(meta_het, h)
    }
  }
}

# Pivot meta results in wide format
meta_results_wide <- meta_results %>%
  pivot_wider(
    names_from = Variable,
    values_from = c(Beta, SE, P),
    names_glue = "{.value}_{Variable}")

# Combine with individual study results
res_tot <- bind_rows(res, meta_results_anc_wide, meta_results_wide)

res_tot <- res_tot %>% 
  mutate(Study = factor(Study, levels = rev(unique(res_tot$Study)))) %>% 
  filter(!Exposure %in% c("WHRadjBMI PGS", "rs2295006", "rs201672448"))

fwrite(res_tot, "/Users/germanja/Documents/glp1-weight-loss-jg/results/wq/meta_GLP1_A_20240424.tsv", sep = "\t")
fwrite(meta_het, "/Users/germanja/Documents/glp1-weight-loss-jg/results/wq/meta_GLP1_A_20240424_ancestry_het.tsv", sep = "\t")
