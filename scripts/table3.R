rm(list = ls())

library(data.table)
library(tidyverse)
library(readxl)
library(meta)

anc <- read_excel('/Users/germanja/Documents/glp1-weight-loss-jg/data/GLP1_ancestry_effects.xlsx', sheet = 2)

# # # GLP1
# AFR
glp1_AFR <- anc %>% filter(Analysis == "GLP1-RA", Variable == "AFR")

meta_glp1_AFR <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "GLP1-RA", Variable == "AFR"),
  studlab = Study)

meta_glp1_AFR$TE.common
meta_glp1_AFR$seTE.common
meta_glp1_AFR$pval.common

# AMR
meta_glp1_AMR <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "GLP1-RA", Variable == "AMR"),
  studlab = Study)

meta_glp1_AMR$TE.common
meta_glp1_AMR$seTE.common
meta_glp1_AMR$pval.common

# EAS
glp1_EAS <- anc %>% filter(Analysis == "GLP1-RA", Variable == "EAS")

meta_glp1_EAS <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "GLP1-RA", Variable == "EAS"),
  studlab = Study)

meta_glp1_EAS$TE.common
meta_glp1_EAS$seTE.common
meta_glp1_EAS$pval.common

# SAS
meta_glp1_SAS <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "GLP1-RA", Variable == "SAS"),
  studlab = Study)

meta_glp1_SAS$TE.common
meta_glp1_SAS$seTE.common
meta_glp1_SAS$pval.common


# # # BS
# AFR
bs_AFR <- anc %>% filter(Analysis == "Bariatric surgery", Variable == "AFR")

meta_bs_AFR <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "Bariatric surgery", Variable == "AFR"),
  studlab = Study)

meta_bs_AFR$TE.common
meta_bs_AFR$seTE.common
meta_bs_AFR$pval.common

# AMR
meta_bs_AMR <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "Bariatric surgery", Variable == "AMR"),
  studlab = Study)

meta_bs_AMR$TE.common
meta_bs_AMR$seTE.common
meta_bs_AMR$pval.common

# EAS
bs_EAS <- anc %>% filter(Analysis == "Bariatric surgery", Variable == "EAS")

meta_bs_EAS <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "Bariatric surgery", Variable == "EAS"),
  studlab = Study)

meta_bs_EAS$TE.common
meta_bs_EAS$seTE.common
meta_bs_EAS$pval.common

# SAS
meta_bs_SAS <- metagen(
  TE = Beta, 
  seTE = SE, 
  data = anc  %>% filter(Analysis == "Bariatric surgery", Variable == "SAS"),
  studlab = Study)

meta_bs_SAS$TE.common
meta_bs_SAS$seTE.common
meta_bs_SAS$pval.common
