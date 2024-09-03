rm(list = ls())

library(data.table)
library(tidyverse)
library(readxl)
library(meta)
library(cowplot)

glp1 <- fread('/Users/germanja/Documents/glp1-weight-loss-jg/results/wq/meta_GLP1_A_20240424.tsv')
bs <- fread('/Users/germanja/Documents/glp1-weight-loss-jg/results/wq/meta_BS_A_20240424.tsv')

glp1 <- glp1 %>%
  mutate(Analysis = "GLP1-RA")
bs <- bs %>%
  mutate(Analysis = "Bariatric surgery")

res <- bind_rows(glp1, bs) %>% 
  mutate(Analysis = factor(Analysis, levels = c("GLP1-RA", "Bariatric surgery")))

res_meta <- res %>% 
  filter(startsWith(Study, "Meta"))

res_meta <- res_meta %>%
  mutate(Meta = factor(ifelse(Ancestry == "ALL", "Multi-ancestry", Ancestry), levels = rev(c("AFR","AMR","EAS","EUR", "MEA", "SAS","Multi-ancestry"))))

cbPalette <- c("#E69F00", "#56B4E9", "#009E73", "#0072B2", "#F0E442", "#CC79A7", "black")

# Plot individual studies and meta-analysis beta for BMI and T2D PGS
res_meta_pgs <- res_meta %>%
  filter(Exposure %in% c("BMI PGS", "T2D PGS"))

p1 <- ggforestplot::forestplot(
  df = res_meta_pgs,
  name = Exposure,
  estimate = Beta_exposure,
  se = SE_exposure,
  pvalue = P_exposure,
  psignif = 0.004,
  colour = Meta,
  shape = Meta,
  xlab = "Percentage change in body weight per 1-SD change in PGS",
  title = "") +
  ggplot2::scale_shape_manual(
     name = "Ancestry",
     values = c(23L, 23L, 23L, 23L, 23L, 23L, 23L)) +
  ggplot2::scale_color_manual(
    name = "Ancestry",
    values = rev(cbPalette)) +
  ggplot2::facet_grid(.~Analysis,
                     scales = "free",
                     space = "free") +
  ggtitle("a.") +
  theme(legend.position = "bottom",
      plot.title.position = "plot")

p1



# Plot only meta-analysis effects for SNPs
res_meta_snp <- res_meta %>% 
  filter(startsWith(Exposure, "rs"),
         Meta == "Multi-ancestry",
         !is.na(Beta_exposure))

p2 <- ggforestplot::forestplot(
  df = res_meta_snp,
  name = Exposure,
  estimate = Beta_exposure,
  se = SE_exposure,
  pvalue = P_exposure,
  psignif = 0.004,
  shape = Meta,
  xlab = "Percentage change in body weight",
  title = "") +
  ggplot2::scale_shape_manual(
    name = "",
    values = c(23L),
    labels = "") +
  ggforce::facet_col(facets = ~Analysis,
                     scales = "free_y",
                     space = "free") +
  ggtitle("b.") +
  theme(legend.position = "none",
        plot.title.position = "plot")

p2

res_meta_snp_exp <- res_meta_snp %>% 
  select(Study, Ancestry, Exposure, Beta_exposure, P_exposure, Analysis)
res_meta_pgs_exp <- res_meta_pgs %>% 
  select(Study, Ancestry, Exposure, Beta_exposure, P_exposure, Analysis) %>% 
  filter(Study == "Meta_analysis")
#For model B
res_meta_snp_exp <- res_meta_snp %>% 
  select(Study, Ancestry, Exposure, Beta_exposure_W0, P_exposure_W0, Beta_exposure, P_exposure, Analysis)
res_meta_pgs_exp <- res_meta_pgs %>% 
  select(Study, Ancestry, Exposure, Beta_exposure_W0, P_exposure_W0, Beta_exposure, P_exposure, Analysis) %>% 
  filter(Study == "Meta_analysis")

png("/Users/cordioli/Projects/GLP1_bmi/figures_tables/Figure3.png", width = 11.5, height = 5, units = "in", res = 300)
#plot_grid(p1, p2, nrow = 1, rel_widths = c(1.5,1))
plot_grid(p1, p2, nrow = 1)
dev.off()