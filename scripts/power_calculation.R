rm(list = ls())

install.packages("pwrss")
library(pwrss)

# Calculate weighted SD from studies
n <- c(255, 191, 614, 117, 91, 857, 156, 614, 728, 369, 77, 133, 336, 120, 856, 1236)
sd <- c(7.6, 6.61, 7.1, 6.25, 5.74, 8.73, 7.68, 5.56, 6.21, 6.71, 6.05, 6.1, 6, 5.5, 7, 10.87)
var <- sd^2

w_var <- n%*%var / sum(n)
w_sd <- sqrt(w_var)

# https://cran.r-project.org/web/packages/pwrss/vignettes/examples.html#32_Single_Regression_Coefficient_(z_or_t_Test)
# For continuous predictors (PGS)
pwrss.t.reg(beta1 = 0.3,  # 0.3% change in body weight for 1-SD change in PGS
            sdy = 7.8, # Combined standard deviation
            k = 25,  # Total number of predictors
            n = 6750,
            r2 = 0.05, # From HUS
            alpha = 0.05,
            alternative = "not equal")


# For SNPs - formula from Masa's paper

# https://www.medrxiv.org/content/10.1101/2021.09.03.21262975v1.full-text
# (see Methods: Fine-mapping replication analysis)
# "We estimated statistical power via the non-centrality parameter (NCP) of the
# chi-square distribution104. We defined NCP = 2 f (1 – f) n β^2 where f is MAF,
# n is the effective sample size, and β is the expected effect size"

Neff <- 6750
maf <- 0.01
beta <- 0.3

q.tresh <- qchisq(0.05, df=1, lower.tail=F)
ncp <- 2 * maf * (1  - maf) * Neff * beta^2
power <- pchisq(q.tresh, df=1, ncp=ncp, lower.tail = F)
power
# [1] 0.8083475
# [1] 0.9342546