setwd("~/NSCI 400 Capstone Project")

if (!require("pacman")) install.packages("pacman") 
pacman::p_load(readxl, plyr, lme4,nlme, robustlmm, car, broom, lsmeans, ggplot2, psych, HLMdiag, tableone,mice)


data.cog <- read_excel("RVCI_MASTER_20240508.xlsx")
data.cog.baseline <- subset(data.cog, Timepoint=="Baseline")
data.cog.baseline1 <- subset(data.cog.baseline, select = -c(2,4:20,28:38,49:75))
data.cog.baseline2 <- rename(data.cog.baseline1, c("...21"="Trails A - Time(sec)", 
                                                   "...22"="Trails A - Comission Errors", 
                                                   "...23"="Trails A - Omission Errors", 
                                                   "...24"="Trails B - Time(sec)", 
                                                   "...25"="Trails B - Comission Errors", 
                                                   "...26"= "Trails B - Omission Errors", 
                                                   "...27"="Trails B minus A",
                                                   "...39"="Stroop 1 Time(sec)", 
                                                   "...40"="Stroop 1 Uncorrected Errors", 
                                                   "...41"="Stroop 1 Corrected Errors", 
                                                   "...42"="Stroop 2 Time(sec)", 
                                                   "...43"="Stroop 2 Uncorrected Errors", 
                                                   "...44"="Stroop 2 Corrected Errors", 
                                                   "...45"="Stroop 3 Time(sec)", 
                                                   "...46"="Stroop 3 Uncorrected Errors", 
                                                   "...47"="Stroop 3 Corrected Errors", 
                                                   "...48"="Stroop 3 - Stroop 2 Time(sec)"))
data.cog.baseline2 <- na.omit(data.cog.baseline2)
data.cog.baseline2 <- data.cog.baseline2[-56, ]
data.cog.baseline2 <- data.cog.baseline2[-47, ]
data.cog.baseline3 <- subset(data.cog.baseline2, select = c(1,9,19))



data.phys <- read_excel("RVCI_MASTER_20240508.xlsx", sheet = "Physical")
data.phys.baseline <- subset(data.phys, Timepoint=="Baseline")
data.phys.baseline1 <- subset(data.phys.baseline, select = c(1,3:5,9,35,65,96)) 
data.phys.baseline2 <- rename(data.phys.baseline1, c("Intake Form"="Age", 
                                                     "...5"="Sex", 
                                                     "...9"="BMI (kg/m^2)", 
                                                     "...35"="CF PWV", 
                                                     "...65"="Grip Strength (Total Score)", 
                                                     "6 Minute Walk Test"="6MWT (Meters Walked)"))
data.phys.baseline2 <- na.omit(data.phys.baseline2)
data.phys.baseline3 <- subset(data.phys.baseline2, select = c(1,6:8))




# Descriptives
data.phys.baseline2$Age <- as.numeric(as.character(data.phys.baseline2$Age))
describeBy(data.phys.baseline2$Age)

table(data.phys.baseline2$Sex)
prop.table(table(data.phys.baseline2$Sex))

data.phys.baseline2$`BMI (kg/m^2)` <- as.numeric(as.character(data.phys.baseline2$`BMI (kg/m^2)`))
describeBy(data.phys.baseline2$`BMI (kg/m^2)`)

data.phys.baseline2$`CF PWV` <- as.numeric(as.character(data.phys.baseline2$`CF PWV`))
describeBy(data.phys.baseline2$`CF PWV`)

data.phys.baseline2$`Grip Strength (Total Score)` <- as.numeric(as.character(data.phys.baseline2$`Grip Strength (Total Score)`))
describeBy(data.phys.baseline2$`Grip Strength (Total Score)`)

data.phys.baseline2$`6MWT (Meters Walked)` <- as.numeric(as.character(data.phys.baseline2$`6MWT (Meters Walked)`))
describeBy(data.phys.baseline2$`6MWT (Meters Walked)`)

describeBy(data.cog.phys.baseline$VO2max)

data.cog.baseline2$`Trails A - Time(sec)` <- as.numeric(as.character(data.cog.baseline2$`Trails A - Time(sec)`))
describeBy(data.cog.baseline2$`Trails A - Time(sec)`)

data.cog.baseline2$`Trails B - Time(sec)` <- as.numeric(as.character(data.cog.baseline2$`Trails B - Time(sec)`))
describeBy(data.cog.baseline2$`Trails B - Time(sec)`)

data.cog.baseline2$`Trails B minus A` <- as.numeric(as.character(data.cog.baseline2$`Trails B minus A`))
describeBy(data.cog.baseline2$`Trails B minus A`)

data.cog.baseline2$`Stroop 1 Time(sec)` <- as.numeric(as.character(data.cog.baseline2$`Stroop 1 Time(sec)`))
describeBy(data.cog.baseline2$`Stroop 1 Time(sec)`)

data.cog.baseline2$`Stroop 2 Time(sec)` <- as.numeric(as.character(data.cog.baseline2$`Stroop 2 Time(sec)`))
describeBy(data.cog.baseline2$`Stroop 2 Time(sec)`)

data.cog.baseline2$`Stroop 3 Time(sec)` <- as.numeric(as.character(data.cog.baseline2$`Stroop 3 Time(sec)`))
describeBy(data.cog.baseline2$`Stroop 3 Time(sec)`)

data.cog.baseline2$`Stroop 3 - Stroop 2 Time(sec)` <- as.numeric(as.character(data.cog.baseline2$`Stroop 3 - Stroop 2 Time(sec)`))
describeBy(data.cog.baseline2$`Stroop 3 - Stroop 2 Time(sec)`)





# Moderation Analysis

data.cog.phys.baseline <- merge(data.phys.baseline3, data.cog.baseline3, by = "ID")
data.cog.phys.baseline <- rename(data.cog.phys.baseline, c("CF PWV"="PWV",
                                                           "Grip Strength (Total Score)"="Grip_Strength",
                                                           "6MWT (Meters Walked)"="Meters_Walked",
                                                           "Trails B minus A"="TMT-B-A",
                                                           "Stroop 3 - Stroop 2 Time(sec)"="Stroop3-2"))
data.cog.phys.baseline <- rename(data.cog.phys.baseline, c("TMT-B-A"="TMT_BA",
                                                           "Stroop3-2"="Stroop_32"))

data.cog.phys.baseline$VO2max <- ((0.1 * (data.cog.phys.baseline$Meters_Walked / 6) + 3.5) / 3.5) * 3.5



# Model 1 CRF and TMT
model.CRF.TMT <- lm(TMT_BA ~ PWV * VO2max, data.cog.phys.baseline)
summary(model.CRF.TMT)

# Visualizing interaction effect
interactions :: interact_plot(model.CRF.TMT, pred = PWV, modx = VO2max)

# Linearity & Additivity assumption
plot(model.CRF.TMT$fitted.values, residuals(model.CRF.TMT),
     xlab = "Fitted Values", 
     ylab = "Residuals", 
     main = "Residuals vs. Fitted Values",
     pch = 20, col = "blue")
abline(h = 0, col = "red", lwd = 2)

# Independence of Residuals assumption 
print(dwt(model.CRF.TMT))

# Homoscedasticity assumption
bptest(model.CRF.TMT)

# Normality of Residuals assumption 
qqnorm(residuals(model.CRF.TMT))  # Q-Q plot
qqline(residuals(model.CRF.TMT), col = "red")  # Add a reference line (ideal normal distribution)

shapiro.test(resid(model.CRF.TMT))

# Multicollinearity assumption 
print(vif(model.CRF.TMT))

# Outliers and Influential Observations
print(outlierTest(model.CRF.TMT))



# Model 2 CRF and Stroop
model.CRF.Stroop <- lm(Stroop_32 ~ PWV * VO2max, data.cog.phys.baseline)
summary(model.CRF.Stroop)

# Visualizing interaction effect
interactions :: interact_plot(model.CRF.Stroop, pred = PWV, modx = VO2max)

# Linearity & Additivity assumption
plot(model.CRF.Stroop$fitted.values, residuals(model.CRF.Stroop),
     xlab = "Fitted Values", 
     ylab = "Residuals", 
     main = "Residuals vs. Fitted Values",
     pch = 20, col = "blue")
abline(h = 0, col = "red", lwd = 2) 

# Independence of Residuals assumption 
print(dwt(model.CRF.Stroop))

# Homoscedasticity assumption
bptest(model.CRF.Stroop)

# Normality of Residuals assumption 
qqnorm(residuals(model.CRF.Stroop))  # Q-Q plot
qqline(residuals(model.CRF.Stroop), col = "red")  # Add a reference line (ideal normal distribution)

shapiro.test(resid(model.CRF.Stroop))

# Multicollinearity assumption 
print(vif(model.CRF.Stroop))

# Outliers and Influential Observations
print(outlierTest(model.CRF.Stroop)) 
