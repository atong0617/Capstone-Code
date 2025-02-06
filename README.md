setwd("~/NSCI 400 Capstone Project")

if (!require("pacman")) install.packages("pacman") 
pacman::p_load(readxl, plyr, lme4,nlme, robustlmm, car, broom, lsmeans, ggplot2, psych, HLMdiag, tableone,mice)


data.cog <- read_excel("RVCI_MASTER_20240508.xlsx")
data.cog.baseline <- subset(data.cog, Timepoint=="Baseline")
data.cog.baseline1 <- subset(data.cog.baseline, select = c(1,27,48,68))
data.cog.baseline2 <- data.cog.baseline1 %>%
  rename(TMT_BA = `...27`,
         Stroop_32 = `...48`,
         MoCA = `...68`)

data.cog.baseline2$`TMT_BA` <- as.numeric(as.character(data.cog.baseline2$`TMT_BA`))
data.cog.baseline2$`Stroop_32` <- as.numeric(as.character(data.cog.baseline2$`Stroop_32`))
data.cog.baseline2$`MoCA` <- as.numeric(as.character(data.cog.baseline2$`MoCA`))
data.cog.baseline2 <- na.omit(data.cog.baseline2)

data.cog.baseline2 <- data.cog.baseline2[-56, ]
data.cog.baseline2 <- data.cog.baseline2[-47, ]


data.phys <- read_excel("RVCI_MASTER_20240508.xlsx", sheet = "Physical")
data.phys.baseline <- subset(data.phys, Timepoint=="Baseline")
data.phys.baseline1 <- subset(data.phys.baseline, select = c(1,4,5,9,35,65,96)) 
data.phys.baseline2 <- rename(data.phys.baseline1, 
                              Age = `Intake Form`,
                              Sex = `...5`,
                              `BMI` = `...9`,
                              PWV = `...35`,
                              Grip_Strength = `...65`,
                              Meters_Walked = `6 Minute Walk Test`)

data.phys.baseline2 <- data.phys.baseline2 %>%
  mutate(Sex = recode(Sex, "Female" = 0, "Male" = 1))
data.phys.baseline2$`Age` <- as.numeric(as.character(data.phys.baseline2$`Age`))
data.phys.baseline2$`BMI` <- as.numeric(as.character(data.phys.baseline2$`BMI`))
data.phys.baseline2$`PWV` <- as.numeric(as.character(data.phys.baseline2$`PWV`))
data.phys.baseline2$`Grip_Strength` <- as.numeric(as.character(data.phys.baseline2$`Grip_Strength`))
data.phys.baseline2$`Meters_Walked` <- as.numeric(as.character(data.phys.baseline2$`Meters_Walked`))
data.phys.baseline2 <- na.omit(data.phys.baseline2)

data.phys.baseline2$VO2max <- ((0.1 * (data.phys.baseline2$Meters_Walked / 6) + 3.5) / 3.5) * 3.5


# Descriptives
describeBy(data.phys.baseline2$Age)

table(data.phys.baseline2$Sex)
prop.table(table(data.phys.baseline2$Sex))

describeBy(data.phys.baseline2$`BMI`)

describeBy(data.phys.baseline2$`PWV`)

describeBy(data.phys.baseline2$`Grip_Strength`)

describeBy(data.phys.baseline2$`Meters_Walked`)

describeBy(data.phys.baseline2$VO2max)

describeBy(data.cog.baseline2$`TMT_BA`)

describeBy(data.cog.baseline2$`Stroop_32`)

describeBy(data.cog.baseline2$`MoCA`)



# Moderation Analysis

data.cog.phys.baseline <- merge(data.phys.baseline2, data.cog.baseline2, by = "ID")



# Model 1 CRF(VO2max) and TMT
model.CRF.TMT <- lm(TMT_BA ~ PWV * VO2max + Age + Sex + BMI + MoCA, data.cog.phys.baseline)
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
model.CRF.Stroop <- lm(Stroop_32 ~ PWV * VO2max + Age + Sex + BMI + MoCA, data.cog.phys.baseline)
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



# Model 3 MS and TMT
model.MS.TMT <- lm(TMT_BA ~ PWV * Grip_Strength + Age + Sex + BMI + MoCA, data.cog.phys.baseline)
summary(model.MS.TMT) 

# Visualizing interaction effect
interactions :: interact_plot(model.MS.TMT, pred = PWV, modx = Grip_Strength)

# Linearity & Additivity assumption
plot(model.MS.TMT$fitted.values, residuals(model.MS.TMT),
     xlab = "Fitted Values", 
     ylab = "Residuals", 
     main = "Residuals vs. Fitted Values",
     pch = 20, col = "blue")
abline(h = 0, col = "red", lwd = 2) 

# Independence of Residuals assumption 
print(dwt(model.MS.TMT))

# Homoscedasticity assumption
bptest(model.MS.TMT)

# Normality of Residuals assumption 
qqnorm(residuals(model.MS.TMT))  # Q-Q plot
qqline(residuals(model.MS.TMT), col = "red")  # Add a reference line (ideal normal distribution)

shapiro.test(resid(model.MS.TMT))

# Multicollinearity assumption 
print(vif(model.MS.TMT))

# Outliers and Influential Observations
print(outlierTest(model.MS.TMT))



# Model 4 MS and Stroop
model.MS.Stroop <- lm(Stroop_32 ~ PWV * Grip_Strength + Age + Sex + BMI + MoCA, data.cog.phys.baseline)
summary(model.MS.Stroop) 

# Visualizing interaction effect
interactions :: interact_plot(model.MS.Stroop, pred = PWV, modx = Grip_Strength)

# Linearity & Additivity assumption
plot(model.MS.Stroop$fitted.values, residuals(model.MS.Stroop),
     xlab = "Fitted Values", 
     ylab = "Residuals", 
     main = "Residuals vs. Fitted Values",
     pch = 20, col = "blue")
abline(h = 0, col = "red", lwd = 2) 

# Independence of Residuals assumption 
print(dwt(model.MS.Stroop))

# Homoscedasticity assumption
bptest(model.MS.Stroop)

# Normality of Residuals assumption 
qqnorm(residuals(model.MS.Stroop))  # Q-Q plot
qqline(residuals(model.MS.Stroop), col = "red")  # Add a reference line (ideal normal distribution)

shapiro.test(resid(model.MS.Stroop))

# Multicollinearity assumption 
print(vif(model.MS.Stroop))

# Outliers and Influential Observations
print(outlierTest(model.MS.Stroop))



# Simple Slope Analysis
# Model 1 CRF(VO2max) and TMT

library(emmeans)

ss.model <- lm(TMT_BA ~ PWV * VO2max + Age + Sex + BMI + MoCA, data = data.cog.phys.baseline)

zvals <- c(mean(data.cog.phys.baseline$VO2max) - sd(data.cog.phys.baseline$VO2max), 
           mean(data.cog.phys.baseline$VO2max), 
           mean(data.cog.phys.baseline$VO2max) + sd(data.cog.phys.baseline$VO2max))

emtrends(ss.model, specs = "VO2max", var = "PWV", at = list(VO2max = zvals)) |> test()


# Model 2 CRF(VO2max) and Stroop
ss.model2 <- lm(Stroop_32 ~ PWV * VO2max + Age + Sex + BMI + MoCA, data = data.cog.phys.baseline)

emtrends(ss.model2, specs = "VO2max", var = "PWV", at = list(VO2max = zvals)) |> test()
