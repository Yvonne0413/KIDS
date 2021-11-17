# Written using R version 4.0.0
library(readstata13)
library(survey)
library(xlsx)
library(data.table)

# Load Data
setwd("C:/projects/RA/KID/")
dat <- read.dta13("LT_CoreData_ChildAbuse_2006_2009_2012KID(5.26).dta")

# Looking at only pediatric patients
dat <- dat[dat$AGE < 18 & !is.na(dat$AGE),]

print(colnames(dat))
cols_to_keep = c("DISCWT", "AGE", "RACE", "YEAR", "ZIPINC_QRTL", "FEMALE", "PAY1", "totcost",
                 "Child_abuse", "DIED", "LOS", "ATYPE", "RECNUM", "xiss", "niss", "MDC", 
                 "HOSP_LOCTEACH", "HOSP_BEDSIZE", "HOSP_REGION", "NACHTYPE", "HOSP_CONTROL",
                 "event_primdx_traumaICD9", "event_anydx_traumaICD9")
dat <- dat[, cols_to_keep]

# Adjust using CPI to 2021 prices
dat[dat$YEAR == 2006,]$totcost <- dat[dat$YEAR == 2006,]$totcost * 1.292034
dat[dat$YEAR == 2009,]$totcost <- dat[dat$YEAR == 2009,]$totcost * 1.214122
dat[dat$YEAR == 2012,]$totcost <- dat[dat$YEAR == 2012,]$totcost * 1.134498

dat$AGEGROUP <- NA
dat[dat$AGE < 18,]$AGEGROUP <- "12-17"
dat[dat$AGE <= 11, ]$AGEGROUP <- "8-11"
dat[dat$AGE <= 7, ]$AGEGROUP <- "4-7"
dat[dat$AGE <= 3, ]$AGEGROUP <- "1-3"
dat[dat$AGE <  1, ]$AGEGROUP <- "<1"

dat$SEVGROUP <- NA
dat[dat$niss <= 75,]$SEVGROUP <- ">25"
dat[dat$niss <= 24, ]$SEVGROUP <- "16-24"
dat[dat$niss <= 15, ]$SEVGROUP <- "0-15"

dat['DIED_ED'] = (dat['DIED'] == 1) & (dat['ATYPE'] == 1)
dat['DIED_IN'] = (dat['DIED'] == 1) & (dat['ATYPE'] != 1)

# Define Sex Labels
dat$SEX <- factor(dat$FEMALE, levels = c(0,1), labels = c("Male", "Female"))
# Define Race Labels according to KID Codebook: https://www.hcup-us.ahrq.gov/db/nation/kid/kiddde.jsp
dat$RACE <- factor(dat$RACE, levels = c(1,2,3,4,5,6), labels = c("White", "Black", "Hispanic", "Asian or Pacific Islander", "Native American", "Other"))
# Define Payer Labels according to KID Codebook: https://www.hcup-us.ahrq.gov/db/nation/kid/kiddde.jsp, excluding medicare (irrelevant)
dat$PAYER <- factor(dat$PAY1, levels = c(2,3,4,5,6), labels = c("Medicaid", "Private", "Self-pay", "No Charge", "Other"))
levels(dat$PAYER) <- c("Medicaid", "Private", "Self-pay", "Other", "Other")
# Define Admit Type Labels according to KID Codebook: https://www.hcup-us.ahrq.gov/db/nation/kid/kiddde.jsp
dat$ADMIT_TYPE <- factor(dat$ATYPE, levels = c(1,2,3,4,5,6), labels = c("Emergency", "Urgent", "Elective", "Newborn", "Trauma Center", "Other"))

dat$DIED <- factor(dat$DIED, levels = c(0, 1), labels = c("FALSE", "TRUE"))
dat$Child_abuse <- factor(dat$Child_abuse, levels = c(0,1), labels = c("Non-SCAN", "SCAN"))
dat$Child_abuse_indicator <- as.numeric(dat$Child_abuse == "SCAN")
dat$ZIPINC_QRTL <- factor(dat$ZIPINC_QRTL, levels = c(1,2,3,4), labels = c("1", "2", "3", "4"))
dat$HOSP_LOCTEACH <- factor(dat$HOSP_LOCTEACH, levels = c(1,2,3), labels = c("Rural", "Urban nonteaching", "Urban teaching"))
dat$HOSP_BEDSIZE <- factor(dat$HOSP_BEDSIZE, levels = c(1,2,3), labels = c("Small", "Medium", "Large"))
dat$HOSP_REGION <- factor(dat$HOSP_REGION, levels = c(1,2,3,4), labels = c("Northeast", "Midwest", "South", "West"))
dat$NACHTYPE <- factor(dat$NACHTYPE, levels = c(0,1,2,3), labels = c("Not Children's", "Children's General", "Children's Specialty", "Children's unit of General"))
dat$HOSP_CONTROL <- factor(dat$HOSP_CONTROL, levels = c(0,1,2,3,4), labels = c("Government or private (collapsed)", "Government, nonfederal (public)", "Private, not-for-profit (voluntary)", "Private, investor-owned (proprietary)", "Private (collapsed category"))
dat$PRIMARY_TRAUMA <- factor(dat$event_primdx_traumaICD9, levels = c(0,1), labels = c("False", "True"))
dat$ANY_TRAUMA <- factor(dat$event_anydx_traumaICD9, levels = c(0,1), labels = c("False", "True"))
dat <- dat[dat$ANY_TRAUMA == "True", ]

# Initialize survey design
dat.w <- svydesign(ids = ~1, data = dat, weights = dat$DISCWT)
options(survey.lonely.psu = "adjust")

# BEGIN: ANALYSIS
svy_freq_table <- function(design, categorical_variables, numerical_variables, column_variable, sheet_name, mean = TRUE){
  # Prep file writer
  wb <- loadWorkbook("Results.xlsx")
  sheets <- getSheets(wb)
  removeSheet(wb, sheetName = sheet_name)
  sheet <- createSheet(wb, sheetName = sheet_name)
  
  # Unique values of group
  groups <- unique(dat[column_variable])
  groups <- groups[!is.na(groups)]
  groups <- sort(groups)
  
  row = 1
  for (factor in categorical_variables){
    # Frequency Distribution
    counts <- round(svytable(as.formula(paste(paste("~",factor,sep=""),column_variable,sep="+")),design))
    # Percentage Distribution
    percent <-round(prop.table(svytable(as.formula(paste(paste("~",factor,sep=""),column_variable,sep="+")),design),margin=2)*100,1)

    # Merge counts and percents
    table <- as.data.frame(cbind(counts, percent))
    colnames(table) <- c(groups, paste(groups,"1",sep="."))
    for(i in 1:length(groups)){
      cols <- c(groups[i], paste(groups[i],"1",sep="."))
      combined <- paste(do.call(paste, c(table[cols],sep=" (")), "%)",sep ="")
      table[groups[i]] <- combined
    }
    table <- table[, (names(table) %in% groups)]
    temp <- rownames(table)
    table <- as.data.table(table)
    table <- as.data.frame(cbind(temp, table))
    colnames(table) <- c(factor, colnames(table)[-1])
    
    # Format output
    cs <- CellStyle(wb, alignment = Alignment(wrapText = TRUE, horizontal = "ALIGN_CENTER"))
    dfColIndex <- rep(list(cs), dim(table)[2])
    names(dfColIndex) <- seq(1, dim(table)[2], by = 1)
    addDataFrame(table, sheet, startRow = row, colStyle = dfColIndex, row.names = FALSE)
    autoSizeColumn(sheet, 1:ncol(table))
    row = row + nrow(table)+1
  }
  
  # Create population level statistics for numerical variables
  for (factor in numerical_variables){
    table = data.frame()
    for(i in 1:length(groups)){
      form = as.formula(paste("~",factor,sep=""))
      des = subset(design, eval(parse(text = column_variable)) == groups[i])
      
      # Mean and standard error
      if (mean){
        res = svymean(form, des, na.rm=TRUE)
        res <- paste(round(coef(res), 3), paste(round(SE(res), 3),")", sep = ""), sep=" (")
      }
      # Median and [Q1, Q3]
      else{
        res = round(unname(svyquantile(form, des, na.rm = TRUE, quantiles = c(.25, .5, .75))), 2)
        res = paste(res[,2],paste("  [", paste( res[,1], paste (",", paste(res[,3], "]", sep = "")), sep = ""), sep = ""), sep = "")          
      }
      table = rbind(table, res)
    }
    table <- as.data.table(t(table))
    colnames(table) <- groups
    table <- as.data.frame(table)
    rownames(table) <- factor
    cs <- CellStyle(wb, alignment = Alignment(wrapText = TRUE, horizontal = "ALIGN_CENTER"))
    dfColIndex <- rep(list(cs), dim(table)[2])
    names(dfColIndex) <- seq(1, dim(table)[2], by = 1)
    addDataFrame(table, sheet, startRow = row, colStyle = dfColIndex, row.names = TRUE)
    autoSizeColumn(sheet, 1:ncol(table))
    row = row + nrow(table)+1
    
  }
  saveWorkbook(wb, "Results.xlsx")
}

svy_freq_table_over_group <- function(design, group, factors, outcome, sheet_name, mean = TRUE, verbose = FALSE){
  # Prep file writer
  wb <- loadWorkbook("Results.xlsx")
  sheets <- getSheets(wb)
  removeSheet(wb, sheetName = sheet_name)
  sheet <- createSheet(wb, sheetName = sheet_name)
  
  # Unique values of group
  groups <- unique(dat[group])
  groups <-groups[!is.na(groups)]
  groups <- sort(groups)
  
  row = 1
  for (factor_choice in factors){
    # Frequency Distribution
    counts <- round(svytable(as.formula(paste(paste("~",factor_choice,sep=""),group,sep="+")),design))
    # Percentage Distribution
    percent <-round(prop.table(svytable(as.formula(paste(paste("~",factor_choice,sep=""),group,sep="+")),design),margin=2)*100,1)
    
    # Median Outcome Distribution
    if (outcome != ""){
      outcomes = c()
      vals <- unique(dat[factor_choice])
      vals <- vals[!is.na(vals)]
      for (val in vals){
        temp = c()
        for (group_choice in groups){
          form = as.formula(paste("~",outcome,sep=""))
          des = subset(design, eval(parse(text = factor_choice)) == val & eval(parse(text = group)) == group_choice)
          # Mean for outcome of interest
          if (mean){
            res = round(unname(svymean(form, des, na.rm = TRUE)), 2)
          }
          # Median and [Q1, Q3] for outcome of interest
          else{
            res = round(unname(svyquantile(form, des, na.rm = TRUE, quantiles = c(.25, .5, .75))), 2)
            res = paste(res[,2],paste("  [", paste( res[,1], paste (",", paste(res[,3], "]", sep = "")), sep = ""), sep = ""), sep = "")          
          }
          temp = cbind(temp, res)
        }
        temp <- as.data.frame(temp)
        outcomes = rbind(outcomes, temp)
      }
      outcomes <- as.data.table(outcomes)
      outcomes <- as.data.frame(cbind(vals, outcomes))
      colnames(outcomes) <- c(factor_choice, groups)
    }
    
    # Chisq test for difference in distribution (not outcome)
    p <- svychisq(as.formula(paste(paste("~",group,sep=""),factor_choice,sep="+")),design, "F")$p.value
    p <- signif(p, 3)
    
    # Merge counts and percents
    table <- as.data.frame(cbind(counts, percent))
    colnames(table) <- c(groups, paste(groups,"1",sep="."))
    for(i in 1:length(groups)){
      cols <- c(groups[i], paste(groups[i],"1",sep="."))
      temp <- do.call(paste, c(table[cols],sep=" ("))
      temp <- paste(temp, "%)",sep ="")
      
      cols <- c(groups[i], paste(groups[i],"1",sep="."))
      table <- table[, !(names(table) %in% cols)]
      table[groups[i]] <- temp
    }
    temp <- rownames(table)
    table <- as.data.table(table)
    table <- as.data.frame(cbind(temp, table))
    
    # Merge incidence with outcomes
    if (outcome != ""){
      cols <- c(paste(groups,"2",sep="."))
      colnames(table) <- c(factor_choice, cols)
      table <- merge(table, outcomes, by = factor_choice)
      for(i in 1:length(groups)){
        cols <- c(paste(groups[i],"2",sep="."), groups[i])
        for (j in 1:length(table[groups[i]])){
          table[groups[i]][j] = format(table[groups[i]][j], format="d", big.mark=",")
        }
        temp <- do.call(paste, c(table[cols],sep="\n "))
        
        cols <- c(groups[i], paste(groups[i],"2",sep="."))
        table <- table[, !(names(table) %in% cols)]
        table[groups[i]] <- temp
      }
    }
    # Add p-values to table
    table$p.value <- c(p, rep("",times=nrow(table)-1))
    if(verbose){
      print(table)
    }
    # Format output
    cs <- CellStyle(wb, alignment = Alignment(wrapText = TRUE, horizontal = "ALIGN_CENTER"))
    dfColIndex <- rep(list(cs), dim(table)[2])
    names(dfColIndex) <- seq(1, dim(table)[2], by = 1)
    addDataFrame(table, sheet, startRow = row, colStyle = dfColIndex, row.names = FALSE)
    autoSizeColumn(sheet, 1:ncol(table))
    row = row + nrow(table)+1
  }
  saveWorkbook(wb, "Results.xlsx")
}

# Table 1: Demographics
design = dat.w
categorical_variables <- c('ADMIT_TYPE', 'AGEGROUP', 'ZIPINC_QRTL', 'SEX', 'PAYER',
                           'HOSP_BEDSIZE', 'HOSP_CONTROL', 'HOSP_LOCTEACH', 'HOSP_REGION', 'NACHTYPE',
                           'RACE', 'SEVGROUP', 'DIED', 'DIED_ED', 'DIED_IN')
numerical_variables <- c('LOS',' totcost', 'AGE', 'niss')
svy_freq_table(design, categorical_variables, numerical_variables, "Child_abuse", "Table 1", mean = TRUE)

# Significance testing
svytable(~Child_abuse, design)
signif(svychisq(~Child_abuse+PAYER, design, "F")$p.value, 3)
signif(svyranktest(niss~Child_abuse, design, "wilcoxon")$p.value, 3)


# Table 2: Race demographics (SCAN only)
design = subset(dat.w, Child_abuse == 'SCAN')
categorical_variables <- c('ADMIT_TYPE', 'AGEGROUP', 'ZIPINC_QRTL', 'SEX', 'PAYER',
                           'HOSP_BEDSIZE', 'HOSP_CONTROL', 'HOSP_LOCTEACH', 'HOSP_REGION', 'NACHTYPE',
                           'SEVGROUP', 'DIED', 'DIED_ED', 'DIED_IN')
numerical_variables <- c('LOS',' totcost', 'AGE', 'niss')
svy_freq_table(design, categorical_variables, numerical_variables, "RACE", "Table 2", mean = TRUE)

# Significance testing
svytable(~ADMIT_TYPE+RACE, design)
signif(svychisq(~ADMIT_TYPE+RACE, design, "F")$p.value, 3)
signif(svyranktest(totcost~RACE, design, "wilcoxon")$p.value, 3)


# Create regressions: Linear if outcome is continuous, Logistic if outcome is categorical
summary(svyglm(Child_abuse_indicator ~ RACE, family = "binomial", design = dat.w, na.action = na.omit))
summary(svyglm(Child_abuse_indicator ~ RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + HOSP_CONTROL + NACHTYPE, family = "binomial", design = dat.w, na.action = na.omit))

summary(svyglm(LOS ~ SEVGROUP + RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + HOSP_CONTROL + NACHTYPE, design = dat.w, na.action = na.omit))
summary(svyglm(LOS ~ SEVGROUP + RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + HOSP_CONTROL + NACHTYPE, design = subset(dat.w, Child_abuse == 'SCAN'), na.action = na.omit))

summary(svyglm(DIED ~ SEVGROUP + RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + HOSP_CONTROL + NACHTYPE, family = "binomial", design = dat.w, na.action = na.omit))
summary(svyglm(DIED ~ SEVGROUP + RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + HOSP_CONTROL + NACHTYPE, family = "binomial", design = subset(dat.w, Child_abuse == 'SCAN'), na.action = na.omit))

summary(svyglm(totcost ~ SEVGROUP + RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + LOS + DIED + HOSP_CONTROL + NACHTYPE, design = dat.w, na.action = na.omit))
summary(svyglm(totcost ~ SEVGROUP + RACE  + SEX + AGEGROUP + PAYER + ZIPINC_QRTL+ HOSP_LOCTEACH + HOSP_BEDSIZE + HOSP_REGION + LOS + DIED + HOSP_CONTROL + NACHTYPE, design = subset(dat.w, Child_abuse == 'SCAN'), na.action = na.omit))


