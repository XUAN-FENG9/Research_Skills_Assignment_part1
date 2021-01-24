#load R packages
library(tidyverse)
library(car)
library(readxl)
library(writexl)
library(tidyr)
library(dbplyr)
library(plm)
library(pheatmap)
library(lubridate)


#load data from COMPUSTAT
COMPUSTAT<-read_xlsx("D:/COMPUSTAT.xlsx")
#define variables
COMPUSTAT<-rename(COMPUSTAT,"year"="Data Year - Fiscal")
COMPUSTAT$leverage<-(COMPUSTAT$`Debt in Current Liabilities - Total`+COMPUSTAT$`Long-Term Debt - Total`)/COMPUSTAT$`Assets - Total`
COMPUSTAT$short_term_leverage<-COMPUSTAT$`Debt in Current Liabilities - Total`/COMPUSTAT$`Assets - Total`
COMPUSTAT$long_term_leverage<-COMPUSTAT$`Long-Term Debt - Total`/COMPUSTAT$`Assets - Total`
COMPUSTAT$net_PPE_ratio<-COMPUSTAT$`Property, Plant and Equipment - Total (Net)`/COMPUSTAT$`Assets - Total`
COMPUSTAT$gross_PPE_ratio<-COMPUSTAT$`Property, Plant and Equipment - Total (Gross)`/COMPUSTAT$`Assets - Total`
COMPUSTAT$inventory_ratio<-COMPUSTAT$`Inventories - Total`/COMPUSTAT$`Assets - Total`
COMPUSTAT$receivable_ratio<-COMPUSTAT$`Receivables - Trade`/COMPUSTAT$`Assets - Total`
COMPUSTAT$cash_ratio<-COMPUSTAT$`Cash and Short-Term Investments`/COMPUSTAT$`Assets - Total`
COMPUSTAT$firm_size<-log(COMPUSTAT$`Assets - Total`)
COMPUSTAT$crisis<-ifelse(COMPUSTAT$year%in%c(2007,2008,2009),1,0)


#load data from CRSP
CRSP<-read_xlsx("D:/CRSP.xlsx")

#generate the 12 month rolling-window return from each December for the CRSP data
CRSP$year<-year(CRSP$`Names Date`)
CRSP<-CRSP%>%group_by(`CUSIP Header`,year)%>%mutate(rolling_return=prod(1+Returns)-1)

#Merge the two dataset
#change the colum names of both table and make them same
CRSP<-rename(CRSP,"CUSIP"="CUSIP Header")

#becasue the CUSIP in CRSP dataset only have the first 8 digits, change the CUSIP in COMPUSTAT to keep align
COMPUSTAT$CUSIP<-substr(COMPUSTAT$CUSIP,1,8)
#merge two datasets and create the master file
CRSP1<-CRSP[,c(4,7,6)]
COMPUSTAT1<-COMPUSTAT[,-c(1,2,4,6,7:14)]
master<-merge(CRSP1,COMPUSTAT1,by=c("CUSIP","year"))
#delete duplicate rows
master<-master[!duplicated(master),]
#delete all observations that have NA for one or more variables
master<-na.omit(master)

#save file
write_xlsx(master,"D:/master_file.xlsx")

#Question 1. Provide a table of descriptive statistics, including the mean, median, standard deviation, skewness,kurtosis of every variable.
#Step1: for continuous variables
mystats<-function(x){
  m1<-mean(x)
  m2<-median(x)
  n<-length(x)
  s<-sd(x)
  skew<-sum((x-m1)^3/s^3)/n
  kurt<-sum((x-m1)^4/s^4)/n-3
  low<-quantile(x,0.05)
  high<-quantile(x,0.95)
  nout<-round(sum(x<low |x>high),1)
  return(c(length=n,mean=m1,median=m2,stdev=s,skew=skew,kurt=kurt,lowlier=low,highlier=high,outliernumber=nout))
}
stats<-sapply(master[,-c(1,2,13)],mystats)
stats
#Step2: for binary variables - crisis
sum(master$crisis)
#modify outliers
deoutliers<-function(x){
  low<-quantile(x,0.05)
  high<-quantile(x,0.95)
  x<-ifelse(x<low,low,x)
  x<-ifelse(x>high,high,x)                   
}

master[,-c(1,2,13)]<-sapply(master[,-c(1,2,13)],deoutliers)

#QUestion 2: Indicate how leverage (total, short-term and long-term) is affected by the independent and control variables by providing a correlation table.
Cor1<-cor(master[,-c(1,2)])
pheatmap(Cor1)
# Correlation matrix only for independent and control variables:
Cor2<-cor(master[,-c(1,2,4,5,6)])
pheatmap(Cor2)

#Question 3: Estimate the following pooled OLS regression
# create variables at t-1 time
master_lag<-read_xlsx("D:/master_lag.xlsx")
master_lag<-master_lag[duplicated(master_lag$CUSIP),]
#3.1.with control variables
OLScon<-lm(leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,
           data = master_lag)
summary(OLScon)
ncvTest(OLScon) #test for heteroscedasticity
vif(OLScon)
#3.2. without control variables
OLSnocon<-lm(leverage~net_PPE_lag+inventory_lag+receivable_lag,data = master_lag)
summary(OLSnocon)
ncvTest(OLSnocon) #test for heteroscedasticity

#Question 4: Panel regression
pmaster<-pdata.frame(master_lag)
#4.1.With time fixed effects
pregt<-plm(leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,
           data = pmaster,effect = "time",model = "within")
summary(pregt)
#4.2.With firm fixed effects
pregf<-plm(leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,
           data = pmaster,effect = "individual",model = "within")
summary(pregf)
#4.3.With time and firm fixed effects
pregtf<-plm(leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,
           data = pmaster,effect = "twoways",model = "within")
summary(pregtf)
#4.4 with random effects
pregran<-plm(leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,
             data = pmaster,model = "random")
summary(pregran)
#4.5 compare models
pFtest(pregt,OLScon)
pFtest(pregf,OLScon)
pFtest(pregtf,OLScon)
phtest(pregran,pregtf)

#Question 6
#OLS regression-short term
regs<-lm(short_term_leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,data = master_lag)
summary(regs)
#OLS regression-long term
regl<-lm(long_term_leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,data = master_lag)
summary(regl)
#panel regression-short term leverage
pregs<-plm(short_term_leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,data = pmaster,effect = "time",model = "within")
summary(pregs)
#panel regression-long term leverage
pregl<-plm(long_term_leverage~net_PPE_lag+inventory_lag+receivable_lag+size_lag+crisis+return_lag,data = pmaster,effect = "time",model = "within")
summary(pregl)

#Question 7
reg7<-lm(leverage~net_PPE_lag+inventory_lag+receivable_lag+
           net_PPE_lag:cash_lag+inventory_lag:cash_lag+receivable_lag:cash_lag,data = master_lag)
summary(reg7)
#Question 8
reg8<-lm(leverage~net_PPE_lag+inventory_lag+receivable_lag+
           crisis:net_PPE_lag+crisis:inventory_lag+crisis:receivable_lag,data = master_lag)
summary(reg8)
#Question 9
#generate a dummy variable naming 'D9' that equals 1 if the firm's rolling return is in the top 20% in the given year.
master_lag<-master_lag%>%group_by(year)%>%mutate(D9=ifelse(return_lag>=quantile(master_lag$return_lag,0.8),1,0))
#interaction regression
reg9<-lm(leverage~net_PPE_lag+inventory_lag+receivable_lag+
           D9:net_PPE_lag+D9:inventory_lag+D9:receivable_lag,data = master_lag)
summary(reg9)