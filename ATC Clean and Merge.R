
# remove all R objects in the working directory
rm(list=ls(all=TRUE))

#Increase memory size (java out of memory and must be before pkg load)
options(java.parameters = "-Xmx8000m")

# id packages
packages <- c("plyr","dplyr","readxl", "zoo",  "stringi", "xlsx","ggplot2","tidyr","stringr","xlsx")

#install and library
lapply(packages, require, character.only = TRUE)

#set folders
proj_name = "ATC TMP"
base_path="U:/NRMC/Data Science"
SP_path="Y:/"

#set project path
proj_path = paste(base_path,"/",proj_name,sep = "")
out_path = file.path(proj_path, "Output")
cleaned_path = file.path(proj_path, "CleanData")
data_path = file.path(proj_path, "RawData")
share_path  = file.path(SP_path, proj_name)

dir.create(file.path(proj_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(out_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(cleaned_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(data_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(share_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")

#returns string w/o leading or trailing whitespace
trim <- function (x) gsub("^\\s+|\\s+$", "", x)

date <- Sys.Date()

hrp<-read.csv('U:/NRMC/Data Science/ATC TMP/RawData/HRPs.csv',header=T,sep = ',')
hrp$DMIS=str_pad(hrp$DMIS, 4, side = c("left"), pad = "0")
hrp<-hrp[ ,c(5,7)]

scrn<-hrp$DMIS

dat.file <- list.files(file.path(data_path), pattern="*.xlsx", recursive = F, full.names=T, all.files = T, include.dirs = T)
UR="U:/NRMC/Data Science/ATC TMP/RawData/Copy of RHC-A(P) Utilization Rate data.xlsx"
Enroll<- read_excel(UR, sheet = 2, col_names = TRUE, col_types = NULL, na = "",skip = 0)
Enroll=Enroll[ ,1:4];colnames(Enroll)=c('Enroll','Year','Mon','MTF')
Enroll=select(Enroll, everything()) %>%
  group_by(MTF,Year,Mon) %>%
  summarise(Enroll=sum(Enroll))

FC_Work<- read_excel(UR, sheet = 3, col_names = TRUE, col_types = NULL, na = "",skip = 0)
FC_Work=FC_Work[ ,c(6,1,2,3)];colnames(FC_Work)=c('Encounters','Year','Mon','MTF')
FC_Work=select(FC_Work, everything()) %>%
  group_by(MTF,Year,Mon) %>%
  summarise(Encounters=sum(Encounters))

PC_Work<- read_excel(UR, sheet = 4, col_names = TRUE, col_types = NULL, na = "",skip = 0)
PC_Work=PC_Work[ ,c(1,6,7,5)];colnames(PC_Work)=c('Visits','Year','Mon','MTF')
PC_Work=select(PC_Work, everything()) %>%
  group_by(MTF,Year,Mon) %>%
  summarise(Visits=sum(Visits))

UR<-merge(Enroll,FC_Work, by=c('MTF','Year','Mon'),all=T)
UR<-merge(UR,PC_Work,by=c('MTF','Year','Mon'),all=T)
UR$Pure.UR=round((UR$Encounters+UR$Visits)/UR$Enroll,3)
UR$Refine.UR=round((UR$Encounters+(UR$Visits/2))/UR$Enroll,3)
#write.csv(UR,'sample file WIDE.csv',row.names=F)

tmp=as.data.frame(matrix(nrow=0,ncol=6))
Det<- list.files(file.path(data_path, 'DetalCode') , pattern="*.xlsx", recursive = F, full.names=T, all.files = T, include.dirs = T)
Primary.codes<-c('BAZ', 'BAA', 'BGA', 'BGZ', 'BHA', 'BHZ', 'BJA')
PBO_WEX<-c("PBO","WEX")
for (i in Det) {
  #i=Det[3]
  Detail<- read_excel(i, sheet = 'rawdata', col_names = TRUE, col_types = NULL, na = "",skip = 0)
  Detail<-Detail[Detail$txtMEPRS3 %in% Primary.codes, ]
  Detail<-Detail[Detail$txtDetail_Code_None %in% PBO_WEX, ]
  Detail=Detail[ ,c(19,18,9,10,5)]
  colnames(Detail)=c('Service','MTF','Planned','Det.Codes','ClinicName')
  #Detail<-filter(Detail,Service=="NRMC")
  Detail$MTF<-stri_sub(as.character(Detail$MTF), 1, 4)
  Detail$Mon=stri_sub(as.character(i), -7, -6)
  Detail$Year=stri_sub(as.character(i), -11, -8)
  Detail=select(Detail, -Service) %>%
    group_by(MTF,Year,Mon) %>%
    summarise(Planned=mean(Planned), Det.Codes=sum(Det.Codes), Percent.DetCode=Det.Codes/Planned)
  tmp=rbind(tmp,Detail)
}
####Filter by PCP only... BGA BAA BDA 
#one of total plan numbers
#filter for PBO only. (Inhibiting from booking) ask brenda.
tmp$Mon=toupper(month.abb[as.numeric(tmp$Mon)])

Master<-merge(UR,tmp,by=c('MTF','Year','Mon'),all=T)

farm='U:/NRMC/Data Science/ATC TMP/RawData/FILTERED_TOC_FEB.xlsx'
Demand<- read_excel(farm, sheet = 'Parent', col_names = TRUE, col_types = NULL, na = "",skip = 0)
#Demand<-filter(Demand,RHC=="RHC-A")
Demand<-Demand[ ,c(4,32)];colnames(Demand)=c("MTF","Demand")
Demand$MTF=str_pad(Demand$MTF, 4, side = c("left"), pad = "0")
Master<-merge(Master,Demand,by=c('MTF'),all=T)


########################
#Net.Leak<-read_excel('U:/NRMC/Data Science/ATC TMP/RawData/People Campaign Metrics.xlsx', sheet = 'rawdata1', col_names = TRUE, col_types = NULL, na = "",skip = 0)

tmp=as.data.frame(matrix(nrow=0,ncol=4))
NL<- list.files('U:/NRMC/G-8 ACSRM/Decision Support/People Campaign/M2-Leakage' , pattern="*.xls", recursive = F, full.names=T, all.files = T, include.dirs = T)
for (i in NL) {
  #i=NL[4]
  #Net.Leak<- read_excel(i, sheet = 'TEDNI-Detail', col_names = F, col_types = NULL, na = "",skip = 0)
  Net.Leak<- read.xlsx2(i, sheetName='TEDNI-Detail', startRow=3,
                        colIndex=NULL, endRow=NULL, as.data.frame=TRUE, header=TRUE,
                        colClasses="character", stringsAsFactors=FALSE)
  Net.Leak<-Net.Leak[ ,c(4,1,2,20)]; Net.Leak$Number.of.Visits..Total=as.numeric(Net.Leak$Number.of.Visits..Total)
  colnames(Net.Leak)=c('MTF','Year','Mon','Net.Leak.Visits')
  Net.Leak=select(Net.Leak, everything()) %>%
    group_by(MTF,Year,Mon) %>%
    summarise(Net.Leak.Visits=sum(Net.Leak.Visits))
  tmp=rbind(tmp,Net.Leak)
}
#write.csv(tmp,"Net.Leak.csv")
Net.Leak=tmp
Master<-merge(Master,Net.Leak,by=c('MTF'),all=T)

AMR<- list.files(file.path(data_path, 'AMR') , pattern="*.xlsx", recursive = F, full.names=T, all.files = T, include.dirs = T)
tmp=as.data.frame(matrix(nrow=0,ncol=6))
for (i in AMR) {
  #i=AMR[3]
  AO<- read_excel(i, sheet = 'rawdata1', col_names = TRUE, col_types = NULL, na = "",skip = 0)
  AO=AO[ ,c(1,3,32,30)]
  colnames(AO)=c('Branch','Region','MTF','Apts.Offered')
  AO$Region=trim(AO$Region);   AO$Branch=trim(AO$Branch)
  #AO<-filter(AO,Region=="North")
  AO<-filter(AO,Branch=="Army")
  AO$MTF<-stri_sub(as.character(AO$MTF), 1, 4)
  AO$Mon=toupper(month.abb[as.numeric(stri_sub(as.character(i), -7, -6))])
  AO$Year=stri_sub(as.character(i), -11, -8)
  AO=AO[complete.cases(AO), ]
  AO=select(AO, -Region, -Branch) %>%
    group_by(MTF,Year,Mon) %>%
    summarise(Apts.Offered=sum(Apts.Offered))
  tmp=rbind(tmp,AO)
}
Master<-merge(Master,tmp,by=c('MTF','Year','Mon'),all=T)

test=Master[complete.cases(Master), ]
#Master<-merge(Master,tmp,by=c('MTF'),all=T)
library(lubridate)
T2.path<- file.path(data_path, 'Time2BSeen')
T2.file<-file.path(T2.path, '3rd_M.xlsx')
T2<- read_excel(T2.file, sheet = 'rawdata1', col_names = TRUE, col_types = NULL, na = "",skip = 0)
T2<-filter(T2,FacilityService=="Army")
#T2<-filter(T2,IntermediateCommand=="RHC-A")
T2$PDMIS=str_pad(T2$PDMIS, 4, side = c("left"), pad = "0")
T2$Mon=toupper(month.abb[as.numeric(month(T2$Period, label = FALSE, abbr = TRUE))])
T2$Year=year(T2$Period)
T2<-T2[ ,c(4,34,33,14)]
colnames(T2)=c('MTF','Year','Mon','Acute3rd')

T2=T2[complete.cases(T2), ]
T2=select(T2, everything()) %>%
  group_by(MTF,Year,Mon) %>%
  summarise(Acute3rd=mean(Acute3rd))
Master<-merge(Master,T2,by=c('MTF','Year','Mon'),all=T)

APLS<-'U:/NRMC/Data Science/ATC TMP/RawData/APLSS/ATC APLSS Data Collection Worksheet as of 8 Apr 2016.xlsx'  
APLS<- read_excel(APLS, sheet = 2, col_names = TRUE, col_types = NULL, na = "",skip = 0)
APLS$MTF=toupper(APLS$MTF)
APLS<-merge(APLS,hrp, by.x='MTF', by.y='Short', all=T)
APLS<-APLS[ ,c(7,2,3,4,5)]
APLS$Mon=toupper(month.abb[as.numeric(month(APLS$Month, label = FALSE, abbr = TRUE))])
APLS$Year=year(as.POSIXct(APLS$Month))
APLS$Month=NULL
colnames(APLS)[1]='MTF'


Master<-merge(Master,APLS,by=c('MTF','Year','Mon'),all=T) 

leak='U:/NRMC/Data Science/ATC TMP/RawData/Primary Care Portal (Network Leakage) RHC-A.xlsx'
leakage<- read_excel(leak, sheet ='EPCP' , col_names = TRUE, col_types = NULL, na = "",skip = 4)
leakage=leakage[ ,c(1,4,5,6,9,8,15,17,18)]
colnames(leakage)[1:3]<-c('MTF','Year','Mon')
leakage$Mon=toupper(month.abb[as.numeric(month(leakage$Mon, label = FALSE, abbr = TRUE))])

Master<-merge(Master,leakage,by=c('MTF','Year','Mon'),all=T) 


Master$Date.ID=paste(Master$Year,str_pad(match(paste(toupper(stri_sub(Master$Mon, 1, 1))
                                                     ,tolower(stri_sub(Master$Mon, -2, -1)),sep="")
                                               ,month.abb),2,side = c("left"), pad = "0")
                     ,sep=".")

#UR to DemandMaster (what demand we need to have)
#this is what the monthly demand is  for my population
Master$Demand.Needed.NEWby12=round((Master$Refine.UR*Master$Enroll)/12,2)
Master$Demand.Needed.NEW=round((Master$Refine.UR*Master$Enroll),2)
Master$'MTF_Enc/Demand.Needed' =round(Master$Encounters/Master$Demand.Needed.NEW,2)
Master$'MTF_Enc/Demand' =round(Master$Encounters/Master$Demand,2)
Master$'MTF_Enc/Demand.Needed' =round(Master$Encounters/Master$Demand.Needed.NEW,2)
Master$'Net_Enc/Demand.Needed' =round(Master$Net.Leak.Visits/Master$Demand.Needed,2)
Master$'Planned/Demand'=round(Master$Planned/Master$Demand,2)
write.csv(Master, file.path(out_path,'All Data.csv'), row.names = F)
Region=Master[Master$MTF %in% scrn , ]
write.csv(Region, file.path(out_path,'All Region Data.csv'), row.names = F)
#test=Region[complete.cases(Region), ]
#write.csv(test, file.path(out_path,'Region Complete Only.csv'), row.names = F)
test=Region
test=test[ ,c(1:3,24,25,4:23)]
tst=test
dec.cols<-colnames(test)[5:25]
for (i in dec.cols) {
  #i=dec.cols[1]
  name.dec.col=paste(i,"_Decile",sep="") #Create col name
  new.col=ncol(tst)+1 #ID Location of col
  tst[ ,new.col]<-ntile(tst[ ,i],10) #Add data to col
  colnames(tst)[new.col]=name.dec.col #Rename col
}
#Deciled.Remove.Raw=tst[ ,-4:-23]
#Deciled.Remove.Raw=gather(Deciled.Remove.Raw,key=Metric,value=Value,-MTF,-Year,-Mon, -Date.ID)
#Deciled.Remove.Raw=Deciled.Remove.Raw[complete.cases(Deciled.Remove.Raw), ]
write.csv(tst,file.path(out_path,'Region file WIDE DECILE.csv'),row.names=F)

tst=test
dec.cols<-colnames(test)[4:23]
for (i in dec.cols) {
  #i=dec.cols[1]
  name.dec.col=paste(i,"_Decile",sep="")
  new.col=ncol(tst)+1
  tst[ ,new.col]<-ntile(tst[ ,i],50)
  colnames(tst)[new.col]=name.dec.col
}
#Fifty.Remove.Raw=tst[ ,-4:-23]
#Fifty.Remove.Raw=gather(Fifty.Remove.Raw,key=Metric,value=Value,-MTF,-Year,-Mon, -Date.ID)
#Fifty.Remove.Raw=Fifty.Remove.Raw[complete.cases(Fifty.Remove.Raw), ]
write.csv(tst,file.path(out_path,'Region file WIDE FIFTY.csv'),row.names=F)

###################################



Final=gather(Region,key=Metric,value=Value,-MTF,-Year,-Mon, -Date.ID)
Final=Final[complete.cases(Final), ]
Final2=spread(Final, key=Metric, value=Value, fill = NA, convert = FALSE, drop = TRUE)

Summary=select(Final, everything()) %>%
  group_by(MTF, Metric) %>%
  summarise(Mean=round(mean(Value),2), Max=round(mean(Value),2), Min=round(min(Value),2),Median=round(median(Value),2))
write.csv(Summary,file.path(out_path,'Summary by Metric.csv'),row.names=F)
#######################

write.csv(Final,'Region file LONG.csv',row.names=F)
Final2=Final[complete.cases(Final), ]
write.csv(Final2,'Region file LONG with NO NAs.csv',row.names=F)


