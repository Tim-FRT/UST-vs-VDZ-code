library(MatchIt)
library(xlsx)
library(survey)
library(tableone)
library(xlsx)

# Step 1: Read Data
aa=read.xlsx("F:\\multicenter UST vs VDZ.xlsx",sheetIndex=1)
aa <- aa[aa$use26 == 1, ]
View(aa)
names(aa)
attach(aa)
vars=c("sex","age","course.of.disease","smoke","Previous.bowel.resection",
       "location","L4","behavior","fistula","perianal.diseases","TNF.naive",
       "previous.EEN","previous.steroid","previous.immunomodulators","basic.CDAI","CDAI.150",
       "EEN.at.baseline","corticosteroid.at.baseline","immunomodulators.at.baseline",
       "CRP0","CRP.5","ALB0","ALB.35",
       "clinical.remission26","steroid.free.remission26","no.optimization.clinical.remission26","no.optimization.steroid.free.remission26",
       "objective.response26","objective.remission26","no.optimization.objective.response26","no.optimization.objective.remission26",
       "clinical.remission52","steroid.free.remission52","no.optimization.clinical.remission52","no.optimization.steroid.free.remission52",
       "objective.response52","objective.remission52","no.optimization.objective.response52","no.optimization.objective.remission52",
       "CRPremission26","CRPremission52") 
fvars<-c("sex","smoke","Previous.bowel.resection",
         "location","L4","behavior","fistula","perianal.diseases","TNF.naive",
         "previous.EEN","previous.steroid","previous.immunomodulators","CDAI.150",
         "EEN.at.baseline","corticosteroid.at.baseline","immunomodulators.at.baseline",
         "CRP.5","ALB.35",
         "clinical.remission26","steroid.free.remission26","no.optimization.clinical.remission26","no.optimization.steroid.free.remission26",
         "objective.response26","objective.remission26","no.optimization.objective.response26","no.optimization.objective.remission26",
         "clinical.remission52","steroid.free.remission52","no.optimization.clinical.remission52","no.optimization.optimization.remission52",
         "objective.response52","objective.remission52","no.optimization.objective.response52","no.optimization.objective.remission52",
         "endoscopic.response26","endoscopic.remission26","endoscopic.response52","endoscopic.remission52",
         "ultrasound.response26","ultrasound.remission26","ultrasound.response52","ultrasound.remission52",
         "radiologic.response26","radiologic.remission26","radiologic.response52","radiologic.remission52",
         "CRPremission26","CRPremission52") 

# step 2: Create a baseline table before adjustment  
tab2<-CreateTableOne(vars=vars,strata="group",data=aa, factorVars=fvars,addOverall=TRUE)
print(tab2)

# step 3: Construct logistic regression models
m.out<-matchit(group~sex+age+course.of.disease+smoke+Previous.bowel.resection+location+L4+
                 behavior+fistula+perianal.diseases+TNF.naive+previous.EEN+previous.steroid+
                 previous.immunomodulators+basic.CDAI+CDAI.150+EEN.at.baseline+corticosteroid.at.baseline+
                 immunomodulators.at.baseline+CRP0+CRP.5+ALB0+ALB.35,data=aa,method="nearest",caliper=0.15,ratio=2)
summary(m.out)
print(m.out)
matchdata1=match.data(m.out)

# Step4：Create a baseline table after adjustment
mBL1<-tableone::CreateTableOne(vars=vars,strata="group",data=matchdata1, factorVars=fvars)
print(mBL1,showAllLevels=TRUE)
library(foreign)
matchdata1$id<-1:nrow(matchdata1)
write.xlsx(matchdata1,file="F:\\c.xlsx")
