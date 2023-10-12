library(tableone)
library(survey)
library(reshape2)
rm(list = ls())
library(xlsx)
library(ggplot2)

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

for (i in fvars){aa[,i] <- as.factor(aa[,i])}


# step 2: Create a baseline table before adjustment  
tab_Unmatched <- CreateTableOne(vars = vars, 
                                strata = "group", 
                                factorVars=fvars,
                                data = aa, 
                                test =T)

print(tab_Unmatched,showAllLevels=TRUE,smd = TRUE)
addmargins(table(ExtractSmd(tab_Unmatched) > 0.1))

mydata1<-print(tab_Unmatched,showAllLevels=TRUE,smd = TRUE)

# step 3: Construct logistic regression models
psModel<-glm(group1~sex+age+course.of.disease+smoke+Previous.bowel.resection+location+L4+
               behavior+fistula+perianal.diseases+TNF.naive+previous.EEN+previous.steroid+
               previous.immunomodulators+basic.CDAI+CDAI.150+EEN.at.baseline+corticosteroid.at.baseline+
               immunomodulators.at.baseline+CRP0+CRP.5+ALB0+ALB.35,
             family=binomial(link="logit"),data=aa)

aa$ps=predict(psModel,type="response")

head(aa$ps)

aa$IPTW<-ifelse(aa$group1==1,1/aa$ps,1/(1-aa$ps))

dataIPTW=svydesign(ids=~1,data=aa,weights= ~IPTW)

# Step4：Create a baseline table after adjustment
tab_IPTW=svyCreateTableOne(vars=vars, strata="group1",data=dataIPTW,test=T)
print(tab_IPTW,showAllLevels=TRUE,smd=TRUE)


# Step5: Calculate P value, OR and 95%CI
bb<-glm(formula=clinical.remission26~group1,family=binomial(link="logit"),data=aa,weights=aa$IPTW) 
summary(bb)
coefficients(bb)
confint(bb) 
exp(coefficients(bb) )
exp(confint(bb))

# Step 6：subgroup analysis 
aa <- aa[aa$TNF.naive == 1, ]

dataIPTW=svydesign(ids=~1,data=aa,weights= ~IPTW)

tab_IPTW=svyCreateTableOne(vars=vars, strata="group1",data=dataIPTW,test=T)
print(tab_IPTW,showAllLevels=TRUE,smd=TRUE)

bb<-glm(formula=clinical.remission26~group1,family=binomial(link="logit"),data=aa,weights=aa$IPTW) 
summary(bb)
coefficients(bb)
confint(bb) 
exp(coefficients(bb) )
exp(confint(bb))

