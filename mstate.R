library(openxlsx)
library(survival)
library(mstate)
library(dplyr)

setwd("F:\\10 SSS\\10\\data")

da2<-read.xlsx("data265794.xlsx",colNames = T)

da2$sumisoqf<-as.factor(da2$sumisoqf)
da2$sumloneqf<-as.factor(da2$sumloneqf)
da2$sumisof<-as.factor(da2$sumisof)
da2$sumlonef<-as.factor(da2$sumlonef)

#pattern A in model 1 as sample
#calculate loneliness and social isolation 
tmat <- mstate::transMat(x = list(c(2, 4), 
                                  c(3, 4), 
                                  c(4), 
                                  c()),
                         names = c("base", "CVD", "Dep", "Death"))
print(tmat)

msebmt <- msprep(data = da2, trans = tmat, 
                 time = c(NA, "cvd", "dep","death"), 
                 status = c(NA, "cvd.s", "dep.s", "death.s"), 
                 keep = c("sumisoq","sumloneq","sumisoqf","sumloneqf","scoial","visit","house","confide","lone","sumiso","sumlone","sumisof","sumlonef","joint","joint2","sex", "age", "BMI", "edu_qua", "income", "alcohol","smoke","ethic","MET","employ"))
msebmt$years<-msebmt$time/365.25
events(msebmt)
head(msebmt)

c<-c("N_case","N_control",'HR','HR_down','HR_up','Pvalue')
cc<-data.frame(c)
result3<-data.frame(cc)
colcc<-c(1)

b<-c("N_case","N_control",'Q2HR','Q2HR_down','Q2HR_up','Q2Pvalue','Q3HR','Q3HR_down','Q3HR_up','Q3Pvalue')
bb<-data.frame(b)
result2<-data.frame(bb)
colbb<-c(1)

j<-colnames(msebmt)
head(msebmt)
for(i in 9:23){
  for (k in 1:5){
    da = subset(msebmt, trans == k)
    da<-na.omit(da)
    feel<-da[,c(i)]
    fit1 <- coxph(Surv(years, status) ~ feel+sex+age+BMI+edu_qua+income+alcohol+smoke+ethic+MET+employ,
                  data = da)
    #fit1 <- coxph(Surv(years, status) ~ feel+sex+age+BMI+edu_qua+income+alcohol+smoke+ethic+MET+employ+hypertension+diabete,
    #                        data = da)
    if(nrow(summary(fit1)$coefficients)==11){
      N_case<-length(which(da[,c(8)]==1))
      N_control<-length(which(da[,c(8)]==0))
      aaa<-exp(confint(fit1,level=0.95))
      test<-c(N_case,N_control,exp(coef(fit1))[1],aaa[1,][c(1)],aaa[1,][c(2)],summary(fit1)$coefficients[1,5])
      name=paste(j[i],"*",k)
      colcc<-c(colcc,name)
      cc<-cbind(cc,test)
      colnames(cc)<-colcc}
    if(nrow(summary(fit1)$coefficients)==12){
      N_case<-length(which(da[,c(8)]==1))
      N_control<-length(which(da[,c(8)]==0))
      aaa<-exp(confint(fit1,level=0.95))
      test<-c(N_case,N_control,exp(coef(fit1))[1],aaa[1,][c(1)],aaa[1,][c(2)],summary(fit1)$coefficients[1,5],exp(coef(fit1))[2],aaa[2,][c(1)],aaa[2,][c(2)],summary(fit1)$coefficients[2,5])
      name=paste(j[i],"*",k)
      colbb<-c(colbb,name)
      bb<-cbind(bb,test)
      colnames(bb)<-colbb}
  }
}
result2<-cbind(result2,bb)
result3<-cbind(result3,cc)

write.xlsx(result2,"cvd_dep_soiso.xlsx")
write.xlsx(result3,"cvd_dep_loneliness.xlsx")

#pattern B in model 2 as sample
da2_3<-read.xlsx("datacvds_14t_262887.xlsx",colNames = T)

da2_3$sumisoqf<-as.factor(da2_3$sumisoq)
da2_3$sumloneqf<-as.factor(da2_3$sumloneq)

tmat <- mstate::transMat(x = list(c(2,3,4,5,7), 
                                  c(6,7), 
                                  c(6, 7),
                                  c(6, 7),
                                  c(6, 7),
                                  c(7),
                                  c()),
                         names = c("base", "ihd", "af", "hf", "stroke", "Dep", "Death"))
print(tmat)


msebmt1 <- msprep(data = da2_3, trans = tmat, 
                  time = c(NA, "ihdt","aft","hft","stroket","dep","death"), 
                  status = c(NA, "ihd.s","af.s","hf.s","stroke.s","dep.s", "death.s"), 
                  keep = c("eid","sumisoq","sumloneq","sumisoqf","sumloneqf","sex", "age", "BMI", "edu_qua", "income", "alcohol","smoke","ethic","MET","employ","hypertension","diabete"))

msebmt1$years<-msebmt1$time/365.25
events(msebmt1)
head(msebmt1)

a<-c("N_case","N_control",'Q2HR','Q2HR_down','Q2HR_up','Q2Pvalue','Q3HR','Q3HR_down','Q3HR_up','Q3Pvalue')
c<-data.frame(a)
result2<-data.frame(a)
colc<-c(1)

e<-c("N_case","N_control",'HR','HR_down','HR_up','Pvalue')
d<-data.frame(e)
result3<-data.frame(e)
cold<-c(1)

j<-colnames(msebmt1)
for(i in 10:13){
  for (k in 1:14){
    da = subset(msebmt1, trans == k)
    da<-na.omit(da)
    feel<-da[,c(i)]
    fit1 <- coxph(Surv(years, status) ~ feel+sex+age+BMI+edu_qua+income+alcohol+smoke+ethic+MET+employ+hypertension+diabete,
                  data = da)
    #fit1 <- coxph(Surv(years, status) ~ feel+sex+age+BMI+edu_qua+income+alcohol+smoke+ethic+MET+employ,
    #            data = da)
    if(nrow(summary(fit1)$coefficients)==14){
      N_case<-length(which(da[,c(8)]==1))
      N_control<-length(which(da[,c(8)]==0))
      aaa<-exp(confint(fit1,level=0.95))
      test<-c(N_case,N_control,exp(coef(fit1))[1],aaa[1,][c(1)],aaa[1,][c(2)],summary(fit1)$coefficients[1,5],exp(coef(fit1))[2],aaa[2,][c(1)],aaa[2,][c(2)],summary(fit1)$coefficients[2,5])
      name=paste(j[i],"*",k)
      colc<-c(colc,name)
      c<-cbind(c,test)
      colnames(c)<-colc}
    if(nrow(summary(fit1)$coefficients)==13){
      N_case<-length(which(da[,c(8)]==1))
      N_control<-length(which(da[,c(8)]==0))
      aaa<-exp(confint(fit1,level=0.95))
      test<-c(N_case,N_control,exp(coef(fit1))[1],aaa[1,][c(1)],aaa[1,][c(2)],summary(fit1)$coefficients[1,5])
      name=paste(j[i],"*",k)
      cold<-c(cold,name)
      d<-cbind(d,test)
      colnames(d)<-cold}
  }
}
result2<-cbind(result2,c)
result3<-cbind(result3,d)

write.xlsx(result2,"lone_fac_cvds_sen.xlsx")
write.xlsx(result3,"lone_num_cvds_sen.xlsx")

