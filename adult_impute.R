# import library
library(tidyverse)
library(data.table)
library(epicalc)
library(openxlsx)
library(extrafont)
library(mice)

# import pre-test
X <- setDT(read.xlsx("D:/commed/pre_adult.xlsx"))
colnames(X)
colnames(X)[2] <- "SUM"
colnames(X)[25:29] <- c("Q20", "Q21", "Q22", "Q23", "Q24")
X1 <- X[,c(-1,-30)]
colnames(X1)
X1 <- X1[,-3]
colnames(X1)

# change variable pre-test
X1[,Q1:=ifelse(Q1=="ใช่",1,0)]
X1[,Q2:=ifelse(Q2=="ใช่",1,0)]
X1[,Q3:=ifelse(Q3=="ใช่",1,0)]
X1[,Q4:=ifelse(Q4=="ไม่ใช่",1,0)]
X1[,Q5:=ifelse(Q5=="ไม่ใช่",1,0)]
X1[,Q6:=ifelse(Q6=="ใช่",1,0)]
X1[,Q7:=ifelse(Q7=="ไม่ใช่",1,0)]
X1[,Q8:=ifelse(Q8=="ใช่",1,0)]
X1[,Q9:=ifelse(Q9=="ไม่ใช่",1,0)]
X1[,Q10:=ifelse(Q10=="ใช่",1,0)]
X1[,Q11:=ifelse(Q11=="ใช่",1,0)]
X1[,Q12:=ifelse(Q12=="ใช่",1,0)]
X1[,Q13:=ifelse(Q13=="ไม่ใช่",1,0)]
X1[,Q15:=ifelse(Q15=="ไม่ใช่",1,0)]
X1[,Q16:=ifelse(Q16=="ไม่ใช่",1,0)]
X1[,Q17:=ifelse(Q17=="ไม่ใช่",1,0)]
X1[,Q18:=ifelse(Q18=="ไม่ใช่",1,0)]
X1[,Q19:=ifelse(Q19=="ใช่",1,0)]
X1[,Q20:=ifelse(Q20=="ใช่",1,0)]
X1[,Q21:=ifelse(Q21=="ใช่",1,0)]
X1[,Q22:=ifelse(Q22=="ไม่ใช่",1,0)]
X1[,Q23:=ifelse(Q23=="ไม่ใช่",1,0)]
X1[,Q24:=ifelse(Q24=="ใช่",1,0)]

# import post-test
Y <- setDT(read.xlsx("D:/commed/post_adult.xlsx"))
colnames(Y)
colnames(Y)[2] <- "SUM"
colnames(Y)[25:29] <- c("Q20", "Q21", "Q22", "Q23", "Q24")
Y1 <- Y[,-1]
colnames(Y1)
Y1 <- Y1[,-3]
colnames(Y1)

# change variable post-test
Y1[,Q1:=ifelse(Q1=="ใช่",1,0)]
Y1[,Q2:=ifelse(Q2=="ใช่",1,0)]
Y1[,Q3:=ifelse(Q3=="ใช่",1,0)]
Y1[,Q4:=ifelse(Q4=="ไม่ใช่",1,0)]
Y1[,Q5:=ifelse(Q5=="ไม่ใช่",1,0)]
Y1[,Q6:=ifelse(Q6=="ใช่",1,0)]
Y1[,Q7:=ifelse(Q7=="ไม่ใช่",1,0)]
Y1[,Q8:=ifelse(Q8=="ใช่",1,0)]
Y1[,Q9:=ifelse(Q9=="ไม่ใช่",1,0)]
Y1[,Q10:=ifelse(Q10=="ใช่",1,0)]
Y1[,Q11:=ifelse(Q11=="ใช่",1,0)]
Y1[,Q12:=ifelse(Q12=="ใช่",1,0)]
Y1[,Q13:=ifelse(Q13=="ไม่ใช่",1,0)]
Y1[,Q15:=ifelse(Q15=="ไม่ใช่",1,0)]
Y1[,Q16:=ifelse(Q16=="ไม่ใช่",1,0)]
Y1[,Q17:=ifelse(Q17=="ไม่ใช่",1,0)]
Y1[,Q18:=ifelse(Q18=="ไม่ใช่",1,0)]
Y1[,Q19:=ifelse(Q19=="ใช่",1,0)]
Y1[,Q20:=ifelse(Q20=="ใช่",1,0)]
Y1[,Q21:=ifelse(Q21=="ใช่",1,0)]
Y1[,Q22:=ifelse(Q22=="ไม่ใช่",1,0)]
Y1[,Q23:=ifelse(Q23=="ไม่ใช่",1,0)]
Y1[,Q24:=ifelse(Q24=="ใช่",1,0)]

# join pre-test and post-test
Z <- X1 %>% full_join(Y1, by = c("ID"))
colnames(Z)
Z[,.N,.(GENDER.x,GENDER.y)]
Z <- Z[,-29]
Z <- Z[,-29]
colnames(Z)[3:4] <- c("GENDER", "STATUS")
colnames(Z)
col <- c("ID","GENDER", "STATUS", "Q1.x", "Q2.x", "Q3.x", "Q4.x", "Q5.x",
         "Q6.x", "Q7.x", "Q8.x", "Q9.x", "Q10.x", "Q11.x", "Q12.x", 
         "Q13.x", "Q15.x", "Q16.x", "Q17.x", "Q18.x", "Q19.x", "Q20.x", 
         "Q21.x", "Q22.x", "Q23.x", "Q24.x", "SUM.x", "Q1.y", "Q2.y", "Q3.y", 
         "Q4.y", "Q5.y",
         "Q6.y", "Q7.y", "Q8.y", "Q9.y", "Q10.y", "Q11.y", "Q12.y", 
         "Q13.y", "Q15.y", "Q16.y", "Q17.y", "Q18.y", "Q19.y", "Q20.y", 
         "Q21.y", "Q22.y", "Q23.y", "Q24.y", "SUM.y")
Z <- Z[,..col]
colnames(Z)

#multiple imputation
imp1 <- mice(Z, maxit = 0)
meth<-imp1$method
predM <- imp1$predictorMatrix

logis <- c("Q1.y", "Q2.y", "Q3.y", 
           "Q4.y", "Q5.y",
           "Q6.y", "Q7.y", "Q8.y", "Q9.y", "Q10.y", "Q11.y", "Q12.y", 
           "Q13.y", "Q15.y", "Q16.y", "Q17.y", "Q18.y", "Q19.y", "Q20.y", 
           "Q21.y", "Q22.y", "Q23.y", "Q24.y")
meth[logis] <- "logreg"
meth["SUM.y"] <- ""
imp2 <- mice(Z, m = 100, predictorMatrix = predM,  
             method = meth,  printFlag = FALSE)
Z3 <- complete(imp2)
Z3 <- setDT(Z3)

Z4 <- Z3 %>% dplyr::select(ID, SUM.x, SUM.y)
shapiro.test(Z4$SUM.x)
shapiro.test(Z4$SUM.y)

wilcox.test(Z4$SUM.x, Z4$SUM.y, paired = T)

Z5 <- Z4 %>% dplyr::select(ID, SUM.x)
Z5[,TEST:=list("Pre", "Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre",
               "Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre",
               "Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre", 
               "Pre","Pre","Pre","Pre","Pre","Pre","Pre","Pre")]
colnames(Z5)[2] <- "SUM"

Z6 <- Z4 %>% dplyr::select(ID, SUM.y)
Z6[,TEST:=c("Post", "Post", "Post","Post","Post","Post","Post",
            "Post","Post","Post","Post","Post","Post","Post",
            "Post","Post","Post","Post","Post","Post","Post",
            "Post","Post","Post","Post","Post","Post","Post",
            "Post", "Post", "Post","Post","Post","Post","Post",
            "Post","Post","Post")]
l <- list(Z5, Z6)
B <- rbindlist(l)
Z4[,GAIN:=ifelse(SUM.x<SUM.y,"gain",ifelse(SUM.x>SUM.y,"loss","non"))]

Z7 <- Z4 %>% dplyr::select(ID, GAIN)
B1 <- B %>% full_join(Z7, by = "ID")
B1$TEST <- factor(B1$TEST, levels = c("Pre", "Post"))
B1 <- B1 %>% mutate(colors = case_when(GAIN == "gain" ~ "#006400", 
                                     GAIN == "loss" ~ "#FF0000", TRUE ~ "gray"))

# boxplot
p <- ggplot(B1, aes(as.factor(TEST), as.numeric(SUM), fill = TEST)) +
  geom_boxplot() +
  geom_line(aes(group = ID, color = B1$colors), size=1, alpha=0.6) +
  geom_point(aes(fill = TEST ,group = ID),size=5,shape=21) +
  scale_color_manual(values = c("#006400", "#FF0000", "grey"), 
                     name = "การเปลี่ยนแปลง", labels = c("เพิ่มขึ้น", "ลดลง", "เท่าเดิม")) +
  theme_classic() +
  labs(x = "การทดสอบ", y ="คะแนน"
       , title = "แผนภาพกล่องคะแนน Pre-test และ Post-test ของผู้มีส่วนได้เสีย") +
  theme(plot.title=element_text(size=20, hjust = -0.7, family = "TH Sarabun New", 
                                face = "bold"), 
        axis.title=element_text(size=18, family = "TH Sarabun New", 
                                face = "bold"), 
        legend.title=element_text(size=18, family = "TH Sarabun New", 
                                  face = "bold"), 
        legend.text=element_text(size=16, family = "TH Sarabun New"),
        axis.text.x=element_text(size=16, family = "TH Sarabun New", 
                                 color = "black", face = "bold"),
        axis.text.y=element_text(size=16, family = "TH Sarabun New", 
                                 color = "black")
  ) +
  scale_x_discrete(labels = c("Pre-test", "Post-test")) +
  scale_fill_manual(name = "การทดสอบ", 
                    values=c("#90EE90", "#87CEEB"), 
                    labels = c("Pre-test", "Post-test")) +
  ylim(0,23)
ggsave("boxplot_adult_impute.png", p, path = "D:/commed/กราฟค่ะกราฟ",height  = 8)

# NG
Z4[,NG:=(SUM.y-SUM.x)/(23-SUM.x)]
Z4[,NG2:=ifelse(NG<=0,"Non",ifelse(NG<0.3,"Low",ifelse(NG<0.7,"Medium","High")))]
Z4$NG2 <- factor(Z4$NG2, levels = c("High", "Medium", "Low"))
Z4 <- Z4 %>% filter(NG > 0)
Z4 <- Z4 %>% mutate(PASS=ifelse(SUM.y >= 0.7*23, "pass", "fail"))
Z4$PASS <- factor(Z4$PASS, levels = c("pass", "fail"))
p <- ggplot(Z4, aes(reorder(ID, -NG), NG, fill = PASS)) +
  geom_col() +
  geom_hline(yintercept = 0.7, color = "red") +
  geom_hline(yintercept = 0.3, color = "red") +
  labs(x = "คนที่", y = "normalized gain"
       , title = "กราฟแท่งแสดง normalized gain และการสอบผ่านของผู้มีส่วนได้เสีย") +
  geom_text(aes(label=round(NG, 2)), hjust = -0.8,
            family="TH Sarabun New", size = 5, angle = 90) +
  ylim(0,1) +
  theme_classic() +
  theme(plot.title=element_text(size=20, hjust = -0.5, family = "TH Sarabun New", 
                                face = "bold"), 
        axis.title=element_text(size=16, family = "TH Sarabun New", 
                                face = "bold"), 
        legend.title=element_text(size=14, family = "TH Sarabun New", 
                                  face = "bold"), 
        legend.text=element_text(size=14, family = "TH Sarabun New"),
        axis.text.x=element_text(size=12, family = "TH Sarabun New", 
                                 color = "black"),
        axis.text.y=element_text(size=12, family = "TH Sarabun New", 
                                 color = "black")
  ) +
  scale_x_discrete(labels=c(1,2,3,4,5,6,7,8,9,10,11,12,13,
                            14,15,16,17,18,19,20,21,22,23,24,25,
                            26,27,28,29,30,31,32,33)) +
  scale_fill_manual(name = "การสอบผ่าน 70%", 
                    values=c("#00008B", "#87CEEB")) +
  annotate(geom="text", y = 0.75, x = 38, label = "High gain\ncutpoint", 
           family = "TH Sarabun New", size = 5, 
           color = "red", fontface = "bold") +
  annotate(geom="text", y = 0.25, x = 38, label = "Medium gain\ncutpoint", 
           family = "TH Sarabun New", size = 5, 
           color = "red", fontface = "bold") +
  coord_cartesian(xlim=c(1,33),clip="off")
ggsave("NG_adult_pass_impute.png", p, path = "D:/commed/กราฟค่ะกราฟ",height  = 4.5)
       