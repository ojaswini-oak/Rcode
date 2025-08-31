library(readxl)
library(dplyr)
library(tidyr)
library(psych)
library(officer)
library(flextable)

df <- read_excel("data1.xlsx")[,c("USUBJID","TRT01P","TRT01PN","AGE","AGEGR1","SEX", "RACE","RACEN","ETHNIC","MMSETOT","DURDIS","DURDSGR1","EDUCLVL","WEIGHTBL","HEIGHTBL","BMIBL","BMIBLGR1")]


#Function for df------------------------------------

make_df <- function(tbl1, name_of_table){
  
  rownames(tbl1) <- tbl1[,"group1"]
  
  tbl2 <- as.data.frame(t(tbl1[,c("n","mean","sd","median","min","max")]))
  
  
  tbl2["n",] <- round(tbl2["n",],0)
  tbl2[c("mean","median","min","max"),] <- round(tbl2[c("mean","median","min","max"),],1)
  tbl2["sd",] <- round(tbl2["sd",],2)
  
  tbl2$extra <- c(name_of_table,rep("",nrow(tbl2) - 1))
  
  tbl2$names <- c("n","Mean","SD","Median","Min","Max")
  
  tbl2 <- tbl2[,c("extra","names","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]
  
  colnames(tbl2) <- c("extra","names","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")
  
  tbl2[nrow(tbl2) + 1,] <-  rep("",ncol(tbl2))
  tbl2[nrow(tbl2) + 1,] <-  rep("",ncol(tbl2))
  
  return(tbl2)
}

#creating doc----------------------------------------------
doc <- read_docx('demo-tmplt.docx')
dim <- docx_dim(doc)

w <- dim$page['width'] - dim$margins['left'] - dim$margins['right']
col_width <- c(w*0.2,w*0.2,w*0.15,w*0.15,w*0.15,w*0.15)


make_flex <- function(tbl) {
  f <- flextable(tbl) |>
    width(j=1:6,width=col_width) |>
    set_table_properties(width=1) |>
    delete_part(part="header") |>
    hline_bottom(part="body",border = fp_border_default(width = 0)) |>
    height_all(0.2, part = "all", unit = "in") |>
    hrule(rule="exact")
  return(f)
}



#AGE----------------------------------------

df2 <- df

df2[,"TRT01P"] <- "total"

df3 <- rbind(df,df2)

age <- describeBy(AGE~TRT01P,data = df3[df3$TRT01PN %in% c(0,81,54),c("AGE","TRT01P")],mat = T) %>%
  select(n,mean,sd,median,min,max)

age$n <- round(age$n, 0)
age$mean <- round(age$mean, 1)
age$sd <- round(age$sd, 2)
age$median <- round(age$median, 1)
age$min <- round(age$min, 1)
age$max <- round(age$max, 1)


age <- t(age)[,c("AGE1","AGE4","AGE3","AGE2")]

age <- as.data.frame(age)

colnames(age) <- c("pl","xl","xh","total")

age$name <- rownames(age)
age$value <- c("AGE",rep("",nrow(age)-1))

age <- age[,c("value","name","pl","xl","xh","total")]


age_2 <- df3[,c("TRT01P", "AGEGR1")] %>%
  group_by(TRT01P,AGEGR1) %>% count() 


age_gr <- age_2 %>%
  pivot_wider(names_from = TRT01P,
              values_from = n)

age_gr <- as.data.frame(age_gr)

age_gr$AGEGR1 <- paste(age_gr$AGEGR1, "yrs")

rownames(age_gr) <- age_gr[,"AGEGR1"]


age_gr <- age_gr[c("<65 yrs","65-80 yrs",">80 yrs"),c("Placebo","Xanomeline Low Dose","Xanomeline High Dose", "total")]
colnames(age_gr) <- c("pl","xl","xh","total")


make_percent <- function(df_att,tot_val) {
  
  ifelse(class(df_att) != "numeric", df_att <- as.numeric(df_att), df_att <- df_att)
  ifelse(class(tot_val) != "numeric", tot_val <- as.numeric(tot_val), tot_val <- tot_val)
  
  ifelse(round((df_att/tot_val)*100, 0) == 0,
         paste0("  ",df_att," ( <1%)"),
         ifelse(nchar(round((df_att/tot_val)*100, 0)) > 1,
                ifelse(nchar(df_att) == 3,
                       paste0(df_att," ( ",round((df_att/tot_val)*100, 0),"%)"),
                       ifelse(nchar(df_att) ==  2, paste0(" ",df_att," ( ",round((df_att/tot_val)*100, 0),"%)"),
                              paste0("  ",df_att," ( ",round((df_att/tot_val)*100, 0),"%)"))),
                paste0("  ",df_att," (  ",round((df_att/tot_val)*100, 0),"%)"))
  )
}


age_gr$pl <- make_percent(age_gr$pl,age["n","pl"])
age_gr$xl <- make_percent(age_gr$xl, age["n","xl"])
age_gr$xh <- make_percent(age_gr$xh, age["n","xh"])
age_gr$total <- make_percent(age_gr$total, age["n","total"])


age_gr$grp <- rownames(age_gr)
age_gr$extra <- ""

age_gr <- age_gr[,c("extra","grp","pl","xl","xh","total")]


age[nrow(age)+1,] <- rep("",ncol(age))

age_gr[nrow(age_gr)+1,] <- rep("",ncol(age_gr))
age_gr[nrow(age_gr)+1,] <- rep("",ncol(age_gr))

#flextables for age-----------------------------------------
age_gr_tbl <- set_flextable_defaults(font.family="Courier New",font.size=9)

age_gr_tbl <- make_flex(age_gr)

age_tbl <-  flextable(age) |>
  width(j=1:6,width=col_width) |>
  set_table_properties(width=1) |>
  delete_part(part="header") |>
  hline_bottom(part="body",border = fp_border_default(width = 0)) |>
  add_header_row(values=c("","","Placebo\n(N=84)","Treatment 1\n(N=86)","Treatment 2\n(N=86)","Total\n(N=86)")) |>
  hline_top(part = "body",border=fp_border_default(width=2)) |>
  hline_top(part = "header",border=fp_border_default(width=2)) |>
  height_all(0.2, part = "all", unit = "in") |>
  height(height = 0.1, part = "header", unit = "in") |>
  hrule(rule="exact") |>
  fontsize(part="header",size=10)


#SEX-----------------------------------------------

sex <- df3 %>%
  select(TRT01P, SEX) %>%
  group_by(TRT01P, SEX) %>%
  tally() 

sex_2 <- sex %>%
  pivot_wider(names_from = TRT01P,
              values_from = n)

sex_2 <- as.data.frame(sex_2)


sex_2["n","SEX"] <- "n"
sex_2["n",2:5] <- sex_2[1,2:5] + sex_2[2,2:5] 

rownames(sex_2) <- sex_2$SEX
sex_2 <- sex_2[,-1]

sex_2<- sex_2[c("n","M","F"),c("Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]


sex_2["M",] <- make_percent(sex_2["M",],sex_2["n",])
sex_2["F",] <- make_percent(as.numeric(sex_2["F",]),as.numeric(sex_2["n",]))

sex_2$Extra <- c("SEX",rep("",nrow(sex_2)-1))

sex_2$Extra2 <- c("n","Male","Female")

sex_final <- sex_2[,c("Extra", "Extra2" ,"Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]


sex_final[nrow(sex_final) + 1,] <-  rep("",ncol(sex_final))
sex_final[nrow(sex_final) + 1,] <-  rep("",ncol(sex_final))

#flextable for sex--------------------------------------------
sex_tbl <- make_flex(sex_final)

#Race-------------------------------------
head(df)

race <- df3[,c("USUBJID","TRT01P","RACE","RACEN","ETHNIC")]

race <- race%>%
  group_by(TRT01P,RACE,ETHNIC) %>%
  count() %>%
  pivot_wider(names_from = TRT01P,
              values_from = n) %>%
  mutate_at(vars(-c(RACE,ETHNIC)),~replace(.,is.na(.),0))


race <- race[,c("Placebo", "Xanomeline Low Dose", "Xanomeline High Dose", "total")]



race <- as.data.frame(race)

race["n",] <- c(sum(race[,1]), sum(race[,2]), sum(race[,3]), sum(race[,4])) 

race[1,] <- make_percent(race[1,],race["n",]) 
race[2,] <- make_percent(race[2,],race["n",]) 
race[3,] <- make_percent(race[3,],race["n",]) 
race[4,3:4] <- make_percent(race[4,3:4],race["n",3:4]) 

race$RACE = c("African Descent", "Hispanic", "Caucasian", "Other", "n")

rownames(race) <- race$RACE
race$extra <- ""

race <- race[c("n","Caucasian","African Descent","Hispanic","Other"),c("extra","RACE","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]

race$extra <- c("RACE(Origin)",rep("",nrow(race) - 1))
race[nrow(race) + 1,] <-  rep("",ncol(race))
race[nrow(race) + 1,] <-  rep("",ncol(race))
#flextable for race------------------------------------
race_tbl <- make_flex(race)

#MMSE----------------------------------
mmse <- describeBy(MMSETOT~TRT01P, data = df3[df3$TRT01PN %in% c(0,54,81), c("MMSETOT", "TRT01P")],mat = T)

mmse <- mmse[,c("group1","n","mean","sd","median","min","max")]

rownames(mmse) <- mmse[,"group1"]
mmse <- mmse[,-1]
mmse <- t(mmse)[,c("Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]


mmse <- as.data.frame(mmse)

mmse["n",] <- round(mmse["n",],0)
mmse["mean",] <- round(mmse["mean",],1)
mmse["sd",] <- round(mmse["sd",],2)
mmse["median",] <- round(mmse["median",],1)
mmse["min",] <- round(mmse["min",],1)
mmse["max",] <- round(mmse["max",],1)


mmse$extra <- c("MMSE",rep("",nrow(mmse)-1))

mmse$names <- c("n","Mean","SD","Median","Min","Max")



mmse <- mmse[,c("extra","names","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]


mmse[nrow(mmse) + 1,] <-  rep("",ncol(mmse))

mmse[nrow(mmse) + 1,] <-  rep("",ncol(mmse))

#flextable for mmse-----------------------------------------
mmse_tbl <- make_flex(mmse)

#Duration Of Disease-------------------


dod<- describeBy(DURDIS~TRT01P, data = df3[df3$TRT01PN %in% c(0,54,81),c("DURDIS","TRT01P")],mat =T)

rownames(dod) <- dod[,"group1"]

dod <- as.data.frame(t(dod[,c("n","mean","sd","median","min","max")]))


dod["n",] <- round(dod["n",],0)
dod[c("mean","median","min","max"),] <- round(dod[c("mean","median","min","max"),],1)
dod["sd",] <- round(dod["sd",],2)

dod$extra <- c("Duration of Disease",rep("",nrow(dod) - 1))
dod$names <- c("n","Mean","SD","Median","Min","Max")

dod <- dod[,c("extra","names","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]


dod[nrow(dod) + 1,] <-  rep("",ncol(dod))

#DURATION OF DISEASE SECOND TABLE
dod_grp <- df3 %>%
  group_by(TRT01P, DURDSGR1) %>%
  count() %>%
  pivot_wider(values_from = n, names_from = TRT01P) 

dod_grp <- as.data.frame(dod_grp)

dod_grp$Placebo <- make_percent(dod_grp$Placebo, dod["n","Placebo"])
dod_grp$`Xanomeline Low Dose` <- make_percent(dod_grp$`Xanomeline Low Dose`, dod["n","Xanomeline Low Dose"])
dod_grp$`Xanomeline High Dose` <- make_percent(dod_grp$`Xanomeline High Dose`, dod["n","Xanomeline High Dose"])
dod_grp$total <- make_percent(dod_grp$total, dod["n","total"])


dod_grp$DURDSGR1 <- paste(dod_grp$DURDSGR1, "months")

dod_grp$extra <- rep("",nrow(dod_grp))

dod_grp <- dod_grp[,c("extra","DURDSGR1","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]


dod_grp[nrow(dod_grp) + 1,] <-  rep("",ncol(dod_grp))
dod_grp[nrow(dod_grp) + 1,] <-  rep("",ncol(dod_grp))

#flextables for dod------------------------------------

dod_tbl <- make_flex(dod)
dod_grp_tbl <- make_flex(dod_grp)

#Years of Education-----------------------------------------

yoe <- describeBy(EDUCLVL~TRT01P, data = df3[df3$TRT01PN %in% c(0,54,81),c("EDUCLVL","TRT01P")],mat =T)


yoe <- make_df(yoe, "Years of Education")

#flextable for years of education---------------------------
yoe_tbl <- make_flex(yoe)

#Baseline Weight-----------------------------

baseline_weight <- describeBy(WEIGHTBL~TRT01P, data = df3[df3$TRT01PN %in% c(0,54,81), c("WEIGHTBL", "TRT01P")], mat = T)

baseline_weight <- make_df(baseline_weight, "Baseline weight(kg)")

#flextable for baseline weight--------------------------
baseline_weight_tbl <- make_flex(baseline_weight)

#Baseline Height-----------------------------


baseline_height <- describeBy(HEIGHTBL~TRT01P, data = df3[df3$TRT01PN %in% c(0,54,81), c("HEIGHTBL","TRT01P")], mat = T)


baseline_height <- make_df(baseline_height, "Baseline height(cm)")

#flextable for baseline height--------------------------
baseline_height_tbl <- make_flex(baseline_height)

#Baseline BMI-------------------------------

baseline_bmi <- describeBy(BMIBL~TRT01P, data = df3[df3$TRT01PN %in% c(0,54,81), c("BMIBL", "TRT01P")], mat = T)

#BASELINE BMI SECOND TABLE
bmi_grp <- df3 %>%
  group_by(TRT01P, BMIBLGR1) %>%
  count() %>%
  pivot_wider(values_from = n, names_from = TRT01P) 

bmi_grp <- as.data.frame(bmi_grp)

bmi_grp$Placebo <- make_percent(bmi_grp$Placebo, baseline_bmi["BMIBL1","n"])
bmi_grp$`Xanomeline Low Dose` <- make_percent(bmi_grp$`Xanomeline Low Dose`, baseline_bmi["BMIBL4","n"])
bmi_grp$`Xanomeline High Dose` <- make_percent(bmi_grp$`Xanomeline High Dose`, baseline_bmi["BMIBL3","n"])
bmi_grp$total <- make_percent(bmi_grp$total, baseline_bmi["BMIBL2","n"])


bmi_grp$extra <- rep("",nrow(bmi_grp))

rownames(bmi_grp) <- bmi_grp$BMIBLGR1

str(bmi_grp)
bmi_grp <- bmi_grp[c("<25","25-<30",">=30"),c("extra","BMIBLGR1","Placebo","Xanomeline Low Dose","Xanomeline High Dose","total")]

#flextables for bmi-------------------------------------

baseline_bmi <- make_df(baseline_bmi, "Baseline BMI")

baseline_bmi_tbl <- make_flex(baseline_bmi)

bmi_grp_tbl <- make_flex(bmi_grp)

#---------------------

title <- fpar(
  ftext("Table 14-2.01",prop = fp_text(font.size=12,font.family="Courier New",italic=TRUE,bold=TRUE)),
  run_linebreak(),
  ftext("Summary of Demographic and Baseline Characteristics",prop = fp_text(italic=TRUE,bold=TRUE,font.size=12,font.family="Courier New")),
  fp_p = fp_par(text.align = "center")
  )


doc |> 
  body_add_fpar(title,pos="before") |>
  body_add_flextable(age_tbl) |>
  body_add_flextable(age_gr_tbl) |>
  body_add_flextable(sex_tbl)|>
  body_add_flextable(race_tbl) |>
  body_add_flextable(mmse_tbl) |>
  body_add_flextable(dod_tbl) |>
  body_add_flextable(dod_grp_tbl) |> 
  body_add_flextable(yoe_tbl) |>
  body_add_flextable(baseline_weight_tbl) |>
  body_add_flextable(baseline_height_tbl) |>
  body_add_flextable(baseline_bmi_tbl) |>
  print("demography.docx")


