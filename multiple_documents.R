library(readxl)
library(flextable)
library(tidyverse)
library(dplyr)
library(officer)
library(ggplot2)


#Reading data from excel
df <- as.data.frame(read_excel("feedback.xlsx",sheet=1))
qb <- as.data.frame(read_excel("feedback.xlsx",sheet=2))

#removing extra row for person no.
rownames(df) <- df[,1]
df[,1] <- NULL



#summary table
s_tbl <- as.data.frame(matrix(nrow=ncol(df) - 3,ncol=11))
colnames(s_tbl) <- c("n","Mean","SD","Median","Min","Max","","Response of 4 [n(%)]","Response of 3 [n(%)]","Response of 2 [n(%)]","Response of 1 [n(%)]")

#for ggplot
pl_tbl <- as.data.frame(matrix(ncol = nrow(qb) + 1 - 3, nrow = 4))

pl_tbl[,1] <- c("a","b","c","d")

v<- vector()
v[1] <- "key"
for(i in seq(1,ncol(pl_tbl) - 1)){
  v[i+1] <- paste0("Q",i)
}

colnames(pl_tbl) <- v

pl_tbl


pl_tbl_2 <- as.data.frame(matrix(ncol = ncol(df), nrow = 5))

pl_tbl_2[,1] <- c("a","b","c","d","e")

v2<- vector()
v2[1] <- "opt"
n=1
for(i in seq(11,13)){
  v2[n+1] <- paste0("Q",i)
  n= n+1
}

colnames(pl_tbl_2) <- v2

pl_tbl_2



s_tbl_2 <- as.data.frame(matrix(nrow=ncol(df),ncol=12))
colnames(s_tbl_2) <- c("n","Mean","SD","Median","Min","Max","","Response of 5 [n(%)]","Response of 4 [n(%)]","Response of 3 [n(%)]","Response of 2 [n(%)]","Response of 1 [n(%)]")


#Function for Calculating statistics
calc_stat <- function(df,x) {
  
  
  n = nrow(df)
  avg = round(mean(unlist(df[,x])),2)
  s = round(sd(df[,x]),2)
  med = round(median(df[,x]),2)
  mn = round(min(df[,x]),2)
  mx = round(max(df[,x]),2)
  
  
  occ_4 <- length(which(as.data.frame(df[,x]) == 4))
  occ_3 <- length(which(as.data.frame(df[,x]) == 3))
  occ_2 <- length(which(as.data.frame(df[,x]) == 2))
  occ_1 <- length(which(as.data.frame(df[,x]) == 1))
  
  
  per_4<- round(((occ_4 / as.double(n)) * 100) , 2)
  per_3<- round(((occ_3 / as.double(n)) * 100) , 2)
  per_2<- round(((occ_2 / as.double(n)) * 100) , 2)
  per_1<- round(((occ_1 / as.double(n)) * 100) , 2)
  
  p_4 <- paste0(occ_4," (",per_4,"%)")
  p_3 <- paste0(occ_3," (",per_3,"%)")
  p_2 <- paste0(occ_2," (",per_2,"%)")
  p_1 <- paste0(occ_1," (",per_1,"%)")
  
  
  if(x<11){
    pl_tbl[,x+1] <<- c(per_4,per_3,per_2,per_1) 
    
    s_tbl[x,] <<- c(n,avg,s,med,mn,mx,"",p_4,p_3,p_2,p_1)}
  
  if(x>10) {
    occ_5 <- length(which(as.data.frame(df[,x]) == 5))
    per_5 <- round(((occ_5 / as.double(n)) * 100) , 2)
    p_5 <- paste0(occ_5," (",per_5,"%)")
    
    pl_tbl_2[,x+1] <<- c(per_5,per_4,per_3,per_2,per_1) 
    
    s_tbl_2[x,] <<- c(n,avg,s,med,mn,mx,"",p_5,p_4,p_3,p_2,p_1)
    
  }
  
}



#For loop to calculate stats of all the columns(questions)
for (i in seq(1,ncol(df))) {
  calc_stat(df,i)
}

final <- list()

ques_stats <- function(x) {
  
  if(x < 11) {
    q <- c(paste0("\n4 = ",qb[x,2],"\n\n3 = ",qb[x,3],"\n\n2 = ",qb[x,4],"\n\n1 = ",qb[x,5]),rep(0,10))
    st <- as.vector(unlist(colnames(s_tbl)))
    va <- as.vector(unlist(s_tbl[x,]))
    
    
    final[[x]] <<- as.data.frame(cbind(q,st,va))
    colnames(final[[x]]) <<- c("Key","Statistic","Value") 
  }
  
  if(x > 10) {
    
    q <- c(paste0("1 = ",qb[x,2],"\nTo\n5 = ",qb[x,3]),rep(0,11))
    st <- as.vector(unlist(colnames(s_tbl_2)))
    va <- as.vector(unlist(s_tbl_2[x,]))
    
    
    final[[x]] <<- as.data.frame(cbind(q,st,va))
    colnames(final[[x]]) <<- c("Scale","Statistic","Value") 
    
  }
} 

for(i in seq(1,nrow(qb))) {
  
  ques_stats(i)
}


pl_tbl_2 <- pl_tbl_2[,c(1,12:14)]



#making table

tabs <- function(x) {
  
  s <- paste0("f",x)
  n <- as.name(s)
  print(s)
  assign(s,final[[x]],envir = .GlobalEnv)
  
  return(n)
}


z <- read_docx()
z <- body_add_par(z,"FEEDBACK REPORT\n",style = "heading 1",pos='before')



flexx <- function(x,n) {
  question_no <- n
  
  x<- flextable(x)
  
  set_flextable_defaults(font.size = 10, theme_fun = theme_box(x))
  x<- border_inner(x,border=fp_border(style="solid"),part="all")
  x<- border_outer(x,border=fp_border(style="solid"),part="all")
  x<- width(x,j=1,width=6,unit = "cm")
  x<- width(x,j=2,width=4,unit = "cm")
  x<- width(x,j=3,width=4,unit = "cm")
  x<- align(x,align="center",part = "all")
  x <- align(x,j=1,align="left")
  x<- padding(x, padding.left = 3, part="all")
  
  
  if(n<11){
    x<- merge_at(x,i=1:11,j=1)
    q <- paste0("Q",question_no," ",qb[n,1])
    
    test <- pl_tbl
    
    test$labs <- test[,n+1]
    
    test <- test %>%
      arrange(labs) %>%
      mutate(pos = cumsum(labs) - 0.5*labs)
    
    test$labs <- test[,n+1]
    
    g <- ggplot(pl_tbl, aes(x="",y=pl_tbl[,n+1], fill = key)) +
      geom_bar(stat="identity",position = "stack") +
      theme_void() +
      coord_polar("y", start=0) + labs(fill="Key") + 
      geom_text(data = subset(test,labs!=0),aes(y = pos,label= paste0(labs,"%"))) +
      scale_fill_manual(labels = c("4", "3", "2","1"),values = c("#40D60B","#e8f347","#FFA023","#ff0017"))

  }
  if(n > 10) {
    x<- merge_at(x,i=1:12,j=1)
    q <- paste0("Q",question_no," ",qb[n,1])
    n = n - 10
    
    test <- pl_tbl_2
    
    test$labs <- test[,n+1]
    
    test <- test %>%
      arrange(labs) %>%
      mutate(pos = cumsum(labs) - 0.5*labs)
    
    test$labs <- test[,n+1]
    
    g <- ggplot(pl_tbl_2, aes(x="",y=pl_tbl_2[,n+1], fill = opt)) +
      geom_bar(stat="identity",position = "stack") +
      theme_void() +
      coord_polar("y", start=0) + labs(fill="Key") + 
      scale_fill_manual(labels = c("5","4", "3", "2","1"), values = c("#40D60B","#e8f347","#ff9f00","#ff4600","#b30017")) +
      geom_text(data = subset(test,labs!=0),aes(y = pos,label= paste0(labs,"%"))) 
  }
  
  # body_add_break(z)
  
  title_str <- paste("Question",question_no)
  fname <- paste0("Question",question_no,".docx")
  
  z <- read_docx()
  z <- body_add_par(z," ")
  z <- body_add_par(z,title_str,style = "table title")
  z <- body_add_par(z,"\n\n")
  z <- body_add_par(z,q,style="centered",pos='after')
  z <- body_add_par(z,"\n\n")
  z <- body_add_flextable(z,x,pos='after')
  z <- body_add_par(z,"\n",pos='after')
  z <- body_add_gg(z,g,width=4,height=4,pos='after')
  z <- body_add_par(z,"\n",pos='after')
  cat(fname," is being printed")
  print(z,fname)
}



for(i in seq(1,13)){
  flexx(final[[i]],i)
  tabs(i)
}
