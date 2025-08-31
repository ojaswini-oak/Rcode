library(flextable)
library(dplyr)
library(ggplot2) #For the diamonds Dataset
library(officer)


#FLEXTABLES-------------------------------------------------------------
df <- iris %>%
  group_by(Species) %>%
  summarize(Sepal.Length = mean(Sepal.Length), Sepal.Width = mean(Sepal.Width),
            Petal.Length = mean(Petal.Length), Petal.Width = mean(Petal.Width)) 

tbl1 <- flextable(df) %>%
  color(i  = ~ Petal.Length == max(df$Petal.Length),  
        j = 'Petal.Length',color = "red") %>%
  color(i  = ~ Sepal.Length == max(df$Sepal.Length),  
        j = 'Sepal.Length',color = "red") %>%
  color(i  = ~ Petal.Width == max(df$Petal.Width),  
        j = 'Petal.Width',color = "red") %>%
  color(i  = ~ Sepal.Width == max(df$Sepal.Width),  
        j = 'Sepal.Width',color = "red") %>%
  separate_header() %>%
  theme_zebra(odd_body = "#dcf7ed",odd_header = "#F7DCE7",even_header = "#F7DCE7") %>%
  valign(part = "header",valign = "center") %>%
  align(part = "all",align = "center") %>%
  border_inner(border = fp_border_default()) %>%
  border_outer(border = fp_border_default()) %>%
  set_table_properties(width = 0.8,layout="autofit")


tbl2 <- (diamonds) %>%
  group_by(cut) %>%
  summarise(Carat = mean(carat),Price = mean(price,na.rm = T), Depth = round(mean(depth),4))

tbl2 <- tbl2 %>%
  flextable() %>%
  mk_par(j = 'Price', value = as_paragraph(
    linerange(value = Price, max = max(Price)) )) %>%
  mk_par(j = 'Carat', value = as_paragraph(
    linerange(value = Carat, max = max(Carat),stickcol = "blue"))) %>%
  set_header_labels(values = c(cut = "Cut",Carat = "Mean Carat Size",Price = "Mean Price",Depth = "Depth")) %>%
  set_table_properties(width = 0.8,layout="autofit") %>%
  width(width = 1) %>% align(part = "all",align = "center") %>%
  align(j = "cut",part = "all",align = "left")




#OFFICER-----------------------------------------------------------


page_properties <- prop_section(
  page_size = page_size(orient = "landscape")
)

doc <- read_docx() %>%
  body_add_par("Iris",style="heading 1",pos="before") %>% body_add_par("\n")%>%
  body_add_flextable(tbl1) %>% 
  body_add_par("\n")%>% body_add_par("\n")%>% 
  body_add_par("Diamonds",style="heading 1") %>% body_add_par("\n")%>%
  body_add_flextable(tbl2) %>% 
  body_set_default_section(page_properties)

print(doc,"multiple_tables.docx")