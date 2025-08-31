library(ggplot2)
library(flextable)
library(officer)
library(dplyr)

myTable <- (diamonds) %>% group_by(cut) %>%
  summarise(Carat = mean(carat,na.rm = T),Price = mean(price,na.rm = T), Depth = mean(depth, na.rm =T)) %>%
  flextable() %>%
  mk_par(j = 'Price', value = as_paragraph(linerange(value = Price, max = max(Price)))) %>%
  mk_par(j = 'Carat', value = as_paragraph(linerange(value = Carat, max = max(Carat),stickcol = "blue"))) %>% set_table_properties(width = 1, layout = "autofit") %>%
  add_header_row(values=c("This a header row"),colwidths=4) %>%
  align(i=1,align="center",part="header") %>%
  footnote(i=4:5,j=1,
           value = as_paragraph(c("This is footer 1","This is footer 2" )),
           ref_symbols = c("1","2"),part="body"
  )

doc <- read_docx() %>% 
  body_add_par("Diamonds", style = "Table Caption") %>% 
  body_add_par("\n") %>%
  body_add_flextable(myTable) %>% print("Single-table.docx")