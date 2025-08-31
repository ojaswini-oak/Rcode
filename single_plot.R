library(ggplot2)
library(dplyr)
library(officer)


myPlot <- ggplot(diamonds, aes(price,carat,group = color))+
  geom_point(aes(colour = color))+ labs(title = "Carat and Price of Diamonds by Colour") + facet_grid(.~color) + 
  theme(axis.text.x = element_text(angle = 45,hjust = 0.7), plot.title= element_text(hjust= 0.5))

doc2 <- read_docx() %>%
  body_add_gg(myPlot,height= 4) %>%
  print("Plot-only.docx")
