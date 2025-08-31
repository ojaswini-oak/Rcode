
library(officer)


#Some placeholder text read from text.txt 

t <- system.file(package="officer",
                 "doc_examples","text.txt"
                 )
data <- readLines(t)

#sample ggplot


myplot <- ggplot(mtcars , aes(x=wt,y=mpg)) +
  geom_point(aes(colour= as.factor(gear)),shape = 19,size = 2) +
  geom_smooth(method = "lm",se=F,colour = "darkgrey") + 
  labs(x = "Weight",y = "Miles per Gallon",
       title = "Fuel Efficiency Decreases with Increase in Weight", colour = "Gear") + 
  theme(plot.title = element_text(hjust = 0.5))
ggsave("myplot.png", myplot, width = 6, height = 4)


centered_plot <- fpar(
  run_linebreak(),
  external_img("myplot.png",
               width=5.8,height=4.5),
  fp_p = fp_par(text.align = "center")
)




#Document
doc <- read_docx("MyTemplate.docx")

doc |>
  body_add_par(data[1],pos="before") |>
  body_add_fpar(centered_plot) |>
  body_add_break() |>
  body_add_par(data[2]) |>
  body_add_par(data[3]) |>
  print("Tempate-based-doc.docx")
  

