
library(officer)
library(dplyr)

doc<- read_docx()

#Page Dimensions
page_dim <- docx_dim(doc) 
full_width <- page_dim$page["width"] - (page_dim$margins["left"] + page_dim$margins["right"])
full_height <- page_dim$page["height"] - (page_dim$margins["left"] + page_dim$margins["right"])

#ggplot 
myplot <- ggplot(mtcars , aes(x=wt,y=mpg)) +
  geom_point(aes(colour= as.factor(gear)),shape = 19,size = 2) +
  geom_smooth(method = "lm",se=F,colour = "darkgrey") + 
  labs(x = "Weight",y = "Miles per Gallon",
       title = "Fuel Efficiency Decreases with Increase in Weight", colour = "Gear") + 
  theme(plot.title = element_text(hjust = 0.5))
ggsave("myplot.png", myplot, width = 6, height = 4)

#centering ggplot
centered_plot <- fpar(
  run_linebreak(),
  external_img("myplot.png",
               width=6,height=4.8),
  fp_p = fp_par(text.align = "center")
)

#custom layout settings
landscape <- prop_section(page_size = page_size(orient = "landscape"))
portrait <- prop_section(page_size = page_size(orient = "portrait"))

#adding all the elements into doc
doc <- doc %>%
  body_add_par("This is a page with Portrait Layout",style="heading 1") %>%
  body_add_fpar(centered_plot) %>% body_add_par("\n") %>%
  body_end_block_section(value = block_section(portrait)) %>%
  body_add_par("This is a page with Landscape Layout",style="heading 1") %>%
  body_add_par("\n") %>%
  body_add_fpar(centered_plot) %>% body_add_par("\n") %>%
  body_end_block_section(value = block_section(landscape)) %>%
  body_add_par("This is a page with Portrait Layout",style="heading 1") %>%
  body_add_fpar(centered_plot) %>%
  body_end_block_section(value = block_section(portrait)) %>%
  body_set_default_section(portrait)

print(doc, target = "custom_layout.docx")
