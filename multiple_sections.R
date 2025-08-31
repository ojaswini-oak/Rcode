library(officer,ggplot2,dplyr)


#ggplots-----------------------------------------------------------
df <- mtcars

df$gear <- as.factor(df$gear)
df$cyl <- as.factor(df$cyl)
df$am <- as.factor(df$am)


d2 <- df %>%
  group_by(am,cyl) %>%
  summarize(mpg = mean(mpg)) 

plt <- ggplot(d2,aes(cyl,mpg,group = am)) +
  geom_bar(stat="identity",position = "dodge",aes(fill= am))+
  scale_y_continuous(expand = c(0,0))+
  scale_fill_manual(values = c("#E06C83","#6CE0C8"),labels=c("Automatic","Manual"),name = "Transmission") + 
  theme_classic() + 
  labs(y = "Miles Per Gallon",x = "Number of Cylinders")+
  theme(plot.title.position = "panel")

plt

df$vs <- factor(df$vs, levels = c(0, 1), labels = c("V-shaped", "Straight"))


plt2 <- ggplot(df,aes(mpg,disp)) + labs(x=  "Miles Per Gallon",y="Displacement")+
  geom_point(aes(color = gear),size=2) + 
  facet_grid(rows = vars(vs),space = "free",switch = "y") +
  theme(plot.title.position = "panel")

plt2


#OFFICER------------------------------------------------------------

one_column <- block_section(
  prop_section(
   type = "continuous"
  )
)

two_columns <- block_section(
  prop_section( type = "continuous",
    section_columns = section_columns(widths = c(3.7, 1.5),space = 0.5)
  )
)



doc <- read_docx()

doc <- doc %>%  body_add_par("Multiple Plots", style = "heading 1") %>%
  body_add_par("\nComparing MPG and Displacement Across Engine Types",style="heading 2") %>%
  body_add_gg(plt2,height=3.6) %>% body_add_par("\n") %>% body_add_par("\n") %>%
  body_end_block_section(value = one_column) %>%  
  body_add_par("Comparison of Fuel Efficiency across 4-, 6- and 8-cylinder engines",style = "heading 2") %>%
  body_add_gg(plt, width= 3.6,height = 2) %>%
  body_add_par("This page has been divided into two sections:") %>%
  body_add_par("\nThe first section contains one plot, in a single column.") %>%
  body_add_par("The second section contains one plot, and this text, arranged into two columns.")

doc <- body_end_block_section(doc,value = two_columns)

print(doc,"multi_section.docx")
