
# load up those packages or install them if you don't have them -----------

if (!require("tidyverse")) install.packages("tidyverse")
if (!require("officer")) install.packages("officer")
if (!require("glue")) install.packages("glue")



# my custom ggplot theme because defaults are ugly ------------------------

thematics<-
  theme_bw()+
  theme(legend.position = "top",
                            axis.text = element_text(color="black",size=14),
                            axis.title = element_text(face=2,size=18),
                            plot.title = element_text(face=2,size=20),
                            legend.title = element_text(face=2),
                            plot.caption = element_text(face=1,size=10),
                            panel.border = element_blank(),
                            axis.line = element_line(color="black",linewidth =.25),
                            strip.text = element_text(face=2),title = element_text(face = 2),
                            panel.grid = element_blank(),
                            strip.background = element_blank(),
                            text = element_text(family = "sans"))

safePal <-
  c("#ee6677",
    "#4477aa" ,
    "#ccbb44" ,
    "#228833",
    "#66ccee",
    "#aa3377"
    )



# some chickie data -------------------------------------------------------

data("ChickWeight")

chickie<-
  ChickWeight


# make some visuals -------------------------------------------------------

(sillyOLS<-
  chickie %>% 
  ggplot(aes(x=Time,
             y=log(weight)))+
  geom_point(alpha=.25,
             size=1.25)+
  geom_smooth(se=T,
              linewidth=1.5,
              method='lm')+
  labs(title="Log Weight By Time",
       y="Log(GM)",
       x="Time")+
  thematics+
  scale_color_manual(values = safePal))


(chickieEndWeight<-
    chickie %>%
    filter(Time %in% c(21)) %>% 
    ggplot(aes(x=factor(Diet),
               y=weight,
               color=factor(Diet)))+
    geom_boxplot(size=1,
                 show.legend = F)+
    geom_jitter(height = 0,
                width = .25,
                size=3,
                alpha=.7,
                show.legend = F)+
    thematics+
    scale_color_manual(values = safePal)+
    scale_y_continuous(expand = expansion(mult = c(.15)),
                       breaks = seq(0,500,50))+
    labs(title="Chick Weights, Day 21",
         y="Weight (GM)",
         x="Diet"))
  
(dietChart<-
  chickie %>% 
  ggplot(aes(x=Time,
             y=log(weight),
             color=Diet))+
  geom_point(alpha=.25,
             size=1.25)+
  geom_smooth(se=F,
              linewidth=1.5)+
  labs(title="Log Weight By Time",
       y="Log(GM)",
       x="Time")+
  thematics+
  scale_color_manual(values = safePal))




# Some text for our PPT ---------------------------------------------------


  
dietMeans<-
  chickie %>% 
  group_by(Diet,Time) %>% 
  summarise(avg=mean(weight,rm.na=T)) %>% 
  filter(Time %in% c(0,21)) %>% 
  pivot_wider(id_cols = Diet,
              names_from = Time,
              values_from = avg,
              names_prefix ="time_" ) %>% 
  mutate(delta=time_21-time_0)

dynamicText<-
  glue(
  "Diet 1 saw a mean weight increase of {round(dietMeans$delta[dietMeans$Diet==1],1)} from T0 to T21\n
  Diet 2 saw a mean weight increase of {round(dietMeans$delta[dietMeans$Diet==2],1)} from T0 to T21\n  
  Diet 3 saw a mean weight increase of {round(dietMeans$delta[dietMeans$Diet==3],1)} from T0 to T21\n  
  Diet 4 saw a mean weight increase of {round(dietMeans$delta[dietMeans$Diet==4],1)} from T0 to T21
")



# bringing in the PPT template --------------------------------------------

ppt<-read_pptx("my_ppt.pptx")

#editing a slide that already exists
ppt<-on_slide(ppt,
              index = 1)

ppt<-ph_with(ppt, 
             value ="Counting Chickens", 
             location = ph_location_label(ph_label ="title"))

ppt<-ph_with(ppt, 
             value =paste0(format(today(),"%D")), 
             location = ph_location_label(ph_label ="subtitle"))             

#adding a new slide from a layout and dropping a fullsize plot
ppt<-add_slide(ppt,layout = "Blank",master = "Retrospect")

ppt<-ph_with(ppt,
             value=sillyOLS,
             location = ph_location_fullsize())

#adding a slide with some text into predefined elements
ppt<-add_slide(ppt,layout = "text slide",master = "Retrospect")

ppt<-ph_with(ppt, 
             value ="Lorem Ipsum", 
             location = ph_location_label(ph_label ="Title 1"))  

ppt<-ph_with(ppt, 
             value ="
Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor 
incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud 
exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. 
Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. 
Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.", 
location = ph_location_label(ph_label ="Content Placeholder 2"))  

#adding a slide with predefined elements for text and visuals
ppt<-add_slide(ppt,layout = "viz",master = "Retrospect")

ppt<-ph_with(ppt, 
             value ="Boxplots Are Fun", 
             location = ph_location_label(ph_label ="Title 1"))  

ppt<-ph_with(ppt, 
             value = chickieEndWeight, 
             location = ph_location_label(ph_label ="visual"))  

#a more complex example
ppt<-add_slide(ppt,layout = "dual",master = "Retrospect")

ppt<-ph_with(ppt, 
             value ="Chickens Get Bigger When You Feed Them", 
             location = ph_location_label(ph_label ="Title"))  

ppt<-ph_with(ppt, 
             value = dietChart, 
             location = ph_location_label(ph_label ="left"))  

ppt<-ph_with(ppt, 
             value = dynamicText, 
             location = ph_location_label(ph_label ="right"))  

#Generate your slide deck!
print(ppt,
      target = paste0(today(),"-ppt.pptx"))
