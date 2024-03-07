
#Mendelian Randomization in Ovarian Cancer
# In this script I looked at #MR data extracted from Ovarian cancer studies 

#I need the following librarries

library(dplyr)
library(openxlsx)
library(janitor)
library(stringr)
library(meta)
library(metafor)
#Data
library(readxl)
mr_oc <- read_excel("U:/Studien/HOC/WCRF CUP/15_Statistical_analysis_Final_summary/5.OC/MR/Dataset_MR_OC_20240124.xlsx",
                    na = c(NA, "", "."))
mr_oc <- clean_names(mr_oc)

mr_oc$author_year <- paste(mr_oc$forname, mr_oc$year_publication, sep =" ")

table(mr_oc$method_mr) 
mr_oc |> count(author_year, method_mr, article_id)


#Ruth KS 2020: MR methods was not extracted but it was 2-sample MR replace the missing value
mr_oc$method_mr <- ifelse(mr_oc$author_year=="Ruth KS 2020", "2 samples R", 
                          mr_oc$method_mr)

#We have 9 MR studies in OC : 8 of them were two-sample MR study 
#Zhu M 2022 was not two-sample MR: correct
mr_oc$method_mr <- ifelse(mr_oc$author_year=="Zhu M 2022", "1 samples R", 
                          mr_oc$method_mr)

mr_oc |> filter(!is.na(explained_variance_percent),
                mr_stat=="IVW") |> 
  group_by(author_year) |>  
  count(i,explained_variance_percent) |> arrange(i)

mr_oc |> count(author_year, method_mr, article_id)
#Filter out rows with NA on 
mr_oc  <- mr_oc |> filter(!is.na(io_stat))
mr_oc |> count(author_year, method_mr, article_id)

unique(mr_oc$i)
mr_oc$i <- ifelse(mr_oc$i=="Bioavailable testosterone", 
                  "Bioavailable Testosterone",mr_oc$i)

ss <- c("Androstendione","DHEA", "DHEAS","Estradiol",
        "Estradiol_free","Estrone", "SHBG","Androstenedione",
        "Free Estradiol","Bioavailable Testosterone",
        "Testosterone","Free testosterone","Unconjugated estradiol",
        "Unconjugated estrone","SHBG-bound estradiol %",
        "Testosterone continuous", "Free testosterone continuous",
        "DHEAS", "16-alpha_OHE-1","2-OHE-1","2OHE1_to_16OHE1_ratio")
#insuling&IGF
ins <- c("Insulin", "C-peptide","IGFBP-1","IGFBP-2",
         "IGF-1","IGFBP-3", "Free IGF-1",
         "IGFB-1","IGFB-2", "IGFB-3")
#inflammation
inf <- c("leptin","Adiponectin","Visfatin", "Leptin",             
         "Adiponectin/leptin","Adiponectin:Both",
         "Adiponectin:Pre","Adiponectin:Post",
         "sTNFR1","sTNFR2",
         "TNF-alpha", "CRP","sOB-R","PAI-1",
         "IL-1ra", "IL-6",
         "TNF-alpha", "IL-8","Il-1ra",
         "IL-1_receptor_type_2")
mr_oc <- mr_oc |> 
  mutate(pathway = case_when(i %in% ss ~ "Sex steroids",
                             i %in% ins ~ "Insulin & IGF",
                             i %in% inf ~ "Inflammation"
  ))

mr_oc$es_CI <- paste(mr_oc$io_statv,
                      "[", mr_oc$io_lower, 
                      ",", mr_oc$io_upper,
                      ']')

#Check data sources
table(mr_oc$studysetting)
mr_oc$source <- ifelse(mr_oc$studysetting=="UK Bio Bank", "UK Biobank",
                       mr_oc$studysetting)
table(mr_oc$source)


#Effect size for Zhu M 2022 is duplicated, remove the duplicated

mr_oc <- mr_oc |> filter(unique_key!=12) #remove rwo #25 it was duplicated on the plot
# MR study by Zhu M 2022


# Now do the risk of bias assessment
# 44 yes/No questions
# A maximum possible score of 44 

mr_oc_rob <- mr_oc |>
  mutate(stat = mr_stat) |> 
  select(author_year, article_id,i,stat, contains("mr_"),
  -mr_stat)
code_vars <- names(mr_oc_rob[,-c(1:4)])
mr_oc_rob<- mr_oc_rob|> 
  mutate(across(all_of(code_vars),
                ~ifelse(.x=="Yes", 1, 0)) 
  )
mr_oc_rob<- mr_oc_rob |> 
  mutate(sum = rowSums(across(starts_with("mr_"))))
#Study with greater or equal to 0.7*44 = Low Risk of bias
0.7*44 
#Stdies with sum >= 30 are low risk of bias
mr_oc_rob$rob <- ifelse(mr_oc_rob$sum>=30, 
                        "Low ROB", "High ROB")
mr_oc_rob |> count(author_year, rob)
# One of the row for Yuan S 2020 is NA replace by the other row for Yuan
mr_oc_rob$rob <- ifelse(mr_oc_rob$author_year=="Yuan S 2020", 
                        "Low ROB",mr_oc_rob$rob)

mr_rob <- mr_oc_rob |> select(author_year, 
                              article_id,i,stat, sum, rob) 
#Now merge it with mr_oc
mr_oc$stat = mr_oc$mr_stat

mr_oc <- cbind(mr_oc,mr_rob)[,-c(93:96)] #Remove duplicated variables
  #Merge didn't work well 

#Now plot
forest(
  metagen(TE=log(io_statv), lower=log(io_lower), upper = log(io_upper), 
          sm="RR",data=mr_oc |> 
            arrange(desc(pathway),mr_stat),
          overall = F,fixed=F, random=F,
          subgroup = i),
  subgroup.name  = "", fontsize=6,spacing = 0.55,
  backtransf = T,
  leftcols = c("source","case_control", "mr_stat"),
  leftlabs = c("Source", "Case Control ratio","MR Stat"),
  rightcols = c("es_CI", "author_year","no_gen_instrument", "rob"),
  rightlabs = c("RR/OR(95%CI)", "Author_year","No of gen instruments",
                "Risk of Bias"
  ),
  squaresize = 0.6,
  col.subgroup = "#00868B"
)



