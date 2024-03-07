# Ovarian Cancer

#Example analysis

library(dplyr)
library(openxlsx)
library(janitor)
library(stringr)
library(meta)
library(metafor)
library(stringr)
library(grid)

#Data
oc <- read.xlsx("U:/Studien/HOC/WCRF CUP/15_Statistical_analysis_Final_summary/5.OC/Blood/Dataset/final_OC_CC_20240208_Mihiretu.xlsx",
                na.strings=c(NA, "", "."))

oc <- clean_names(oc)
oc <- oc |> 
  mutate(author_year = paste(forname, year_publication))

sum(is.na(oc$i )) #1964 missing 
sum(is.na(oc$i_other )) #2892

table(oc$i)
table(oc$i_other)

sum(is.na(oc$author_year))
sum(is.na(oc$io_statv))
oc$i <- ifelse(oc$i =="Other", oc$i_other, oc$i)

all_missing <- sapply(oc, function(x) all(is.na(x)))
oc <- oc[, !all_missing]
oc <- oc |> filter(!is.na(io_statv)) #No effect size

#oc <- oc[rowSums(is.na(oc))< ncol(oc),] #the same
#oc <- oc |> filter(is.na(generalcomment)) #filter out studies that were excluded in epiinfo 
oc |> count(i, io_statv, author_year)


oc |> count(i, forname) |> count(i)
#Unconjugated Estradiol 
#Unconjugated estradiol 
#Unconjugated estriol 
#Unconjugated estrone 
#IGFBP-2
#IGFBP2

oc <- oc |> mutate(
  i = case_when(i == "Unconjugated estradiol"~"Unconjugated Estradiol",
                i == "IGFBP1"~"IGFBP-1",
                i == "IGFBP2"~"IGFBP-2",
                i == "IGFBP3"~"IGFBP-3",
                i=="IFN-Î³"~"IFN-gamma",
                TRUE~i)
)
table(oc$i)
oc |> count(i, forname) |> count(i)

oc |> count(i, author_year, io_statv)

x <- oc |> group_by(author_year) |> count(i, io_statv)

#Unique studies
unique(oc$author_year)
table(oc$io_adj)

#We have two studies with similar author_year Ose, J 2017: rename
oc$author_year <- ifelse(oc$author_year=="Ose, J 2017"&oc$article_id=="p_28205047",
                         "Ose, J et.al 2017", oc$author_year)


#Cohort
oc <- oc |> 
  mutate(cohort = case_when(author_year == "Chen, T 2011" ~ "FMC", #Finish Maternity Cohort 
                            author_year == "Kabat, G 2018" ~ "WHI",
                            author_year == "Helzlsouer K J 1995" ~ "CLUE I", #Assumption
                            author_year == "Lukanova, A 2003"~ "ORDET+NYUWHS+NSHDS",
                            author_year == "Lukanova, Annekatrin 2003"~ "ORDET+NYUWHS+NSHDS",
                            author_year=="Ose, J 2017" ~"OC3",#Ovarian Cancer Cohort Consortium
                            author_year=="Ose, J et.al 2017" ~ "OC3",
                            author_year == "Knuppel, A 2020" ~ "UK Biobank",
                            author_year == "Watts, E 2021" ~ "UK Biobank",
                            author_year == "Qian, F 2020" ~ "UK Biobank",
                            author_year=="Peeters, P 2007" ~ "EPIC",
                            author_year=="Peres, L 2019" ~"OC3",
                            author_year=="Peres, L 2021" ~ "NHS & NHSII",
                            author_year=="Trabert, B 2019"~"WHI-OS",
                            author_year == "Trabert, B 2016"~"WHI-OS",
                            author_year == "Trabert, B 2014"~ "PLCO",
                            author_year == "Tworoger, SS 2007" ~ "NHS, NHSII & WHS",
                            author_year == "Schock, H 2015"~"FMC & NSMC" #Finish Maternity Cohort and Northern Sweden Maternity Cohort
  ))
#Replace Schock, H 2015 by Schock, H 2014. The year of publication was 2014 not 2015
oc$author_year <- ifelse(oc$author_year=="Schock, H 2015", 
                         "Schock, H 2014", oc$author_year)
oc |> count(author_year, cohort) 

#Pregnancy
table(oc$pregnancy)

oc |> count(author_year, pregnancy) 
## Two studies(Chen, T 2011 and Schock, H 2014) were in pregnant women
## Three studies(Ose, J 2017,Ose, J et.al 2017, Peres, L 2019 ) included pregnant and non-pregnant

#All the aothers are assumed to be "Not pregnant"
oc$pregnancy <- ifelse(oc$pregnancy!="Pregnant"&oc$pregnancy!="Study included both",
                       "Not pregnant", oc$pregnancy)


oc$pregnancy <- ifelse(oc$pregnancy=="Study included both", "Both", oc$pregnancy)
oc$pregnancy <- ifelse(oc$pregnancy=="Not reported", "Not pregnant", oc$pregnancy)
oc$pregnancy <- ifelse(is.na(oc$pregnancy), "Not pregnant", oc$pregnancy)

#HRT use
table(oc$hrto_cuse)
oc |> count(author_year, hrto_cuse)

# Nine studies excluded women on HRT
# Eight included women on HRT but adjusted for it in the analysis

#Medication
table(oc$medication_e_o) #None reported for all
table(oc$fasting)
oc |> filter(fasting=="Fasting") |>  count(author_year, i)
# Only one study used Fasting sample for Insulin

table(oc$fasting_other)
table(oc$generalcomment)
table(oc$matchingfactors)
#Physiolo
table(oc$physiological_condition_cases)
table(oc$i_all_change_stat)
table(oc$o_s1)
table(oc$o_s2)

cc <- oc |> count(author_year, i, o_s1, o_s2, io)

#Creat Oc_new data frame
oc_new <- oc |> 
  select(author_year,cohort, title, country, article_id, 
         studydesign, ethnicity_race, study_population, study_setting,
         menopause_status, menstrual_cycle, hrto_cuse, 
         fasting, fasting_other, total_sample_size,io_n, io_n, 
         sample_size_cases, sample_size_control,
         i, i_transformation, i_units,i_units_other,  
         i_specimen, i_measure, 
         o_s1, o_s1_n, o_s1_average,o_s1_av_value,
         o_s1_variance,
         o_s1_pval,
         o_s2, o_s2_n, o_s2_average,
         o_s2_variance,
         io,io_ref, io_comp, io_n, io_stat,io_statv,
         io_var,io_varv, io_p, io_pt, io_adj, 
         responserate, missingproportion, missinghandeled,commentsonmissing,
         impactbias, impactbias_just,indirectness, indirectness_justification,
         starts_with("cc_"),
         starts_with("c_")) 

table(oc_new$menopause_status)
oc_new$menopause_status <- ifelse(str_detect(oc_new$menopause_status,
                                             "Not adjusted for"), 
                                  "Both-not adjusted",
                                  oc_new$menopause_status)
oc_new$menopause_status <- ifelse(str_detect(oc_new$menopause_status,
                                             "Both - Adjusted for"), 
                                  "Both-adjusted",
                                  oc_new$menopause_status)

table(oc_new$io_var)
oc_new$io_var <- ifelse(oc_new$io_var =="95%Ci", "95%CI",
                        oc_new$io_var)

unique(oc_new$io)
oc_new |> filter(io=="HR") |> count(author_year) #Replaced by OC VS Control

oc_new <- oc_new |> 
  mutate(
    io = case_when(
      io == "Non-serous*Control" ~ "Non-serous OC VS Control",
      io == "Nonserous OC*Control" ~ "Non-serous OC VS Control",
      io == "Non-serous OC*Control" ~ "Non-serous OC VS Control",
      io == "Non-Serous OC*Control"~ "Non-serous OC VS Control",
      io == "Clear Cell OC*Control" ~ "Clear cell OC VS Control",
      io == "Clear cell OC*Control" ~ "Clear cell OC VS Control",
      io == "Clear cell*Control" ~ "Clear cell OC VS Control",
      io == "EC*control" ~ "Endometroid OC VS Control",
      io == "Endometrioid OC*Control" ~ "Endometroid OC VS Control",
      io == "Endometroid*Control" ~ "Endometroid OC VS Control",
      io == "OC*Control (Mendelian Randomization)" ~ "OC VS Control",
      io == "HR (Mendelian Randomization)" ~ "OC VS Control",
      io == "HR (Mendelian Randomization" ~"OC VS Control",
      io == "OC*Contro" ~"OC VS Control",
      io == "OC*Control "~ "OC VS Control",
      io == "OC*Control" ~ "OC VS Control",
      io == "Serous*Control" ~ "Serous OC VS Control",
      io == "Serous OC*Control" ~ "Serous OC VS Control",
      io == "Mucinous OC*Control" ~ "Mucinous OC VS Control",
      io == "SCST OC*Control" ~ "SCST OC VS Control",
      io == "GCT OC*Control"~ "GCT OC VS Control",
      io == "Borderline serous*Control" ~ "Borderline serous OC VS Control",
      io == "Invasive serous*Control" ~ "Invasive serous OC VS Control",
      io =="Borderline mucinous" ~ "Borderline mucinous OC VS Control",
      io == "Borderline mucinous*Control"~ "Borderline mucinous OC VS Control",
      io == "Invasive mucinous" ~ "Invasive mucinous OC VS Control",
      io == "Invasive mucinous*Control"~ "Invasive mucinous OC VS Control",
      io == "HR"~"OC VS Control", # Kabat study compared risk in OC V Control 
      TRUE ~ io
    )
  )
oc_new |> filter(is.na(io)) |> count(author_year)
oc_new$io <- ifelse(is.na(oc_new$io)&oc_new$author_year=="Helzlsouer K J 1995",
                 "OC VS Control", oc_new$io)
unique(oc_new$io)


#Overall, it is best to keep them separate. 
#However, we can combine clear cell, mucinous, 
            #endometrioid to a non-serous group.
# oc_new <- oc_new |> 
#   mutate(io_cat= case_when(io == "Mucinous OC*Control" ~"Non-serous OC*Control",
#                            io == "Clear cell OC*Control"~"Non-serous OC*Control",
#                            io =="Endometrioid OC*Control"~ "Non-serous OC*Control",
#                            TRUE ~ oc_new$io))




adj <- oc_new |> count(author_year, i,io_adj, io_statv, io, io_comp)
oc_new$io_adj <- ifelse(oc_new$author_year=="Trabert, B 2019", 
                        "Multivariable adjusted", oc_new$io_adj) #All are adjusted
#adjust: age (model 1)
oc_new$io_adj <- ifelse(oc_new$author_year=="Kabat, G 2018"&oc_new$io_adj=="adjust: age (model 1)", 
                        "Unadjusted", oc_new$io_adj) #
oc_new$io_adj <- ifelse(oc_new$author_year=="Kabat, G 2018"&str_detect(oc_new$io_adj, "model 2"), 
                       "Multivariable adjusted", oc_new$io_adj) 

oc_new$io_adj <- ifelse(str_detect(oc_new$io_adj, "model 2"),
                                    "Multivariable adjusted", oc_new$io_adj)
oc_new$io_adj <- ifelse(is.na(oc_new$io_adj)&oc_new$author_year=="Schock, H 2014",
                        "Multivariable adjusted", oc_new$io_adj)

#Check Tworoger 207
tworoger <- oc_new |> 
  filter(author_year=="Tworoger, SS 2007") |> 
  count(i, io_adj, io_comp, io_statv)

#IGF-1: Multivariable adjusted plus IGFBP-2 and IGFBP-3 , iostatv = 0.51
#IGFBP-2: Multivariable adjusted plus IGF-1 and IGFBP-3, iostatv = 0.83
#IGFBP-3: Multivariable adjusted plus IGF-1 and IGFBP-2, iostatv = 1.07

unique(oc_new$io_adj)

oc_new <- oc_new |> 
  mutate(io_adj = case_when(
    author_year == "Tworoger, SS 2007" & i == "IGF-1" & str_detect(io_adj, "IGF") ~ "Multivariable adjusted + IGFBP-2 & IGFBP-3",
    author_year == "Tworoger, SS 2007" & i == "IGFBP-2" & str_detect(io_adj, "IGF") ~ "Multivariable adjusted + IGF-1 & IGFBP-3",
    author_year == "Tworoger, SS 2007" & i == "IGFBP-3" & str_detect(io_adj, "IGF") ~ "Multivariable adjusted+IGF-1 & IGFBP-2",
    author_year == "Tworoger, SS 2007" & !str_detect(io_adj, "IGF") ~ "Multivariable adjusted",
    io_adj == "Multivariable adusted" ~ "Multivariable adjusted",
    io_adj == "Mutlivariable adjusted"~ "Multivariable adjusted",
    TRUE ~ io_adj
  ))


table(oc_new$menopause_status)
table(oc_new$io_var)
oc_new$io_var <- ifelse(oc_new$io_var =="9%CI", "95%CI",
                        oc_new$io_var)
table(oc_new$io_var)
table(oc_new$io)
unique(oc_new$io_varv)
ci <- oc_new |> 
  filter(str_detect(io_varv, "â€")) |>  
  count(author_year, i, io_statv, io_varv)

oc_new$io_varv <- str_replace(oc_new$io_varv, "â€", "-")
oc_new$io_varv <- str_replace(oc_new$io_varv, '“', "-")

adj <- oc_new |> count(author_year, i,io_adj, io_statv,io_varv, io, io_comp)

#Separate CI to lower and upper
library(tidyr)

is.character(oc_new$io_varv)
oc_new$conf_int <- as.character(oc_new$io_varv) 



oc_new$conf_int <- gsub(" ", "", oc_new$conf_int)
table(str_detect(oc_new$conf_int, ","))
oc_new$conf_int <- gsub(",", "-", oc_new$conf_int)
oc_new$conf_int <- gsub("--", "-", oc_new$conf_int)

oc_new <- extract(oc_new, conf_int, 
                  into = c("lower", "upper"), regex = "^(.*)-(.*)$")

#Check
na <- oc_new |> select(io_varv, lower, upper)
sum(is.na(oc_new$lower))
sum(is.na(oc_new$upper))
sum(is.na(oc_new$io_varv))
rm(na)

#Change to numeric 
sum(is.na(oc_new$io_stat))
sum(is.na(oc_new$io_statv))
is.numeric(oc_new$upper)
sum(is.na(oc_new$sample_size_cases))
sum(is.na(oc_new$sample_size_control))

oc_new <- oc_new |> 
  mutate_at(vars(io_statv, lower, upper, sample_size_cases,
                 sample_size_control), as.numeric)


table(oc_new$studydesign)
oc_new |> filter(studydesign=="Case_control") |> count(author_year)
oc_new$studydesign <- ifelse(oc_new$author_year=="Helzlsouer K J 1995", 
                             "nested case-control", oc_new$studydesign) #It is Nested cc study

#Analyze #Risk of bias assessment data
case_control <- oc_new |> filter(studydesign!="Cohort")
cohort <- oc_new |> filter(studydesign=="Cohort")

#Nested-case Control: ROB
code_vars   <- c("cc_objectives", "cc_hypothesis", 
                 "cc_descriptionpopulation",
                 "cc_participationrate",
                 "cc_selectionparticipans", "cc_samplesize", 
                 "cc_samplesize1", "cc_exposureprior", "cc_fu", #confusing
                 "cc_exposurecategory",
                 "cc_omparabilitymeasurements",
                 "cc_measureoutcome", #cc_recallbias(no=1) 
                 "cc_accessorsblinded","cc_confounding",
                 "cc_analysismethod", "cc_missingdata"#NA and Yes are 1
                 #"cs_biasselectiona" : no=1
                 #"cs_biasselectionb", : no=1
                 #"cs_biasselectionc" : n0=1
)


#max score 4
reverse_code <- c("cc_recallbias",
                  "cc_biasselectiona", #no=1
                  "cc_biasselectionb", #no=1
                  "cc_biasselectionc" # n0=1
)



case_control<- case_control|> 
  mutate(across(all_of(code_vars),
                ~ifelse(.x=="Yes", 1, 0)) 
  )


case_control<- case_control|> 
  mutate(across(all_of(reverse_code),
                ~ifelse(.x=="No", 1, 0))
  )

table(case_control$cc_missingdata)
sum(is.na(case_control$cc_missingdata))

case_control$cc_missingdata <- ifelse(is.na(case_control$cc_missingdata),
                                1,case_control$cc_missingdata)


table(case_control$cc_missingdata)
table(case_control$cc_confounding)
sum(is.na(case_control$cc_confounding)) #replace the missing NA in cc_confounding
case_control$cc_confounding <- ifelse(is.na(case_control$cc_confounding), 1,
                                case_control$cc_confounding )



new_rob <- case_control|> 
  mutate(across(all_of(reverse_code),
                ~ifelse(.x=="1", "Yes", "No"))
  )
new_rob <- case_control|> 
  mutate(across(all_of(reverse_code),
                ~ifelse(.x=="1", "No", "Yes"))
  )

x <- lapply(case_control[code_vars], table) |> as.data.frame() |> 
  as.data.frame() |> t() 
rm(new_rob, x)

#x |> filter(str_detect(rownames(x), "Freq")==TRUE)
case_control$impactbias <- as.factor(case_control$impactbias)
table(case_control$impactbias)

case_control<- case_control|>
  mutate(impactbias = case_when(
    impactbias=="maximal" ~ 0,
    impactbias=="medium" ~ 1,
    impactbias=="minimal" ~ 2))

#Replace #Case control by nested case-contrl
# Cohort: ROB

code_vars_cohort   <- c("c_objectives", "c_hypothesis", 
                 "c_descriptionpopulation",
                 "c_participationrate",
                 "c_selectionparticipans", "c_samplesize", 
                 "c_samplesize1", "c_exposureprior", "c_fu", #confusing,
                 "c_lossfu",
                 "c_exposurecategory",
                 "c_comparabilitfmeasurements",
                 "c_measureoutcome", #c_recallbias(no=1) 
                 "c_assessorsblinded","c_confounding",
                 "c_analysismethod", "c_missingdata"#NA and Yes are 1
                 #"cs_biasselectiona" : no=1
                 #"cs_biasselectionb", : no=1
                 #"cs_biasselectionc" : n0=1
)


#max score 4
reverse_code_cohort <- c("c_recallbias",
                  "c_biasselectiona", #no=1
                  "c_biasselectionb", #no=1
                  "c_biasselectionc" # n0=1
)


x <- oc_new |> 
  filter(studydesign=="Cohort") |> 
  select(starts_with("c_"), studydesign, author_year)

cohort |> select(c_recallbias, c_biasselectiona, 
                 c_biasselectionb, c_biasselectionc, 
                 c_lossfu, c_recallbias )

cohort<- cohort|> 
  mutate(across(all_of(code_vars_cohort),
                ~ifelse(.x=="Yes", 1, 0)) 
  )


cohort<- cohort|> 
  mutate(across(all_of(reverse_code_cohort),
                ~ifelse(.x=="No", 1, 0))
  )

table(cohort$c_missingdata)
sum(is.na(cohort$c_missingdata))

cohort$c_missingdata <- ifelse(is.na(cohort$c_missingdata),
                                1,cohort$c_missingdata)


table(cohort$c_missingdata)
table(cohort$c_confounding)
sum(is.na(cohort$c_confounding)) 


table(cohort$impactbias)
sum(is.null(cohort$impactbias))
#Sum
case_control<- case_control |> 
  mutate(sum = rowSums(across(c(starts_with("cc_"), 
                                impactbias))))
cohort<- cohort|>
  mutate(impactbias = case_when(
    impactbias=="maximal" ~ 0,
    impactbias=="medium" ~ 1,
    impactbias=="minimal" ~ 2))

cohort$c_comments <- 0
cohort$c_compliance <- 0
cohort<- cohort |> 
  mutate(sum = rowSums(across(c(starts_with("c_"), 
                                impactbias))))

case_control |> count(author_year, sum)
# The rOB data for the following studies was not correctly reterived
#Add manually calculated from EpiInfo
#Helzlsouer K J 1995: 1+1+1+ 1+1+0+ 0+1+1 +1+1+1+ 0+0+1+ 1+0+1+ 1+1 + 0
#Schock, H 2014: 1+1+1+ 1+1+1+ 1+1+1+ 1+1+1+ 0+1+1+ 1+1+1+ 1+1+2

case_control$sum <- ifelse(case_control$author_year=="Helzlsouer K J 1995", 
                           15, case_control$sum)
case_control$sum <- ifelse(case_control$author_year=="Schock, H 2014", 
                           21, case_control$sum)
#Maximum attainable score for case_control is 22: 0.75*22= 16.5
#Maximum attainable score for cohort is 24: 0.75*22= 18

#Check median
median(case_control$sum, na.rm=T) #19
median(cohort$sum, na.rm=T) #18

max(case_control$sum, na.rm=T) #21
max(cohort$sum, na.rm=T) #21

min(case_control$sum, na.rm=T) #15
min(cohort$sum, na.rm=T) #15

case_control$bias <- ifelse(case_control$sum<16.5, "High risk of bias",
                      "Low risk of bias")
cohort$bias <- ifelse(cohort$sum<18, "High risk of bias",
                            "Low risk of bias")

# Merge 
oc_new <- rbind(case_control, cohort)
oc_new |> group_by(author_year, studydesign) |> count(sum, bias) #

# Remove all the columns with rob variables
oc_new <- oc_new |> 
  select(-starts_with("cc_"), -starts_with("c_"))


unique(oc_new$i) # 31 different intermediate exposures

x <- oc_new |> count(author_year, i) |> count(author_year, i)
oc_new |> filter(is.na(io_comp)) |> 
  count(author_year, i, io_comp)

x <- oc_new |> filter(is.na(io_comp)|is.na(io_ref)) |> 
  count(author_year, i,io,i_transformation,io_ref, io_comp, io_statv)
schock <- oc_new |> filter(is.na(io_comp)|is.na(io_ref), author_year=="Schock, H 2014") |> 
  count(author_year, i,io,i_transformation,io_ref, io_comp, io_statv)

#io_comp is Continuous for Chen, T 2011: Manually checked from the pdfs
oc_new <- oc_new |> 
  mutate(io_comp= case_when(author_year=="Knuppel, A 2020"&is.na(io_comp) ~ "Continuous",
                   author_year == "Chen, T 2011"&is.na(io_comp) ~ "Continuous",
                   author_year=="Ose, J 2017"&is.na(io_comp) ~ "Continuous",
                   author_year == "Ose, J et.al 2017"&is.na(io_comp) ~ "Continuous",
                   author_year == "Peres, L 2019"& is.na(io_comp)~ "Continuous",
                   author_year =="Qian, F 2020" &is.na(io_comp)~ "Continuous",
                   author_year=="Watts, E 2021"&is.na(io_comp)~ "Continuous",
                   author_year =="Schock, H 2014"&is.na(io_comp)~ "Continuous",
                   TRUE~io_comp),
         io_ref = case_when(
           author_year =="Schock, H 2014"&	io_comp == "Tertile 3"~ "Tertile 1",
           TRUE ~ io_ref)
           )

# recode based on the similarity in type of OC 
unique(oc_new$io)
oc_new |> filter(io == "GCT OC VS Control"|io == "SCST OC VS Control") |>  
  count(author_year, i)
#GCT: germ cell tumors 
#SCST: sex cord stromal tumors

oc_new <- oc_new |> 
  mutate(io_new = case_when(
    io == "Borderline mucinous OC VS Control" ~ "Mucinous OC VS Control",
    io == "Invasive mucinous OC VS Control"~ "Mucinous OC VS Control",
    io == "Invasive serous OC VS Control" ~ "Serous OC VS Control",
    io == "Borderline serous OC VS Control"~ "Serous OC VS Control",
    TRUE ~ io
  ))
unique(oc_new$io_new)
#Recode
unique(oc_new$io_comp)
#Tertile 2 is wrong
oc_new |> filter(io_comp=="Tertile 2" ) |> count(author_year, i, io_statv)
#Correct
#oc_new$io_comp <- ifelse(oc_new$io_comp=="Tertile 2", "Tertile 3", 
 #                        oc_new$io_comp)
#Or here together with others
oc_new <- oc_new |> 
  mutate(io_comp = case_when(io_comp=="Quartile 4"~ "Q4 Vs Q1",
                             io_comp=="Tertile 3"~ "T3 Vs T1",
                             io_comp=="Quntile 5"~ "Q5 Vs Q1",
                             io_comp=="Clinical-based Tertile 3"~ "T3 Vs T1",
                             io_comp=="Tertile 2"~ "T3 Vs T1",
                             io_comp=="High"~ "High Vs Low",
                             io_comp=="Quintile 5"~ "Q5 Vs Q1",
                             io_comp == ">10 mg/L" ~ ">10mg/L vs <1mg/dl",
                             io_comp == "Continous"~"Continuous",
                             TRUE~io_comp))

# All categorical comparisons to high/vslow
unique(oc_new$io_comp)
oc_new$io_comp_new_cat <- ifelse(oc_new$io_comp!="Continuous", "High vs Low",
                                 oc_new$io_comp)
unique(oc_new$io_comp_new_cat)

oc_new <- oc_new |> 
  mutate(io_new_cat = case_when(
    io_new == "Endometroid OC VS Control" ~ "Non-serous OC VS Control",
    io_new == "Mucinous OC VS Control" ~ "Non-serous OC VS Control",
    io_new == "Clear cell OC VS Control" ~ "Non-serous OC VS Control",
    
    TRUE ~ io_new
  ))

#Estradiol: 4 studies
e2_meta <- oc_new |> 
  filter(i=="Estradiol") |> unique()
e2_meta |> count(author_year, menopause_status, io, io) # 2 pre, 1 post and 1 both


e2_meta |> filter(io == "OC VS Control") |> 
  count(author_year, menopause_status, io, io_statv, io_varv, io_comp)

e2_pool <- metagen(TE = log(io_statv), lower = log(lower), 
                   upper = log(upper),
                   subgroup = paste(io_new_cat,io_adj, sep = " | "),
                   subgroup.name = "",
                   data = e2_meta |> filter(io == "OC VS Control",
                                            io_comp != "Continuous") |> arrange(desc(io_adj)),
                   sm = "RR",
                   common = FALSE,
                   overall = TRUE,
                   studlab = author_year,
                   backtransf = TRUE,
                   method.tau = "PM")

forest(e2_pool, leftcols = c("author_year","cohort", 
                             "io", 
                             "io_stat","io_n"),
       subgroup.name = "", overall = F,
       leftlabs = c("Study","Cohort", "Comparison", 
                    "Measure of association",
                    "Case/Control_ratio"),
       col.subgroup = "#00868B",spacing = 0.7,
       fontsize = 6)
grid.text("Pre-diagnostic circulating Estradiol levels and the risk of Ovarian cancer in pre- and postmenopausal women",
          x = 0.5, y = 0.95, 
          gp = gpar(fontsize = 6, fontface = "bold"))



e2_pool <- metagen(TE = log(io_statv), lower = log(lower), 
                   upper = log(upper),
                   subgroup = paste(i, io_new_cat,io_adj, sep = " | "),
                   subgroup.name = "",
                   data = e2_meta |> arrange(desc(io_adj)),
                   sm = "OR",
                   common = FALSE,random = FALSE,
                   overall = FALSE,
                   studlab = author_year,
                   backtransf = TRUE,#method.tau = "PM"
                   )

forest(e2_pool, leftcols = c("author_year","cohort", 
                             "io", "io_comp","io_n"),
       subgroup.name = "", overall = F,
       leftlabs = c("Study","Cohort", "Comparison", "Ref group",
                    "Case/Control_ratio"),
       col.subgroup = "#00868B",spacing = 0.7,
       fontsize = 6)
grid.text("Pre-diagnostic circulating Estradiol levels and the risk of Ovarian cancer in pre- and postmenopausal women",
          x = 0.5, y = 0.97, 
          gp = gpar(fontsize = 8, fontface = "bold"))


#DHEAS: 4 studies
dheas_meta <- oc_new |> 
  filter(i=="DHEAS") |> unique()
dheas_meta |> count(author_year, menopause_status, io) # 2 pre, 1 post and 1 both


dheas_meta |> filter(io == "OC VS Control") |> 
  count(author_year, menopause_status, io, io_statv, io_varv, io_comp)

dheas_pool <- metagen(TE = log(io_statv), lower = log(lower), 
                   upper = log(upper),
                   subgroup = paste(io_new,io_adj, sep = " | "),
                   subgroup.name = "",
                   data = dheas_meta |> filter(io == "OC VS Control",
                                            io_comp != "Continuous") |> arrange(desc(io_adj)),
                   sm = "RR",
                   common = FALSE,
                   overall = TRUE,
                   studlab = author_year,
                   backtransf = TRUE,
                   method.tau = "PM")

forest(dheas_pool, leftcols = c("author_year","cohort", 
                             "io", 
                             "io_stat","io_n"),
       subgroup.name = "", overall = F,
       leftlabs = c("Study","Cohort", "Comparison", 
                    "Measure of association",
                    "Case/Control_ratio"),
       col.subgroup = "#00868B",spacing = 0.7,
       fontsize = 6)
grid.text("Pre-diagnostic circulating Estradiol levels and the risk of Ovarian cancer in pre- and postmenopausal women",
          x = 0.5, y = 0.95, 
          gp = gpar(fontsize = 6, fontface = "bold"))


# Function
metagen_func <- function(df) {
  metagen(TE = log(io_statv), lower = log(lower), 
          upper = log(upper),
          subgroup = paste(i, io_new_cat, io_adj, sep = " | "),
          subgroup.name = "",
          data = df |> arrange(desc(io_adj)),
          sm = "OR",
          common = FALSE, random = FALSE,
          overall = FALSE,
          studlab = author_year,
          backtransf = TRUE  #method.tau = "PM"
  )}

forest_func <- function(model, i) {
  forest(model, leftcols = c("author_year", "cohort", 
                             "io", "io_comp", "io_n"),
          subgroup.name = "", overall = FALSE,
          leftlabs = c("Study", "Cohort", "Comparison", "Ref Comparison",
                       "Case/Control_ratio"),
          col.subgroup = "#00868B", spacing = 0.7,
          fontsize = 6)
  
  title_text <- paste("Pre-diagnostic circulating", i, "levels and the risk of Ovarian cancer in pre- and postmenopausal women")
  grid.text(title_text, x = 0.5, y = 0.97, 
            gp = gpar(fontsize = 8, fontface = "bold"))
}

my_forest_plot <- function(df) {
  
  metagen(TE = log(io_statv), lower = log(lower), 
                     upper = log(upper),
                     subgroup = paste(i, io_new, io_adj, sep = " | "),
                     subgroup.name = "",
                     data = e2_meta |> arrange(desc(io_adj)),
                     sm = "OR",
                     common = FALSE, random = FALSE,
                     overall = FALSE,
                     studlab = author_year,
                     backtransf = TRUE  #method.tau = "PM"
  )
  
  forest(pool, leftcols = c("author_year", "cohort", 
                               "io", "io_comp", "io_n"),
         subgroup.name = "", overall = FALSE,
         leftlabs = c("Study", "Cohort", "Comparison", "Ref Comparison",
                      "Case/Control_ratio"),
         col.subgroup = "#00868B", spacing = 0.7,
         fontsize = 6)
  
  title_text <- paste("Pre-diagnostic circulating", df$i, "levels and the risk of Ovarian cancer in pre- and postmenopausal women")
  grid.text(title_text, x = 0.5, y = 0.97, 
            gp = gpar(fontsize = 8, fontface = "bold"))
}

#Estradiol: 4 studies
forest_func(e2_pool <- metagen_func(e2_meta), "Estradiol")
#Unconjugated Estradiol: 2 studies
e2_unconjugated_meta <- oc_new |> filter(i=="Unconjugated Estradiol")
forest_func(e2_unconjugated_pool <- metagen_func(e2_unconjugated_meta), "Unconjugated Estradiol")

dheas_meta <- oc_new |> filter(i=="DHEAS")
forest_func(dheas_pool <- metagen_func(dheas_meta), "DHEAS")
#DHEA: 2 studies
dhea_meta <- oc_new |> filter(i=="DHEA")
forest_func(dhea_pool <- metagen_func(dhea_meta), "DHEA")

#Testosterone: 4 studies
tt_meta <- oc_new |> filter(i=="Testosterone")
forest_func(tt_pool <- metagen_func(tt_meta), "Testosterone")
# Testosterone_free: 2 studies
tt_free_meta <- oc_new |> filter(i=="Testosterone_free")
forest_func(tt_free_pool <- metagen_func(tt_free_meta), "Free testosterone")
# Estron: 3 studies
e1_meta <- oc_new |> filter(i=="Estrone")
forest_func(e1_pool <- metagen_func(e1_meta), "Estrone")
#	Unconjugated estrone: 2 studies
e1_unconjugated_meta <- oc_new |> filter(i=="Unconjugated estrone")
forest_func(e1_unconjugated_pool <- metagen_func(e1_unconjugated_meta), "Unconjugated estrone")
#Androstendione: 4 studies
a1_meta <- oc_new |> filter(i=="Androstendione")
forest_func(a1_pool <- metagen_func(a1_meta), "Androstendione")

##SHBG: 2 studies
shbg_meta <- oc_new |> filter(i=="SHBG")
forest_func(shbg_pool <- metagen_func(shbg_meta), "SHBG")
#IGF-1: 4 studies
igf1_meta <- oc_new |> filter(i=="IGF-1")
forest_func(igf1_pool <- metagen_func(igf1_meta), "IGF-1")

#IGFBP-2: 2 studies
igfbp2_meta <- oc_new |> filter(i=="IGFBP-2")
forest_func(igfbp2_pool <- metagen_func(igfbp2_meta), "IGFBP-2")
##IL-8: 2 studies
il8_meta <- oc_new %>%
  filter(i == "IL-8") %>%
  distinct(io_statv, .keep_all = TRUE)

forest_func(il8_pool <- metagen_func(il8_meta), "IL-8")
#Progesterone
prog_meta <- oc_new |> filter(i=="Progesterone") |> unique()
forest_func(prog_pool <- metagen_func(prog_meta), "Progesterone")

hormones <- oc_new |> count(author_year, i) |> arrange(i)
#C-peptide, 
#C_reactive_protein_CRP, DHT_Dihydroesterone, 
#Estriol, Estriol_conjugated, Estrone_conjugated, IFN-gamma, IGFBP-1, 	
#IGFBP-3, IL-10, IL-1B, IL-1a, IL-2, Insulin
other_hormones <- c("C-peptide", "C_reactive_protein_CRP", "DHT_Dihydroesterone",
                    "Estriol", "Estriol_conjugated", "Estrone_conjugated", 
                    "IFN-gamma", "IGFBP-1", "IGFBP-3","17-alpha-hydroxyprogesterone", 
                    "IL-10", "IL-1B", "IL-1a", "IL-2", "Insulin")

others <- oc_new |> filter(i %in% other_hormones) |> unique() #deduplicate

others_pool <- metagen(TE = log(io_statv), lower = log(lower), 
                   upper = log(upper),
                   subgroup = paste(i, io_new_cat,io_adj, sep = " | "),
                   subgroup.name = "",
                   data = others |> arrange(i),
                   sm = "RR",
                   common = FALSE,random = FALSE,
                   overall = FALSE,
                   studlab = author_year,
                   backtransf = TRUE,
                   method.tau = "PM")

svg("U:/Studien/HOC/WCRF CUP/15_Statistical_analysis_Final_summary/5.OC/Blood/Mihiretu/others_pool.svg", width = 12, height = 23)

forest(others_pool, leftcols = c("author_year", "cohort", 
                                 "io", "io_comp", "io_n"),
       subgroup.name = "", overall = FALSE,
       leftlabs = c("Study", "Cohort", "Comparison", "Ref Comparison",
                    "Case/Control_ratio"),
       col.subgroup = "#00868B",spacing = 0.7,
       fontsize = 8)
dev.off()
#	

#Direction
oc_new <- oc_new |> 
  mutate(direction = case_when(
    io_statv<1&lower<1&upper<1 ~ "Negative",
    io_statv>1&lower>1&upper>1 ~ "Positive",
    TRUE ~ "Null"
  ))
x <- oc_new |> count(i, io_statv, lower, upper, direction)  

#Write data for plotly plots in python
oc_plotly <- oc_new |> 
  select(author_year,article_id,menopause_status, i,io_new, io_comp, io_ref, 
        bias, cohort, io_adj, io_statv, lower, upper, direction) |> unique()
unique(oc_plotly$i)

sex_steroids <- c("17-alpha-hydroxyprogesterone", "Androstendione", 
                  "DHEA", "DHEAS", "DHT_Dihydroesterone", "Estradiol", 
                  "Estradiol_conjugated", "Estriol", "Estriol_conjugated",
                  "Estrone", "Estrone_conjugated", "Progesterone", 
                  "Testosterone", "Testosterone_free", 
                  "Unconjugated Estradiol", "Unconjugated estriol", 
                  "Unconjugated estrone", "SHBG")
iflammation <- c("C_reactive_protein_CRP", 
                "IFN-gamma", "IL-10", 
                "IL-1a", "IL-1B", "IL-2", "IL-8")
isulin_igf <- c("C-peptide", "IGF-1", "IGFBP-1", 
               "IGFBP-2", "IGFBP-3", "Insulin")

oc_plotly <- oc_plotly %>%
  mutate(pathway = case_when(
    i %in% sex_steroids ~ "Sex steroid",
    i %in% iflammation ~ "Iflammation",
    i %in% isulin_igf ~ "Insulin & IGF",
    TRUE ~ "other"
  )) 

table(oc_plotly$pathway)
#write.csv(oc_plotly, "oc_plotly.csv")

#Now meta-analyze

# Function
# We use Paul-Mandel Tau-square estimator
# Bakbergenuly and colleagues (2020) found that the Paule-Mandel estimator is 
# well suited especially when the number of studies is small.
# Paule-Mandel estimator is a good first choice, provided 
# there is no extreme variation in the sample sizes

metagen_func_pool <- function(df) {
  metagen(TE = log(io_statv), lower = log(lower), 
          upper = log(upper),
          subgroup = paste(i,io_comp_new_cat, io_new_cat, io_adj, sep = " | "),
          subgroup.name = "",
          data = df |> arrange(desc(io_adj)),
          sm = "OR",
          common = FALSE, random = TRUE,overall.hetstat = F,
          overall = FALSE,
          studlab = author_year,test.subgroup = F,
          backtransf = TRUE, method.tau = "PM"
  )}

forest_func_pool <- function(model, i) {
  forest(model, leftcols = c("author_year", "cohort", 
                             "io", "io_comp", "io_n"),
         subgroup.name = "", overall = FALSE,
         leftlabs = c("Study", "Cohort", "Comparison", "Ref Comparison",
                      "Case/Control_ratio"),
         col.subgroup = "#00868B", spacing = 0.7,
         fontsize = 7)
  
  title_text <- paste("Pre-diagnostic circulating", i, "levels and the risk of Ovarian cancer in pre- and postmenopausal women")
  grid.text(title_text, x = 0.5, y = 0.97, 
            gp = gpar(fontsize = 8, fontface = "bold"))
}

# Meta-analyzable???
#E2
#Non-serous vs Control, Serous and OC vs control
unique(e2_meta$io_new_cat)
forest_func_pool(e2_pooled <- metagen_func_pool(e2_meta |> filter(
  (
    (io_new_cat == "Non-serous OC VS Control" | 
       io_new_cat == "Serous OC VS Control" | 
       io_new_cat == "OC VS Control") & 
      !(io_new_cat == "OC VS Control" & io_comp_new_cat == "Continuous")
  ) | 
    author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017") #Not to filter out the consortium studies
)), 
    "Estradiol")
#Peres, L 2019, Ose, J et.al 2017, Ose, J 2017: Consortium studies 

forest_func_pool(metagen_func_pool(e2_meta ), 
  "Estradiol")

#Unconjugated E2
# Similar cohort: Meta-analysis not needed 
unique(e2_unconjugated_meta$io_new_cat)

forest_func_pool(unconj_e2_pooled <- metagen_func_pool(e2_unconjugated_meta ), "Unconjugated Estradiol")

#DHEAS
forest_func_pool(metagen_func_pool(dheas_meta ), "DHEAS")

forest_func_pool(dheas_pooled <- metagen_func_pool(dheas_meta |> 
                                                     filter(
                                                       ((io_new_cat == "Non-serous OC VS Control" | io_new_cat == "OC VS Control") &
                                                           !(io_new_cat == "OC VS Control" & io_comp_new_cat == "Continuous")) | 
                                                         author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                 "DHEAS")

# DHEA: onöly one meta-analysis
forest_func_pool(metagen_func_pool(dhea_meta ), "DHEA")

forest_func_pool(dhea_pooled <- metagen_func_pool(dhea_meta |> 
                                                     filter(io_new_cat== "OC VS Control" )),
                 "DHEA")
#TT
forest_func_pool(metagen_func_pool(tt_meta ), "Testosterone")

forest_func_pool(tt_pooled <- metagen_func_pool(tt_meta |> 
                                                  filter(
                                                    ((io_new_cat =="Non-serous OC VS Control"|
                                                           #io_new_cat =="Serous OC VS Control"|
                                                           io_new_cat== "OC VS Control" )&
                                                           !(io_new_cat == "OC VS Control" & io_comp == "T3 Vs T1"))|
                                                      author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                                                         
                 "Testosterone")


# Free TT
forest_func_pool(metagen_func_pool(tt_free_meta ), "Free Testosterone")

forest_func_pool(tt_free_pooled <- metagen_func_pool(tt_free_meta |> 
                                                  filter(
                                                    ((io_new_cat =="Non-serous OC VS Control"|
                                                            #io_new_cat =="Serous OC VS Control"|
                                                            io_new_cat== "OC VS Control" )&
                                                           !(io_new_cat == "OC VS Control" & io_comp == "T3 Vs T1"))|
                                                      author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                 "Free Testosterone")

#Estrone
forest_func_pool(metagen_func_pool(e1_meta ), "Estrone")

forest_func_pool(e1_pooled <- metagen_func_pool(e1_meta |> 
                                                       filter(
                                                         ((io_new_cat== "OC VS Control" )&
                                                            !(io_new_cat == "OC VS Control" & io_adj == "Unadjusted"))|#Only one paper
                                                           author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                 "Estrone")


#A1
#forest_func_pool(metagen_func_pool(a1_meta), "Androstenedione")

forest_func_pool(a1_pooled <- metagen_func_pool(a1_meta |> 
                                                  filter(
                                                    ((io_new_cat =="Non-serous OC VS Control"|
                                                    io_new_cat =="Serous OC VS Control"|
                                                    io_new_cat== "OC VS Control" )&
                                                      !(io_new_cat == "Serous OC VS Control" & io_comp == "Continuous")&
                                                      !(io_new_cat == "OC VS Control" & io_comp == "Continuous"))|
                                                      author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                 "Androstenedione")

#Progestrone
forest_func_pool(prog_pooled <- metagen_func_pool(prog_meta |> filter(
  ((io_new_cat =="Non-serous OC VS Control"|
      io_new_cat =="Serous OC VS Control"|
      io_new_cat== "OC VS Control" )&
     !(io_new_cat =="OC VS Control"& io_comp =="Continuous"))|
    author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))), 
                 "Progestrone")


# SHBG
forest_func_pool(metagen_func_pool(shbg_meta), "SHBG")

forest_func_pool(shbg_pooled <- metagen_func_pool(shbg_meta |> 
                                                  filter(
                                                    ((io_new_cat =="Non-serous OC VS Control"|
                                                       io_new_cat =="Serous OC VS Control"|
                                                       io_new_cat== "OC VS Control" )&
                                                      !(io_new_cat == "Serous OC VS Control" & io_comp == "Continuous")&
                                                      !(io_new_cat == "OC VS Control" & io_comp == "T3 Vs T1")&
                                                      !(io_new_cat == "Serous OC VS Control"& io_comp == "T3 Vs T1"))|
                                                      author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                 "SHBG")

#IGF1
forest_func_pool(metagen_func_pool(igf1_meta), "IGF-1")

forest_func_pool(igf1_pooled <- metagen_func_pool(igf1_meta |> 
                                                    filter(
                                                      ((io_new_cat =="Non-serous OC VS Control"|
                                                         io_new_cat== "OC VS Control" )&
                                                        !(io_new_cat == "OC VS Control" & io_adj == "Multivariable adjusted + IGFBP-2 & IGFBP-3")&
                                                        !(io_new_cat == "OC VS Control"& io_adj == "Minimally adjusted"))|
                                                        author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))),
                 "IGF-1")

#IGFBP-2
forest_func_pool(metagen_func_pool(igfbp2_meta), "IGFBP-2")
forest_func_pool(metagen_func_pool(igfbp2_meta|>filter(
  ((io_new_cat== "OC VS Control")|
     !(io_new_cat=="OC VS Control"& io_adj=="Unadjusted"))|
    author_year %in% c("Peres, L 2019", "Ose, J et.al 2017", "Ose, J 2017"))), 
                 "IGFBP-2")

#IL-8
forest_func_pool(il8_pooled <- metagen_func_pool(il8_meta), "IL-8")



extract_meta_analysis <- function(meta_analysis, meta_name) {
  # Extract the desired values
  Subgroup <- meta_analysis$subgroup.levels
  No_studies <- meta_analysis$k.w
  Studies <- vector("list", length(meta_analysis$k.w))
  for (i in 1:length(meta_analysis$k.w)) {
    Studies[[i]] <- list(meta_analysis$studlab[meta_analysis$k.w == meta_analysis$k.w[i]])
  }  
  OR <- round(exp(meta_analysis$TE.random.w), 2)
  lower <- round(exp(meta_analysis$lower.random.w), 2)
  upper <- round(exp(meta_analysis$upper.random.w), 2)
  I2 <- paste(round(meta_analysis$I2.w, 2) * 100)
  library(purrr)
  
  # Combine the Subgroup and Studies variables
  Subgroup_Studies <- map2_chr(Subgroup, Studies, ~paste(.x, .y, collapse = ", "))
  
  # Combine the values into a data frame and add a column for the meta-analysis name
  data.frame(
    Subgroup = Subgroup,
    Meta_analysis = meta_name,
    `No of studies` = No_studies,
    Studies = Subgroup_Studies,
    OR = OR,Lower = lower, Upper = upper,
    CI = paste("[", lower, ",", upper, "]", sep = ""),
    I2 = paste(I2, "%", sep = "")
  )
}

# Apply the function to each meta-analysis and combine the results
meta_results <- list(
  extract_meta_analysis(e2_pooled, "Estradiol"),
  extract_meta_analysis(unconj_e2_pooled, "Unconjugated Estradiol"),
  extract_meta_analysis(e1_pooled, "Estrone"),
  extract_meta_analysis(shbg_pooled, "SHBG"),
  extract_meta_analysis(a1_pooled, "Androstenedione"),
  extract_meta_analysis(tt_pooled, "Testosterone"),
  extract_meta_analysis(tt_free_pooled, "Free testosterone"),
  extract_meta_analysis(dheas_pooled, "DHEAS"),
  extract_meta_analysis(dhea_pooled, "DHEA"),
  extract_meta_analysis(prog_pooled, "Progestrone"),
  extract_meta_analysis(il8_pooled, "IL-8"),
  extract_meta_analysis(igf1_pooled, "IGF-1"),
  extract_meta_analysis(igfbp2_pooled, "IGFB-2")
)

#
meta_results_df <- do.call(rbind, meta_results)

meta_results_df$new_subgroup <- meta_results_df$Subgroup
library(tidyr)
meta_results_df <- extract(meta_results_df, new_subgroup, 
                  into = c("biomarker","comparison", "oc_type","adjustment" ), 
                  regex =  "^([^|]*)\\s*\\|\\s*([^|]*)\\s*\\|\\s*([^|]*)\\s*\\|\\s*(.*)$")

#Or
# meta_results_df <- separate(meta_results_df, new_subgroup, 
#                             into = c("biomarker", "comparison", "oc_type", "adjustment"), 
#                             sep = "\\s*\\|\\s*", 
#                             remove = TRUE, 
#                             convert = FALSE)
meta_results_df$oc_type <- ifelse(meta_results_df$oc_type=="OC VS Control",
                          "All OC", meta_results_df$oc_type)

meta_results_df <- meta_results_df |> arrange(oc_type)
meta_results_df$type <- ifelse(meta_results_df$No.of.studies==1, "OC3 pooled analysis",
                                "Meta-analysis")

meta_results_df$I2 <- ifelse(meta_results_df$I2=="NA%", "-",
                             meta_results_df$I2)

#Replace the string Testosterone_free with Free Testosterone
meta_results_df$Subgroup <- str_replace(meta_results_df$Subgroup,
                                             "Testosterone_free", "Free Testosterone")

#Save figure
svg("U:/Studien/HOC/WCRF CUP/15_Statistical_analysis_Final_summary/5.OC/Blood/Mihiretu/all_meta_analysis.svg", 
    width = 10, height = 14)

forest(
  metagen(TE=log(OR), lower=log(Lower), upper = log(Upper), sm="OR", 
          data=meta_results_df |> 
            filter(#No.of.studies > 1&
               Meta_analysis != "Unconjugated Estradiol")|>
            mutate(Subgroup = str_remove_all(Subgroup, 
                                             "Multivariable adjusted|VS Control \\|"),
                   Subgroup = str_replace(Subgroup, "\\| OC", "| All OC"),
                   ),
          overall = F,fixed=F, random=F,
          subgroup = Meta_analysis ),
  subgroup.name  = "", fontsize=8,spacing = 0.7,
  backtransf = T,
  leftcols = "Subgroup",
  leftlabs = "Subgroup",
  rightcols = c("OR","CI", "No.of.studies","I2", "type"#, "Studies"
  ),
  rightlabs = c("OR", "95%CI", 
                "Number of studies","I2", "Source"#, "Studies"
  ),
  squaresize = 0.6,
  col.subgroup = "#00868B"
)
dev.off()

#Now call the Mendelian randomization code

source("U:/Studien/HOC/WCRF CUP/15_Statistical_analysis_Final_summary/5.OC/Blood/Mihiretu/MR_studies_analysis_OC.R")

mr_oc$type <- "MR"

names(mr_oc)
mr_oc$No.of.studies <- mr_oc$author_year
mr_oc$Subgroup <- paste(mr_oc$i," | All OC" )
ivw <- mr_oc |> 
  filter(mr_stat=="IVW")

meta_results <- meta_results_df |> 
  filter(oc_type=="OC VS Control ")


table(meta_results$oc_type)

names(meta_results)

ivw <- ivw |> 
  mutate(
    I2 = explained_variance_percent,
    OR = io_statv,
    cohort = source,
    bias = rob,
    Lower = io_lower,
    Upper = io_upper,
    biomarker = i
  )

names(oc_new)

table(ivw$i)
ivw$type <- paste(ivw$type," - ", ivw$cohort)
ivw <- ivw |> 
  select(Subgroup, type, OR, Lower, Upper,
        I2,No.of.studies, type, no_gen_instrument, biomarker )

meta_results$no_gen_instrument <- ""

meta_analysis_df <- meta_results |> 
  select(Subgroup,type, OR, Lower, Upper, 
        I2, No.of.studies,no_gen_instrument, biomarker)
meta_analysis_df$No.of.studies <- as.character(meta_analysis_df$No.of.studies)

is.numeric(ivw$I2) #Explained variance should be charcter other wise it will not 

ivw$I2 <- paste(ivw$I2, "%")
ivw$I2 <- ifelse(ivw$I2=="NA %", "-", ivw$I2)

mr_met_oc <- rbind(meta_analysis_df, ivw)


oc_new |> filter(
  author_year %in% c("Peres, L 2019", 
                     "Ose, J et.al 2017", 
                     "Ose, J 2017")) |> count(author_year, 
                                              i)
# Now label the studies 

mr_met_oc$No.of.studies <- ifelse(mr_met_oc$type=="OC3 pooled analysis", NA, 
                                  mr_met_oc$No.of.studies)
sum(is.na(mr_met_oc$No.of.studies))

#This doesn't work: WHy? No idea!
# mr_met_oc <- mr_met_oc |> 
#   mutate(No.of.studies = case_when(
#     is.na(No.of.studies) & i == "Androstendione" ~ "Ose, J 2017",
#     is.na(No.of.studies) & i == "DHEAS" ~ "Ose, J 2017",
#     is.na(No.of.studies) & i == "SHBG" ~ "Ose, J 2017",
#     is.na(No.of.studies) & i == "Testosterone_free"~ "Ose, J 2017",
#     is.na(No.of.studies) & i == "IGF-1" ~ "Ose, J et.al 2017",
#     is.na(No.of.studies) & i == "C_reactive_protein_CRP" ~ "Peres, L 2019",
#     TRUE ~ No.of.studies
#   ))

# Create a named vector for mapping
author_map <- c("Androstendione" = "Ose, J 2017", 
                "DHEAS" = "Ose, J 2017", 
                "SHBG" = "Ose, J 2017", 
                "Testosterone_free" = "Ose, J 2017", 
                "IGF-1" = "Ose, J et.al 2017", 
                "C_reactive_protein_CRP" = "Peres, L 2019")

#mutate with ifelse and the mapping
mr_met_oc <- mr_met_oc |> 
  mutate(No.of.studies = ifelse(is.na(No.of.studies), 
                                author_map[i], No.of.studies))

#Now plot
mr_met_oc |> count(biomarker, No.of.studies, type)
mr_met_oc$OR_95_CI <- paste(mr_met_oc$OR, 
                            "[",mr_met_oc$Lower, mr_met_oc$Upper,"]", 
                            sep ="")

svg("U:/Studien/HOC/WCRF CUP/15_Statistical_analysis_Final_summary/5.OC/Blood/Mihiretu/mr_met_triangulation_allOC.svg", 
    width = 12, height = 14)
forest(
  metagen(TE=log(OR), lower=log(Lower), upper = log(Upper), sm="OR", 
          data=mr_met_oc |> arrange(biomarker) |> 
            mutate(Subgroup = str_remove_all(Subgroup, 
                                             "Multivariable adjusted|VS Control \\|")),
          
          overall = F,fixed=F, random=F,
          subgroup = biomarker ),
  subgroup.name  = "", fontsize=8,spacing = 0.8,
  backtransf = T,
  leftcols = "Subgroup",
  leftlabs = "Subgroup",
  rightcols = c("OR_95_CI","No.of.studies","I2","type"),
  rightlabs = c("OR[95%CI]", "Number of studies",
                "I-squared/Explained Variance", "Source"),
  squaresize = 0.6,
  col.subgroup = "#00868B"
)
dev.off()


