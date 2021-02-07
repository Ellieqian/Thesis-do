*=========================================================
* Thesis .do file *
*=========================================================

clear
clear matrix
clear mata
capture log close
set maxvar 15000
set more off
numlabel, add


*=========================================================
* Set up macros and clean dataset *
*=========================================================

* Change directory and make sub-folder
global datadir "/Users/Ellie/Desktop/THESIS/Data"
global dofiledir "/Users/Ellie/Desktop/THESIS/Thesis-do"
global output "/Users/Ellie/Desktop/THESIS/Output"

*Year Macros
local COHORT Cohort1 

* Set local/global macros for current date
local today=c(current_date)
local c_today= "`today'"
global date=subinstr("`c_today'", " ", "",.)

* Run .do file to merge baseline and 6-week data 
//run "$dofiledir/6Week_Merge.do"

* Load data
use "/Users/Ellie/Desktop/THESIS/Data/Cohort1_6W_Merged_30Jan2021.dta", clear

* Create log file
cd $output
capture log close
log using "Thesis.log", replace

* Generate strata
drop strata
egen strata=concat(region ur) if region==1 | region==3 | region==4 | region==7, punc(-)
replace strata=string(region) if strata==""

* Create weight for all women with complete forms
gen SWweight=FQweight if baseline_status==3
replace SWweight=SWFUweight if SWresult!=.

* Overall response rate 
replace SWresult=FRS_result if baseline_status==3 // replace with baseline result for 5-9 weeks postpartum women 

gen sw_responserate=0 if SWresult>=1 & SWresult<6 & SWresult!=.
replace sw_responserate=1 if SWresult==1 
label define responselist 0 "Not complete" 1 "Complete"
label val sw_responserate responselist
tab sw_responserate
tabout sw_responserate using "Thesis_output_$date.xls", replace cells(freq col) h2("6W response rate") f(0 1) clab(n %)

* Source of the data 
gen source=1 if FRS_result!=.
replace source=2 if SWresult_orig!=.
replace source=3 if SWresult_cc!=.
label var source "Source of 6-week data"
label define source_list 1 "Baseline (5-9 wks pp at baseline)" 2 "6-Week pre-COVID" 3 "6-Week during COVID"
label val source source_list

tabout source if FRS_result==1 | SWresult==1 using "Thesis_output_$date.xls", append cells(freq col) h2("Source of 6-Week Data-Complete forms") f(0 1) clab(n %) 

* Restrict analysis to women who completed questionnaire 
keep if SWresult==1

* Set survey weights
svyset EA [pweight=SWweight], strata(strata) singleunit(scaled)

* Generate 0/1 urban/rural variable
gen urban=ur==1
label variable urban "Urban/rural place of residence"
label define urban 1 "Urban" 0 "Rural"
label value urban urban

*=========================================================
* Covariates set up *
*=========================================================

* Replace 6w data with baseline for 5-9 weeks postpartum women
foreach var in delivery_place who_assisted_delivery anc_delivery_place anc_delivery_skilled anc_emergency_transport ///
				anc_danger_place anc_danger_migraine anc_danger_hbp anc_danger_edema ///
				anc_danger_convuls anc_danger_bleeding anc_nd_info_yn ///
				anc_bp anc_weight anc_urine anc_blood anc_stool anc_syph_test ///
				anc_syph_result anc_syph_couns anc_hiv_test anc_hiv_result anc_hiv_couns ///
				delivprob_bleed delivprob_leakmemb24hr delivprob_leakmembpre9mo ///
				delivprob_malposition delivprob_prolonglab delivprob_convuls anc_tt_inject{
	replace SW`var'= `var' if baseline_status==3
	}
	
replace SWanc_hew_yn=anc_hew_yn_pp if baseline_status==3
replace SWanc_phcp_yn=anc_phcp_yn_pp if baseline_status==3
replace SWanc_hew_num=anc_hew_num_pp if baseline_status==3
replace SWanc_phcp_num=anc_phcp_num_pp if baseline_status==3

foreach prob in migraine hbp edema convuls vagbleed fever abnormdisch abpain vision {

	replace SWpregprob_`prob' = pregprob_`prob'_pp if baseline_status==3

}
tab SWpregprob_hbp, m

* Recode -99 and -88 to missing
foreach var in age school total_birth pregnancy_desired marital_status SWanc_hew_timing SWanc_phcp_timing ///
				SWanc_hew_yn SWanc_phcp_yn SWanc_hew_num SWanc_phcp_num ///
				SWanc_delivery_place SWanc_delivery_skilled SWanc_emergency_transport ///
				SWanc_danger_place SWanc_danger_migraine SWanc_danger_hbp SWanc_danger_edema ///
				SWanc_danger_convuls SWanc_danger_bleeding SWanc_nd_info_nr ///
				SWanc_bp SWanc_weight SWanc_urine SWanc_blood SWanc_stool SWanc_syph_test ///
				SWanc_syph_result SWanc_syph_couns SWanc_hiv_test SWanc_hiv_result SWanc_hiv_couns ///
				SWanc_lam_couns SWanc_ppfp_couns SWwho_assisted_delivery ///
				SWdelivprob_bleed SWdelivprob_leakmemb24hr SWdelivprob_leakmembpre9mo ///
				SWdelivprob_malposition SWdelivprob_prolonglab SWdelivprob_convuls ///
				SWpostdelivprob_retainpl SWpostdelivprob_fever SWpostdelivprob_bleed ///
				SWpostdelivprob_convuls SWpregprob_hbp SWpregprob_edema SWpregprob_convuls ///
				SWpregprob_vagbleed SWpregprob_fever SWpregprob_abnormdisch SWpregprob_abpain SWpregprob_vision {
	recode `var' (-99 -88 =.) 
	}
	
* Label yes/no response options
capture label define yesno 0 "No" 1 "Yes"	

* Generate age categories
gen age_cat=0
replace age_cat=1 if FQ_age>=20 & FQ_age<35
replace age_cat=2 if FQ_age>=35
label define agel 0 "<20" 1 "20-34" 2 "35+"
label value age_cat agel

* Generate categorical variable for parity 
egen parity4=cut(total_births), at(0, 1, 2, 4, 30) icodes
lab def parity4l 0 "0 children" 1 "1 children" 2 "2-3 children" 3 "4+ children"
lab val parity4 parity4l
lab var parity4 "Parity" 

* Group secondary, technical & vocational, and higher education 
gen education=school
replace education=2 if school>=3
lab def edul 0 "No education" 1 "Primary" 2 "Secondary+" 
lab val education edul
lab var education "Education level" 

gen any_edu=0 if education==0
replace any_edu=1 if education==1 | education==2

* Generate unintended pregnancy binary variable
gen unintended_preg=0
replace unintended_preg=1 if pregnancy_desired==2 | pregnancy_desired==3 
replace unintended_preg=. if pregnancy_desired==.
lab var unintended_preg "Unintended pregnancy" 
label val unintended_preg yesno

* Generate binary marital status variable 
gen married=0 if marital_status!=.
replace married=1 if marital_status==1 | marital_status==2
label define marriedl 0 "Not married" 1 "Married or living with a partner"
label val married marriedl

* Facility versus home delivery
gen facility_deliv=0 if SWdelivery_place==1 | SWdelivery_place==2
replace facility_deliv=1 if SWdelivery_place>2 
replace facility_deliv=. if SWdelivery_place==. // 96 treated as facility delivery 
label val facility_deliv yesno
tab facility_deliv, m

* Skilled birth attendant binary variable
gen sba=0 if SWwho_assisted_delivery==0 | SWwho_assisted_delivery>7 // other treated as no sba
replace sba=1 if SWwho_assisted_delivery<=7
replace sba=. if SWwho_assisted_delivery==.
label val sba yesno
tab sba, m

*=========================================================
* ANC-related covariates *
*=========================================================

/*
		NOTE: 
			1. "0" include women with no ANC 
			2. All the variables that did not come directly from the questionnaire do not have any missing
				-88, -99, and . are all treated as 0 in generated variables 
*/

*** ANC frequency ***
* Total number of ANC 
replace SWanc_hew_num=0 if SWanc_hew_num==.
replace SWanc_phcp_num=0 if SWanc_phcp_num==.
gen anc_tot=SWanc_hew_num + SWanc_phcp_num 
label var anc_tot "Total number of ANC" 

* 0, 1-3 and 4+ ANC categories 
gen anc_num_cat=0 if anc_tot==0
replace anc_num_cat=1 if anc_tot>=1 & anc_tot<=3
replace anc_num_cat=2 if anc_tot>=4
label define anc_num_l 0 "No ANC" 1 "1-3 ANC" 2 "4+ ANC"
label val anc_num_cat anc_num_l

* Any ANC
gen anyanc=0 if anc_tot==0
replace anyanc=1 if anc_tot!=0 & anc_tot!=.
label var anyanc "Had at least one ANC" 
label val anyanc yesno

* 4+ ANC
gen ANC4=0
replace ANC4=1 if anc_num_cat==2
label var ANC4 "Had mroe than 4 ANC" 

* 8+ ANC
gen ANC8=0
replace ANC8=1 if anc_tot>=8
label var ANC8 "Had mroe than 8 ANC" 
label val ANC8 yesno 
tab ANC8

*** GA at first ANC visit ***
* Generate GA at first visit variable
recode SWanc_phcp_timing(0 = 1)
recode SWanc_hew_timing(0 = 1)
gen ga_first_anc=SWanc_hew_timing
replace ga_first_anc=SWanc_phcp_timing if SWanc_phcp_timing<anc_hew_timing & SWanc_phcp_timing!=.
replace ga_first_anc=. if SWanc_hew_timing==. & SWanc_phcp_timing==.
replace ga_first_anc=. if anyanc==0
label var ga_first_anc "GA (in months) at first ANC"

* Create categories 
gen ga_first_anc_cat=0
replace ga_first_anc_cat=1 if ga_first_anc<4 & ga_first_anc>=1 
replace ga_first_anc_cat=2 if ga_first_anc>=4 & ga_first_anc<=6 
replace ga_first_anc_cat=3 if ga_first_anc>6 & ga_first_anc<=9 
replace ga_first_anc_cat=4 if ga_first_anc==.
replace ga_first_anc_cat=0 if anyanc==0
label define ga_anc_l 0 "No ANC" 1 "First trimester" 2 "Second trimester" 3 "Third trimester" 4 "Missing ANC timing"
label val ga_first_anc_cat ga_anc_l
tab ga_first_anc_cat

* ANC by provider type
gen provider_code=0 if anyanc==0
replace provider_code=1 if SWanc_hew_yn==1 & SWanc_phcp_yn!=1 & anyanc==1 
replace provider_code=2  if SWanc_phcp_yn==1 & SWanc_hew_yn!=1 & anyanc==1
replace provider_code=3 if SWanc_phcp_yn==1 & SWanc_hew_yn==1 & anyanc==1
label define providerl 0 "No ANC" 1 "HEW only" 2 "PHCP only" 3 "Both" 
label val provider_code providerl

* Generate binary composite birth/complication readiness indicator
gen birth_readiness_all=0
replace birth_readiness_all=1 if (SWanc_delivery_place==1 & SWanc_delivery_skilled==1 & ///
									SWanc_emergency_transport==1 & SWanc_danger_place==1 & ///
									SWanc_danger_migraine==1 & SWanc_danger_hbp==1 & SWanc_danger_edema==1 & ///
									SWanc_danger_convuls==1 & SWanc_danger_bleeding==1)
label var birth_readiness_all "At ANC discussed all birth/complication readiness topics"
label val birth_readiness_all yesno

/*
continuous variable, then categorize  
look up egen xtile to divide into three even groups
*/

gen birth_readiness_c=SWanc_delivery_place + SWanc_delivery_skilled + SWanc_emergency_transport + SWanc_danger_place + ///
						SWanc_danger_migraine + SWanc_danger_hbp + SWanc_danger_edema + SWanc_danger_convuls + SWanc_danger_bleeding
								
gen birth_readiness_cat=0
replace birth_readiness_cat=1 if birth_readiness_c<=3 & birth_readiness_c>=1 
replace birth_readiness_cat=2 if birth_readiness_c>3 & birth_readiness_c<=6 
replace birth_readiness_cat=3 if birth_readiness_c>6 & birth_readiness_c<=9 
replace birth_readiness_cat=0 if anyanc==0
label var birth_readiness_cat "Birth readiness discussion (categorial)"
label define birthl 0 "None" 1 "1-3 topics" 2 "4-6 topics" 3 "7-9 topics"
label val birth_readiness_cat birthl

* PE/E danger sign counseling
gen danger_sign_coun=0
replace danger_sign_coun=1 if (SWanc_danger_migraine==1 | anc_danger_hbp==1 | anc_danger_edema==1 | SWanc_danger_convuls==1)
label var danger_sign_coun "Have at least one PE/E danger sign counseling at ANC"
label val danger_sign_coun yesno

* Key ANC services binary indicators 
gen anc_key_services=0 
replace anc_key_services=1 if (SWanc_bp==1 & SWanc_iron==1 & SWanc_blood==1 & ///
								SWanc_urine==1 & SWanc_syph_test==1 & SWanc_hiv_test==1) 
label var anc_key_services "At ANC had : BP, urine and blood sampled, tested for syphilis & HIV, + took iron during preg"
label val anc_key_services yesno

* Maternal assessments 
/*foreach var in bp weight urine blood stool {
	recode SWanc_`var'(. = 0)
	}
*/
gen maternal_assess_score=SWanc_bp + SWanc_weight + SWanc_urine + SWanc_blood + SWanc_stool
replace maternal_assess_score=0 if anyanc==0

gen maternal_assess_cat=0 
replace maternal_assess_cat=1 if maternal_assess_score>=1 & maternal_assess_score<=3
replace maternal_assess_cat=2 if maternal_assess_score==4
replace maternal_assess_cat=3 if maternal_assess_score==5
replace maternal_assess_cat=0 if anyanc==0
label define maternal_assess_cat 0 "None" 1 "1-3 assessments" 2 "4 assessments" 3 "All 5 assessments"
label val maternal_assess_cat maternal_assess_cat

* All 5 assessments
gen maternal_assess_all=0
replace maternal_assess_all=1 if (SWanc_bp==1 & SWanc_weight==1 & SWanc_urine==1 & SWanc_blood==1 & SWanc_stool==1)
label var maternal_assess_all "Received all 5 maternal assessments"
label val maternal_assess_all yesno
tab maternal_assess_all

* PE/E-related Pregnancy complications binary variable 
gen preg_comp=0
replace preg_comp=1 if (SWpregprob_hbp==1 | SWpregprob_edema==1 | SWpregprob_convuls==1 | SWpregprob_vision==1) 
label var preg_comp "Had PE/E-related complications during pregnancy" 
label val preg_comp yesno

* Recode missing variable to 0 for women with no ANC 

foreach var in SWanc_delivery_place SWanc_delivery_skilled SWanc_emergency_transport ///
				SWanc_danger_place SWanc_danger_migraine SWanc_danger_hbp SWanc_danger_edema ///
				SWanc_danger_convuls SWanc_danger_bleeding SWanc_nd_info_yn ///
				SWanc_bp SWanc_weight SWanc_urine SWanc_blood SWanc_stool ///
				SWanc_syph_test  SWanc_hiv_test SWanc_tt_inject {
				
				recode `var' (. -88 -99 = 0)
				
				}

* Facility delivery and skilled birth attendant combined variable
gen facility_skilled=0 if facility_deliv!=. & sba!=.
replace facility_skilled=1 if (facility_deliv==0 & sba==0)
replace facility_skilled=2 if (facility_deliv==0 & sba==1)
replace facility_skilled=3 if facility_deliv==1
label define facility_skilledl 1 "Home delivery with no SBA" 2 "Home delivery with SBA" 3 "Facility delivery with SBA"
label val facility_skilled facility_skilledl
tab facility_skilled, m

*=========================================================
* Describe missingness *
*=========================================================
			
preserve 
keep age urban school wealthquintile total_births parity4 pregnancy_desired unintended_preg ///
				SWdelivery_place SWwho_assisted_delivery ///
				SWanc_hew_yn SWanc_phcp_yn SWanc_hew_yn SWanc_phcp_yn /// 
				SWanc_hew_num SWanc_hew_num SWanc_phcp_num SWanc_phcp_num ///
				SWanc_delivery_place SWanc_delivery_skilled SWanc_emergency_transport ///
				SWanc_danger_place SWanc_danger_migraine SWanc_danger_hbp SWanc_danger_edema ///
				SWanc_danger_convuls SWanc_danger_bleeding SWanc_nd_info_yn ///
				SWanc_bp SWanc_weight SWanc_urine SWanc_blood SWanc_stool ///
				SWanc_syph_test  SWanc_hiv_test ///
				SWpregprob_hbp SWpregprob_edema SWpregprob_convuls SWpregprob_vision ///
				SWdelivprob_convuls
				
missings report, percent
restore 

*=========================================================
* EDA *
*=========================================================

* Tabulation of outcome and potential confounders
* Background characteristics
tabout SWdelivprob_convuls age_cat [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls education [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls urban [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls married [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls parity4 [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls unintended_preg [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 


* ANC and delivery care 
tabout SWdelivprob_convuls anyanc [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls ANC4 [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls anc_key_services [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls birth_readiness_all [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls preg_comp [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls facility_deliv [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls sba [aw=SWweight]  using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls SWcaesarean_delivery [aw=SWweight]  using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 


* Lowess smoothed regressions -- continuous Xs
/*
* Age
lowess SWdelivprob_convuls FQ_age, name(age,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds age -vs- any complication")
graph export graphs/lowess_age.png,replace

* Parity
lowess SWdelivprob_convuls total_birth, name(parity,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds parity -vs- any complication")
graph export graphs/lowess_parity.png,replace

* Number of ANC 
lowess SWdelivprob_convuls anc_tot, name(anc,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds ANC -vs- any complication")
graph export graphs/lowess_anc.png,replace

* GA at first ANC
lowess SWdelivprob_convuls ga_first_anc if ga_first_anc!=0, name(ga_anc,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds GA at first ANC -vs- any complication")
graph export graphs/lowess_ga.png,replace

* k x 2 tables -- categorical Xs
for var age_cat:  tabulate X SWdelivprob_convuls, chi2 exact column
for var parity4:  tabulate X SWdelivprob_convuls, chi2 exact column
for var education:  tabulate X SWdelivprob_convuls, chi2 exact column
for var anc_num_cat:  tabulate X SWdelivprob_convuls, chi2 exact column


graph box FQ_age, name(age_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
	marker(1,mlab(FQ_age)) t1("Age by provider type") scale(1.2) ytitle("Age (years)") 
	graph export graphs/age_provider.png, replace

graph box total_birth, name(parity_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
	marker(1,mlab(total_birth)) t1("Parity by provider type") scale(1.2) ytitle("Number of children") 
	graph export graphs/parity_provider.png, replace
*/


*=========================================================
* Table 1 *
*=========================================================

* Run the biggest model to keep sample size consistent 
quietly svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_tt_inject i.anc_num_cat 
gen select = e(sample)

* By outcome
foreach var in age_cat urban education wealthquintile married parity4 unintended_preg preg_comp facility_skilled ///
				anc_num_cat provider_code maternal_assess_all SWanc_bp SWanc_weight SWanc_urine ///
				birth_readiness_all danger_sign_coun SWanc_nd_info_yn SWanc_tt_inject{

	svy: tab `var' SWdelivprob_convuls if select, col 
	
}

*=========================================================
* Logistic regression *
*=========================================================

*** Crude effect among all women ***

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) replace
putexcel A1 = "ANC Frequency", bold overwritefmt
putexcel E2 = "N = ", right bold overwritefmt
putexcel F2 = `e(N)', hcenter bold overwritefmt
putexcel A3 = matrix(results), names nformat(number_d2)

* ANC provider type
svy: logistic SWdelivprob_convuls i.provider_code if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = HEW PHCP both _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A9 = "ANC Provider Type", bold overwritefmt
putexcel E10 = "N = ", right bold overwritefmt
putexcel F10 = `e(N)', hcenter bold overwritefmt
putexcel A11 = matrix(results), names nformat(number_d2)

* Maternal assessment at ANC 
svy: logistic SWdelivprob_convuls maternal_assess_all if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A17 = "Received all 5 assessment at ANC", bold overwritefmt
putexcel E18 = "N = ", right bold overwritefmt
putexcel F18 = `e(N)', hcenter bold overwritefmt
putexcel A19 = matrix(results), names nformat(number_d2)

svy: logistic SWdelivprob_convuls i.maternal_assess_cat if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A23 = "Maternal assessment score at ANC", bold overwritefmt
putexcel E24 = "N = ", right bold overwritefmt
putexcel F24 = `e(N)', hcenter bold overwritefmt
putexcel A25 = matrix(results), names nformat(number_d2)

* BP measurement alone
svy: logistic SWdelivprob_convuls SWanc_bp if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A31 = "BP measurement at ANC", bold overwritefmt
putexcel E32 = "N = ", right bold overwritefmt
putexcel F32 = `e(N)', hcenter bold overwritefmt
putexcel A33 = matrix(results), names nformat(number_d2)

* Weight measurement alone
svy: logistic SWdelivprob_convuls SWanc_weight if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A37 = "Weight measurement at ANC", bold overwritefmt
putexcel E38 = "N = ", right bold overwritefmt
putexcel F38 = `e(N)', hcenter bold overwritefmt
putexcel A39 = matrix(results), names nformat(number_d2)

* Urine test alone
svy: logistic SWdelivprob_convuls SWanc_urine if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A43 = "Urine test at ANC", bold overwritefmt
putexcel E44 = "N = ", right bold overwritefmt
putexcel F44 = `e(N)', hcenter bold overwritefmt
putexcel A45 = matrix(results), names nformat(number_d2)

* Birth readiness discussion 
* Categorical 
svy: logistic SWdelivprob_convuls i.birth_readiness_cat if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4-6 7-9 _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A49 = "Birth readiness discussion (categorial)", bold overwritefmt
putexcel E50 = "N = ", right bold overwritefmt
putexcel F50 = `e(N)', hcenter bold overwritefmt
putexcel A51 = matrix(results), names nformat(number_d2)

* Binary 
svy: logistic SWdelivprob_convuls birth_readiness_all if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A57 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E58 = "N = ", right bold overwritefmt
putexcel F58 = `e(N)', hcenter bold overwritefmt
putexcel A59 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls danger_sign_coun if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A63 = "PE-related danger sign counseling", bold overwritefmt
putexcel E64 = "N = ", right bold overwritefmt
putexcel F64 = `e(N)', hcenter bold overwritefmt
putexcel A65 = matrix(results), names nformat(number_d2)

* Nutrition counseling
svy: logistic SWdelivprob_convuls SWanc_nd_info_yn if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A69 = "Nutritional counseling", bold overwritefmt
putexcel E70 = "N = ", right bold overwritefmt
putexcel F70 = `e(N)', hcenter bold overwritefmt
putexcel A71 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls SWanc_tt_inject if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A75 = "Tetanus injection at ANC", bold overwritefmt
putexcel E76 = "N = ", right bold overwritefmt
putexcel F76 = `e(N)', hcenter bold overwritefmt
putexcel A77 = matrix(results), names nformat(number_d2)

* Facility deliver and SBA 
svy: logistic SWdelivprob_convuls i.facility_skilled if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = Home_SBA Facility_SBA _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A81 = "Facility delivery and SBA", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82 = `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)

*** Crude effect among women with signs of preeclampsia ***

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A1 = "ANC Frequency", bold overwritefmt
putexcel E2 = "N = ", right bold overwritefmt
putexcel F2 = `e(N)', hcenter bold overwritefmt
putexcel A3 = matrix(results), names nformat(number_d2)

* ANC provider type
svy: logistic SWdelivprob_convuls i.provider_code if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = HEW PHCP both _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A8 = "ANC Provider Type", bold overwritefmt
putexcel E9 = "N = ", right bold overwritefmt
putexcel F9 = `e(N)', hcenter bold overwritefmt
putexcel A10 = matrix(results), names nformat(number_d2)

* Maternal assessment at ANC 
svy: logistic SWdelivprob_convuls maternal_assess_all if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A17 = "Received all 5 assessment at ANC", bold overwritefmt
putexcel E18 = "N = ", right bold overwritefmt
putexcel F18 = `e(N)', hcenter bold overwritefmt
putexcel A19 = matrix(results), names nformat(number_d2)

svy: logistic SWdelivprob_convuls i.maternal_assess_cat if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A23 = "Maternal assessment score at ANC", bold overwritefmt
putexcel E24 = "N = ", right bold overwritefmt
putexcel F24 = `e(N)', hcenter bold overwritefmt
putexcel A25 = matrix(results), names nformat(number_d2)

* BP measurement alone
svy: logistic SWdelivprob_convuls SWanc_bp if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A31 = "BP measurement at ANC", bold overwritefmt
putexcel E32 = "N = ", right bold overwritefmt
putexcel F32 = `e(N)', hcenter bold overwritefmt
putexcel A33 = matrix(results), names nformat(number_d2)

* Weight measurement alone
svy: logistic SWdelivprob_convuls SWanc_weight if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A37 = "Weight measurement at ANC", bold overwritefmt
putexcel E38 = "N = ", right bold overwritefmt
putexcel F38 = `e(N)', hcenter bold overwritefmt
putexcel A39 = matrix(results), names nformat(number_d2)

* Urine test alone
svy: logistic SWdelivprob_convuls SWanc_urine if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A43 = "Urine test at ANC", bold overwritefmt
putexcel E44 = "N = ", right bold overwritefmt
putexcel F44 = `e(N)', hcenter bold overwritefmt
putexcel A45 = matrix(results), names nformat(number_d2)

* Birth readiness discussion 
* Categorical 
svy: logistic SWdelivprob_convuls i.birth_readiness_cat if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A49 = "Birth readiness discussion (categorial)", bold overwritefmt
putexcel E50 = "N = ", right bold overwritefmt
putexcel F50 = `e(N)', hcenter bold overwritefmt
putexcel A51 = matrix(results), names nformat(number_d2)

* Binary 
svy: logistic SWdelivprob_convuls birth_readiness_all if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A57 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E58 = "N = ", right bold overwritefmt
putexcel F58 = `e(N)', hcenter bold overwritefmt
putexcel A59 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls danger_sign_coun if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A63 = "PE-related danger sign counseling", bold overwritefmt
putexcel E64 = "N = ", right bold overwritefmt
putexcel F64 = `e(N)', hcenter bold overwritefmt
putexcel A65 = matrix(results), names nformat(number_d2)

* Nutrition counseling
svy: logistic SWdelivprob_convuls SWanc_nd_info_yn if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A69 = "Nutritional counseling", bold overwritefmt
putexcel E70 = "N = ", right bold overwritefmt
putexcel F70 = `e(N)', hcenter bold overwritefmt
putexcel A71 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls SWanc_tt_inject if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
matrix rownames results = Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A75 = "Tetanus injection at ANC", bold overwritefmt
putexcel E76 = "N = ", right bold overwritefmt
putexcel F76 = `e(N)', hcenter bold overwritefmt
putexcel A77 = matrix(results), names nformat(number_d2)

* Facility deliver and SBA 
svy: logistic SWdelivprob_convuls i.facility_skilled if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = Home_SBA Facility_SBA _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A81 = "Facility delivery and SBA", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82 = `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)

*=====================================================================
* adjusted-all results *
*=====================================================================

*==================================*
* Among all women * 
*==================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.anc_num_cat if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A1 = "Adjusted logistic regression results among all women", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.provider_code if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A23 = "ANC Provider Type", bold overwritefmt
putexcel E24 = "N = ", right bold overwritefmt
putexcel F24 = `e(N)', hcenter bold overwritefmt
putexcel A25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.maternal_assess_cat if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel E45 = "N = ", right bold overwritefmt
putexcel F45 = `e(N)', hcenter bold overwritefmt
putexcel A46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled maternal_assess_all if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel E66 = "N = ", right bold overwritefmt
putexcel F66 = `e(N)', hcenter bold overwritefmt
putexcel A67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled danger_sign_coun if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A83 = "PE/E danger sign counseling", bold overwritefmt
putexcel E84 = "N = ", right bold overwritefmt
putexcel F84 = `e(N)', hcenter bold overwritefmt
putexcel A85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.birth_readiness_cat if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel E101 = "N = ", right bold overwritefmt
putexcel F101 = `e(N)', hcenter bold overwritefmt
putexcel A102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled birth_readiness_all if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E123 = "N = ", right bold overwritefmt
putexcel F123 = `e(N)', hcenter bold overwritefmt
putexcel A124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_nd_info_yn if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A140 = "Nutritional counseling", bold overwritefmt
putexcel E141 = "N = ", right bold overwritefmt
putexcel F141 = `e(N)', hcenter bold overwritefmt
putexcel A142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_tt_inject if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A158 = "Tetanus injection", bold overwritefmt
putexcel E159 = "N = ", right bold overwritefmt
putexcel F159 = `e(N)', hcenter bold overwritefmt
putexcel A160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_bp if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A176 = "BP measurement at ANC", bold overwritefmt
putexcel E177 = "N = ", right bold overwritefmt
putexcel F177 = `e(N)', hcenter bold overwritefmt
putexcel A178= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_weight if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A194 = "Weight measurement at ANC", bold overwritefmt
putexcel E195 = "N = ", right bold overwritefmt
putexcel F195 = `e(N)', hcenter bold overwritefmt
putexcel A196= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_urine if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A212 = "Urine test at ANC", bold overwritefmt
putexcel E213 = "N = ", right bold overwritefmt
putexcel F213 = `e(N)', hcenter bold overwritefmt
putexcel A214= matrix(results), names nformat(number_d2)

*======================================================*
* Among women who were pregnancy at enrollment  * 
*======================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.anc_num_cat if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H1 = "Adjusted logistic regression results among women who were pregnancy at enrollment", bold overwritefmt
putexcel H3 = "ANC Frequency", bold overwritefmt
putexcel L4 = "N = ", right bold overwritefmt
putexcel M4 = `e(N)', hcenter bold overwritefmt
putexcel H5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.provider_code if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H23 = "ANC Provider Type", bold overwritefmt
putexcel L24 = "N = ", right bold overwritefmt
putexcel M24 = `e(N)', hcenter bold overwritefmt
putexcel H25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.maternal_assess_cat if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel L45 = "N = ", right bold overwritefmt
putexcel M45 = `e(N)', hcenter bold overwritefmt
putexcel H46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled maternal_assess_all if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel L66 = "N = ", right bold overwritefmt
putexcel M66 = `e(N)', hcenter bold overwritefmt
putexcel H67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled danger_sign_coun if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H83 = "PE/E danger sign counseling", bold overwritefmt
putexcel L84 = "N = ", right bold overwritefmt
putexcel M84 = `e(N)', hcenter bold overwritefmt
putexcel H85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.birth_readiness_cat if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel L101 = "N = ", right bold overwritefmt
putexcel M101 = `e(N)', hcenter bold overwritefmt
putexcel H102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled birth_readiness_all if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel L123 = "N = ", right bold overwritefmt
putexcel M123 = `e(N)', hcenter bold overwritefmt
putexcel H124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_nd_info_yn if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H140 = "Nutritional counseling", bold overwritefmt
putexcel L141 = "N = ", right bold overwritefmt
putexcel M141 = `e(N)', hcenter bold overwritefmt
putexcel H142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_tt_inject if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H158 = "Tetanus injection", bold overwritefmt
putexcel L159 = "N = ", right bold overwritefmt
putexcel M159 = `e(N)', hcenter bold overwritefmt
putexcel H160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_bp if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H176 = "BP measurement at ANC", bold overwritefmt
putexcel L177 = "N = ", right bold overwritefmt
putexcel M177 = `e(N)', hcenter bold overwritefmt
putexcel H178= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_weight if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H194 = "Weight measurement at ANC", bold overwritefmt
putexcel L195 = "N = ", right bold overwritefmt
putexcel M195 = `e(N)', hcenter bold overwritefmt
putexcel H196= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_urine if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H212 = "Urine test at ANC", bold overwritefmt
putexcel L213 = "N = ", right bold overwritefmt
putexcel M213 = `e(N)', hcenter bold overwritefmt
putexcel H214= matrix(results), names nformat(number_d2)

*======================================================*
* Excluding 5-9 postpartum women * 
*======================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.anc_num_cat if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O1 = "Adjusted logistic regression results excluding 5-9 postpartum women", bold overwritefmt
putexcel O3 = "ANC Frequency", bold overwritefmt
putexcel S4 = "N = ", right bold overwritefmt
putexcel T4 = `e(N)', hcenter bold overwritefmt
putexcel O5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.provider_code if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O23 = "ANC Provider Type", bold overwritefmt
putexcel S24 = "N = ", right bold overwritefmt
putexcel T24 = `e(N)', hcenter bold overwritefmt
putexcel O25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.maternal_assess_cat if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel S45 = "N = ", right bold overwritefmt
putexcel T45 = `e(N)', hcenter bold overwritefmt
putexcel O46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled maternal_assess_all if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel S66 = "N = ", right bold overwritefmt
putexcel T66 = `e(N)', hcenter bold overwritefmt
putexcel O67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled danger_sign_coun if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O83 = "PE/E danger sign counseling", bold overwritefmt
putexcel S84 = "N = ", right bold overwritefmt
putexcel T84 = `e(N)', hcenter bold overwritefmt
putexcel O85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled i.birth_readiness_cat if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel S101 = "N = ", right bold overwritefmt
putexcel T101 = `e(N)', hcenter bold overwritefmt
putexcel O102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled birth_readiness_all if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel S123 = "N = ", right bold overwritefmt
putexcel T123 = `e(N)', hcenter bold overwritefmt
putexcel O124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_nd_info_yn if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O140 = "Nutritional counseling", bold overwritefmt
putexcel S141 = "N = ", right bold overwritefmt
putexcel T141 = `e(N)', hcenter bold overwritefmt
putexcel O142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_tt_inject if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O158 = "Tetanus injection", bold overwritefmt
putexcel S159 = "N = ", right bold overwritefmt
putexcel T159 = `e(N)', hcenter bold overwritefmt
putexcel O160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_bp if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O176 = "BP measurement at ANC", bold overwritefmt
putexcel S177 = "N = ", right bold overwritefmt
putexcel T177 = `e(N)', hcenter bold overwritefmt
putexcel O178= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_weight if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O194 = "Weight measurement at ANC", bold overwritefmt
putexcel S195 = "N = ", right bold overwritefmt
putexcel T195 = `e(N)', hcenter bold overwritefmt
putexcel O196= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 preg_comp i.facility_skilled SWanc_urine if select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O212 = "Urine test at ANC", bold overwritefmt
putexcel S213 = "N = ", right bold overwritefmt
putexcel T213 = `e(N)', hcenter bold overwritefmt
putexcel O214= matrix(results), names nformat(number_d2)

*=====================================================*
* Among all women with signs of preeclampsia * 
*=====================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.anc_num_cat if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A1 = "Adjusted logistic regression results among all women with signs preeclampsia", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.provider_code if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A23 = "ANC Provider Type", bold overwritefmt
putexcel E24 = "N = ", right bold overwritefmt
putexcel F24 = `e(N)', hcenter bold overwritefmt
putexcel A25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.maternal_assess_cat if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel E45 = "N = ", right bold overwritefmt
putexcel F45 = `e(N)', hcenter bold overwritefmt
putexcel A46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled maternal_assess_all if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel E66 = "N = ", right bold overwritefmt
putexcel F66 = `e(N)', hcenter bold overwritefmt
putexcel A67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled danger_sign_coun if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A83 = "PE/E danger sign counseling", bold overwritefmt
putexcel E84 = "N = ", right bold overwritefmt
putexcel F84 = `e(N)', hcenter bold overwritefmt
putexcel A85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.birth_readiness_cat if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel E101 = "N = ", right bold overwritefmt
putexcel F101 = `e(N)', hcenter bold overwritefmt
putexcel A102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled birth_readiness_all if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E123 = "N = ", right bold overwritefmt
putexcel F123 = `e(N)', hcenter bold overwritefmt
putexcel A124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_nd_info_yn if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A140 = "Nutritional counseling", bold overwritefmt
putexcel E141 = "N = ", right bold overwritefmt
putexcel F141 = `e(N)', hcenter bold overwritefmt
putexcel A142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_tt_inject if preg_comp==1 & select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A158 = "Tetanus injection", bold overwritefmt
putexcel E159 = "N = ", right bold overwritefmt
putexcel F159 = `e(N)', hcenter bold overwritefmt
putexcel A160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_bp if preg_comp==1 & select 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A176 = "Blood pressure measurement at ANC", bold overwritefmt
putexcel E177 = "N = ", right bold overwritefmt
putexcel F177 = `e(N)', hcenter bold overwritefmt
putexcel A178 = matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_weight if preg_comp==1 & select 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A194 = "Weight measurement at ANC", bold overwritefmt
putexcel E195 = "N = ", right bold overwritefmt
putexcel F195 = `e(N)', hcenter bold overwritefmt
putexcel A196 = matrix(results), names nformat(number_d2)

*** Urine test *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_urine if preg_comp==1 & select 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel A212 = "Urine test at ANC", bold overwritefmt
putexcel E213 = "N = ", right bold overwritefmt
putexcel F213 = `e(N)', hcenter bold overwritefmt
putexcel A214 = matrix(results), names nformat(number_d2)

*==================================================================*
* Among women who were pregnancy at enrollment and had sign of PE* 
*==================================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.anc_num_cat if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H1 = "Adjusted logistic regression results among women who were pregnancy at enrollment and had PE signs", bold overwritefmt
putexcel H3 = "ANC Frequency", bold overwritefmt
putexcel L4 = "N = ", right bold overwritefmt
putexcel M4 = `e(N)', hcenter bold overwritefmt
putexcel H5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.provider_code if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H23 = "ANC Provider Type", bold overwritefmt
putexcel L24 = "N = ", right bold overwritefmt
putexcel M24 = `e(N)', hcenter bold overwritefmt
putexcel H25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.maternal_assess_cat if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel L45 = "N = ", right bold overwritefmt
putexcel M45 = `e(N)', hcenter bold overwritefmt
putexcel H46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled maternal_assess_all if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel L66 = "N = ", right bold overwritefmt
putexcel M66 = `e(N)', hcenter bold overwritefmt
putexcel H67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled danger_sign_coun if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H83 = "PE/E danger sign counseling", bold overwritefmt
putexcel L84 = "N = ", right bold overwritefmt
putexcel M84 = `e(N)', hcenter bold overwritefmt
putexcel H85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.birth_readiness_cat if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel L101 = "N = ", right bold overwritefmt
putexcel M101 = `e(N)', hcenter bold overwritefmt
putexcel H102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled birth_readiness_all if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel L123 = "N = ", right bold overwritefmt
putexcel M123 = `e(N)', hcenter bold overwritefmt
putexcel H124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_nd_info_yn if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H140 = "Nutritional counseling", bold overwritefmt
putexcel L141 = "N = ", right bold overwritefmt
putexcel M141 = `e(N)', hcenter bold overwritefmt
putexcel H142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_tt_inject if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H158 = "Tetanus injection", bold overwritefmt
putexcel L159 = "N = ", right bold overwritefmt
putexcel M159 = `e(N)', hcenter bold overwritefmt
putexcel H160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_bp if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H176 = "Blood pressure measurement at ANC", bold overwritefmt
putexcel L177 = "N = ", right bold overwritefmt
putexcel M177 = `e(N)', hcenter bold overwritefmt
putexcel H178 = matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_weight if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H194 = "Weight measurement at ANC", bold overwritefmt
putexcel L195 = "N = ", right bold overwritefmt
putexcel M195 = `e(N)', hcenter bold overwritefmt
putexcel H196 = matrix(results), names nformat(number_d2)

*** Urine test *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_urine if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H212 = "Urine test at ANC", bold overwritefmt
putexcel L213 = "N = ", right bold overwritefmt
putexcel M213 = `e(N)', hcenter bold overwritefmt
putexcel H214 = matrix(results), names nformat(number_d2)


*==================================================================*
* Among women with PE, exclusing 5-9 weeks postpartum women * 
*==================================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.anc_num_cat if preg_comp==1 & select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O1 = "Adjusted logistic regression results among with PE, exclusing 5-9 weeks postpartum women", bold overwritefmt
putexcel O3 = "ANC Frequency", bold overwritefmt
putexcel S4 = "N = ", right bold overwritefmt
putexcel T4 = `e(N)', hcenter bold overwritefmt
putexcel O5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.provider_code if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O23 = "ANC Provider Type", bold overwritefmt
putexcel S24 = "N = ", right bold overwritefmt
putexcel T24 = `e(N)', hcenter bold overwritefmt
putexcel O25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.maternal_assess_cat if preg_comp==1 & select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel S45 = "N = ", right bold overwritefmt
putexcel T45 = `e(N)', hcenter bold overwritefmt
putexcel O46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled maternal_assess_all if preg_comp==1 & select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel S66 = "N = ", right bold overwritefmt
putexcel T66 = `e(N)', hcenter bold overwritefmt
putexcel O67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled danger_sign_coun if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O83 = "PE/E danger sign counseling", bold overwritefmt
putexcel S84 = "N = ", right bold overwritefmt
putexcel T84 = `e(N)', hcenter bold overwritefmt
putexcel O85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled i.birth_readiness_cat if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel S101 = "N = ", right bold overwritefmt
putexcel T101 = `e(N)', hcenter bold overwritefmt
putexcel O102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled birth_readiness_all if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel S123 = "N = ", right bold overwritefmt
putexcel T123 = `e(N)', hcenter bold overwritefmt
putexcel O124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_nd_info_yn if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O140 = "Nutritional counseling", bold overwritefmt
putexcel S141 = "N = ", right bold overwritefmt
putexcel T141 = `e(N)', hcenter bold overwritefmt
putexcel O142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_tt_inject if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O158 = "Tetanus injection", bold overwritefmt
putexcel S159 = "N = ", right bold overwritefmt
putexcel T159 = `e(N)', hcenter bold overwritefmt
putexcel O160 = matrix(results), names nformat(number_d2)


*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_bp if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O176 = "Blood pressure measurement at ANC", bold overwritefmt
putexcel S177 = "N = ", right bold overwritefmt
putexcel T177 = `e(N)', hcenter bold overwritefmt
putexcel O178 = matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_weight if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O194 = "Weight measurement at ANC", bold overwritefmt
putexcel S195 = "N = ", right bold overwritefmt
putexcel T195 = `e(N)', hcenter bold overwritefmt
putexcel O196 = matrix(results), names nformat(number_d2)

*** Urine test *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity4 i.facility_skilled SWanc_urine if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O212 = "Urine test at ANC", bold overwritefmt
putexcel S213 = "N = ", right bold overwritefmt
putexcel T213 = `e(N)', hcenter bold overwritefmt
putexcel O214 = matrix(results), names nformat(number_d2)


log close
