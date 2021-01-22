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


/*

Outcome: 
- Convulsions/fits during delivery 


Covariates of interest: 
- ANC frequency 
	* 0, 1-3, 4+ 
- ANC components: 
	* Nutrition counseling
	* Key services received (Had BP, urine and blood sampled, tested for syphilis/HIV, + took iron)
	* Birth readiness discussion 
	
Things to adjust for: 
- Background characteristics
	* Age
	* Residence
	* Parity
	* Education 
	* Wealth? 
	* Marital status? 
- Pregnancy intention
- Facility delivery and/or skilled birth attendant
- Complications experienced during pregnancy 


2,853 interviews completed and consented to follow up 
2,578 6-week interviews expected 

*/


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
use "/Users/Ellie/Desktop/THESIS/Data/Cohort1_6W_Merged_10Jan2021.dta", clear

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
				delivprob_malposition delivprob_prolonglab delivprob_convuls {
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

* Recode missing variable to 0 for women with no ANC 

foreach var in SWanc_delivery_place SWanc_delivery_skilled SWanc_emergency_transport ///
				SWanc_danger_place SWanc_danger_migraine SWanc_danger_hbp SWanc_danger_edema ///
				SWanc_danger_convuls SWanc_danger_bleeding SWanc_nd_info_nr ///
				SWanc_bp SWanc_weight SWanc_urine SWanc_blood SWanc_stool ///
				SWanc_syph_test  SWanc_hiv_test {
				
				recode `var' (. = 0) if anyanc==0
				
				}
				
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

* By outcome
foreach var in age_cat urban education wealthquintile parity4 unintended_preg facility_skilled preg_comp {

	svy: tab `var' SWdelivprob_convuls, col 
	
}

*=========================================================
* Logistic regression *
*=========================================================

*** Crude effect among all women ***

* Age
svy: logistic SWdelivprob_convuls i.age_cat 
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 20-34 35+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) replace
putexcel A1 = "Crude logistic regression results", bold overwritefmt
putexcel A2 = "Age", bold overwritefmt
putexcel E3 = "N = ", right bold overwritefmt
putexcel F3 = `e(N)', hcenter bold overwritefmt
putexcel A4 = matrix(results), names nformat(number_d2)

* Education
svy: logistic SWdelivprob_convuls i.education 
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = Primary Secondary+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A9 = "Education", bold overwritefmt
putexcel E10 = "N = ", right bold overwritefmt
putexcel F10 = `e(N)', hcenter bold overwritefmt
putexcel A11 = matrix(results), names nformat(number_d2)

* Wealth
svy: logistic SWdelivprob_convuls i.wealthquintile 
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = Lower Middle Higher Highest _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A16 = "Wealth", bold overwritefmt
putexcel E17 = "N = ", right bold overwritefmt
putexcel F17 = `e(N)', hcenter bold overwritefmt
putexcel A18 = matrix(results), names nformat(number_d2)

* Residence 
svy: logistic SWdelivprob_convuls urban 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Urban _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A25 = "Residence", bold overwritefmt
putexcel E26 = "N = ", right bold overwritefmt
putexcel F26 = `e(N)', hcenter bold overwritefmt
putexcel A27 = matrix(results), names nformat(number_d2)

* Parity 
svy: logistic SWdelivprob_convuls i.parity4 
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1 2-3 4+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A31 = "Parity", bold overwritefmt
putexcel E32 = "N = ", right bold overwritefmt
putexcel F32 = `e(N)', hcenter bold overwritefmt
putexcel A33 = matrix(results), names nformat(number_d2)

* Unintended pregnancy  
svy: logistic SWdelivprob_convuls unintended_preg
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A39 = "Unintended pregnancy", bold overwritefmt
putexcel E40 = "N = ", right bold overwritefmt
putexcel F40 = `e(N)', hcenter bold overwritefmt
putexcel A41 = matrix(results), names nformat(number_d2)

* Signs of preeclampsia 
svy: logistic SWdelivprob_convuls preg_comp
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A45 = "Preeclampsia", bold overwritefmt
putexcel E46 = "N = ", right bold overwritefmt
putexcel F46 = `e(N)', hcenter bold overwritefmt
putexcel A47 = matrix(results), names nformat(number_d2)

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A51 = "ANC Frequency", bold overwritefmt
putexcel E52 = "N = ", right bold overwritefmt
putexcel F52 = `e(N)', hcenter bold overwritefmt
putexcel A53 = matrix(results), names nformat(number_d2)

* ANC provider type
svy: logistic SWdelivprob_convuls i.provider_code
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = HEW PHCP both _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A59 = "ANC Provider Type", bold overwritefmt
putexcel E60 = "N = ", right bold overwritefmt
putexcel F60 = `e(N)', hcenter bold overwritefmt
putexcel A61 = matrix(results), names nformat(number_d2)

* Maternal assessment at ANC 
svy: logistic SWdelivprob_convuls maternal_assess_all
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A67 = "Received all 5 assessment at ANC", bold overwritefmt
putexcel E68 = "N = ", right bold overwritefmt
putexcel F68 = `e(N)', hcenter bold overwritefmt
putexcel A69 = matrix(results), names nformat(number_d2)

svy: logistic SWdelivprob_convuls i.maternal_assess_cat
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A73 = "Maternal assessment score at ANC", bold overwritefmt
putexcel E74 = "N = ", right bold overwritefmt
putexcel F74 = `e(N)', hcenter bold overwritefmt
putexcel A75 = matrix(results), names nformat(number_d2)

* BP measurement alone
svy: logistic SWdelivprob_convuls SWanc_bp
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A81 = "BP measurement at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82 = `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)

* Weight measurement alone
svy: logistic SWdelivprob_convuls SWanc_weight
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A87 = "Weight measurement at ANC", bold overwritefmt
putexcel E88 = "N = ", right bold overwritefmt
putexcel F88 = `e(N)', hcenter bold overwritefmt
putexcel A89 = matrix(results), names nformat(number_d2)

* Urine test alone
svy: logistic SWdelivprob_convuls SWanc_urine
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A93 = "Urine test at ANC", bold overwritefmt
putexcel E94 = "N = ", right bold overwritefmt
putexcel F94 = `e(N)', hcenter bold overwritefmt
putexcel A95 = matrix(results), names nformat(number_d2)

* Birth readiness discussion 
* Categorical 
svy: logistic SWdelivprob_convuls i.birth_readiness_cat
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4-6 7-9 _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A99 = "Birth readiness discussion (categorial)", bold overwritefmt
putexcel E100 = "N = ", right bold overwritefmt
putexcel F100 = `e(N)', hcenter bold overwritefmt
putexcel A101 = matrix(results), names nformat(number_d2)

* Binary 
svy: logistic SWdelivprob_convuls birth_readiness_all
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A107 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E108 = "N = ", right bold overwritefmt
putexcel F108 = `e(N)', hcenter bold overwritefmt
putexcel A109 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls danger_sign_coun 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A113 = "PE-related danger sign counseling", bold overwritefmt
putexcel E114 = "N = ", right bold overwritefmt
putexcel F114 = `e(N)', hcenter bold overwritefmt
putexcel A115 = matrix(results), names nformat(number_d2)

* Nutrition counseling
svy: logistic SWdelivprob_convuls SWanc_nd_info_yn  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A119 = "Nutritional counseling", bold overwritefmt
putexcel E120 = "N = ", right bold overwritefmt
putexcel F120 = `e(N)', hcenter bold overwritefmt
putexcel A121 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls SWanc_tt_inject  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A125 = "Tetanus injection at ANC", bold overwritefmt
putexcel E126 = "N = ", right bold overwritefmt
putexcel F126 = `e(N)', hcenter bold overwritefmt
putexcel A127 = matrix(results), names nformat(number_d2)

* Facility deliver and SBA 
svy: logistic SWdelivprob_convuls i.facility_skilled  
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = Home_SBA Facility_SBA _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A131 = "Facility delivery and SBA", bold overwritefmt
putexcel E132 = "N = ", right bold overwritefmt
putexcel F132 = `e(N)', hcenter bold overwritefmt
putexcel A133 = matrix(results), names nformat(number_d2)


*** Crude effect among women with signs of preeclampsia ***

* Age
svy: logistic SWdelivprob_convuls i.age_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 20-34 35+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A1 = "Crude logistic regression results among women with signs of preeclampsia", bold overwritefmt
putexcel A2 = "Age", bold overwritefmt
putexcel E3 = "N = ", right bold overwritefmt
putexcel F3 = `e(N)', hcenter bold overwritefmt
putexcel A4 = matrix(results), names nformat(number_d2)

* Education
svy: logistic SWdelivprob_convuls i.education if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = Primary Secondary+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A9 = "Education", bold overwritefmt
putexcel E10 = "N = ", right bold overwritefmt
putexcel F10 = `e(N)', hcenter bold overwritefmt
putexcel A11 = matrix(results), names nformat(number_d2)

* Wealth
svy: logistic SWdelivprob_convuls i.wealthquintile if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = Lower Middle Higher Highest _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A16 = "Wealth", bold overwritefmt
putexcel E17 = "N = ", right bold overwritefmt
putexcel F17 = `e(N)', hcenter bold overwritefmt
putexcel A18 = matrix(results), names nformat(number_d2)

* Residence 
svy: logistic SWdelivprob_convuls urban if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Urban _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A25 = "Residence", bold overwritefmt
putexcel E26 = "N = ", right bold overwritefmt
putexcel F26 = `e(N)', hcenter bold overwritefmt
putexcel A27 = matrix(results), names nformat(number_d2)

* Parity 
svy: logistic SWdelivprob_convuls i.parity4 if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1 2-3 4+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A31 = "Parity", bold overwritefmt
putexcel E32 = "N = ", right bold overwritefmt
putexcel F32 = `e(N)', hcenter bold overwritefmt
putexcel A33 = matrix(results), names nformat(number_d2)

* Unintended pregnancy  
svy: logistic SWdelivprob_convuls unintended_preg if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A39 = "Unintended pregnancy", bold overwritefmt
putexcel E40 = "N = ", right bold overwritefmt
putexcel F40 = `e(N)', hcenter bold overwritefmt
putexcel A41 = matrix(results), names nformat(number_d2)

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A45 = "ANC Frequency", bold overwritefmt
putexcel E46 = "N = ", right bold overwritefmt
putexcel F46 = `e(N)', hcenter bold overwritefmt
putexcel A47 = matrix(results), names nformat(number_d2)

* ANC provider type
svy: logistic SWdelivprob_convuls i.provider_code if preg_comp==1 
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = PHCP both _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A53 = "ANC Provider Type", bold overwritefmt
putexcel E54 = "N = ", right bold overwritefmt
putexcel F54 = `e(N)', hcenter bold overwritefmt
putexcel A55 = matrix(results), names nformat(number_d2)

* Maternal assessment at ANC 
svy: logistic SWdelivprob_convuls maternal_assess_all if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A61 = "Received all 5 assessment at ANC", bold overwritefmt
putexcel E62 = "N = ", right bold overwritefmt
putexcel F62 = `e(N)', hcenter bold overwritefmt
putexcel A63 = matrix(results), names nformat(number_d2)

svy: logistic SWdelivprob_convuls i.maternal_assess_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A67 = "Maternal assessment score at ANC", bold overwritefmt
putexcel E68 = "N = ", right bold overwritefmt
putexcel F68 = `e(N)', hcenter bold overwritefmt
putexcel A69 = matrix(results), names nformat(number_d2)

* BP measurement alone
svy: logistic SWdelivprob_convuls SWanc_bp if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A75 = "BP measurement at ANC", bold overwritefmt
putexcel E76 = "N = ", right bold overwritefmt
putexcel F76 = `e(N)', hcenter bold overwritefmt
putexcel A77 = matrix(results), names nformat(number_d2)

* Weight measurement alone
svy: logistic SWdelivprob_convuls SWanc_weight if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A81 = "Weight measurement at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82 = `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)

* Urine test alone
svy: logistic SWdelivprob_convuls SWanc_urine if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A87 = "Urine test at ANC", bold overwritefmt
putexcel E88 = "N = ", right bold overwritefmt
putexcel F88 = `e(N)', hcenter bold overwritefmt
putexcel A89 = matrix(results), names nformat(number_d2)

* Birth readiness discussion 
* Categorical 
svy: logistic SWdelivprob_convuls i.birth_readiness_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A93 = "Birth readiness discussion (categorial)", bold overwritefmt
putexcel E94 = "N = ", right bold overwritefmt
putexcel F94 = `e(N)', hcenter bold overwritefmt
putexcel A95 = matrix(results), names nformat(number_d2)

* Binary 
svy: logistic SWdelivprob_convuls birth_readiness_all if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A102 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E103 = "N = ", right bold overwritefmt
putexcel F103 = `e(N)', hcenter bold overwritefmt
putexcel A104 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls danger_sign_coun if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A108 = "PE-related danger sign counseling", bold overwritefmt
putexcel E109 = "N = ", right bold overwritefmt
putexcel F109 = `e(N)', hcenter bold overwritefmt
putexcel A110 = matrix(results), names nformat(number_d2)

* Nutrition counseling
svy: logistic SWdelivprob_convuls SWanc_nd_info_yn if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A114 = "Nutritional counseling", bold overwritefmt
putexcel E115 = "N = ", right bold overwritefmt
putexcel F115 = `e(N)', hcenter bold overwritefmt
putexcel A116 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls SWanc_tt_inject if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A120 = "Tetanus injection at ANC", bold overwritefmt
putexcel E121 = "N = ", right bold overwritefmt
putexcel F121 = `e(N)', hcenter bold overwritefmt
putexcel A122 = matrix(results), names nformat(number_d2)

* Facility deliver and SBA 
svy: logistic SWdelivprob_convuls i.facility_skilled if preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = Home_SBA Facility_SBA _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A126 = "Facility delivery and SBA", bold overwritefmt
putexcel E127 = "N = ", right bold overwritefmt
putexcel F127 = `e(N)', hcenter bold overwritefmt
putexcel A128 = matrix(results), names nformat(number_d2)

*=========================================================
* Adjusted results *
*=========================================================

*==================================*
* Among all women * 
*==================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.parity4 preg_comp i.facility_skilled  i.anc_num_cat 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A1 = "Adjusted logistic regression results", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)


*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  i.provider_code 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A28 = "ANC Provider Type", bold overwritefmt
putexcel E29 = "N = ", right bold overwritefmt
putexcel F29 = `e(N)', hcenter bold overwritefmt
putexcel A30 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  i.maternal_assess_cat 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A58 = "Maternal assessment (categorial)", bold overwritefmt
putexcel E59 = "N = ", right bold overwritefmt
putexcel F59 = `e(N)', hcenter bold overwritefmt
putexcel A60 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  danger_sign_coun
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A87 = "PE/E danger sign counseling", bold overwritefmt
putexcel E88 = "N = ", right bold overwritefmt
putexcel F88 = `e(N)', hcenter bold overwritefmt
putexcel A89 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  i.birth_readiness_cat
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A113 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel E114 = "N = ", right bold overwritefmt
putexcel F114 = `e(N)', hcenter bold overwritefmt
putexcel A115 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  SWanc_nd_info_yn
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A142 = "Nutritional counseling", bold overwritefmt
putexcel E143 = "N = ", right bold overwritefmt
putexcel F143 = `e(N)', hcenter bold overwritefmt
putexcel A144 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  SWanc_tt_inject
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A168 = "Tetanus injection", bold overwritefmt
putexcel E169 = "N = ", right bold overwritefmt
putexcel F169 = `e(N)', hcenter bold overwritefmt
putexcel A170 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled maternal_assess_all
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A196 = "Received all 5 maternal assessments", bold overwritefmt
putexcel E197 = "N = ", right bold overwritefmt
putexcel F197 = `e(N)', hcenter bold overwritefmt
putexcel A198 = matrix(results), names nformat(number_d2)

*=========================================*
* Among women with signs of preeclampsia * 
*=========================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.parity4 i.facility_skilled  i.anc_num_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A1 = "Adjusted logistic regression results among women with signs preeclampsia", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled  i.provider_code if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A28 = "ANC Provider Type", bold overwritefmt
putexcel E29 = "N = ", right bold overwritefmt
putexcel F29 = `e(N)', hcenter bold overwritefmt
putexcel A30 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled  i.maternal_assess_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A58 = "Maternal assessment (categorial)", bold overwritefmt
putexcel E59 = "N = ", right bold overwritefmt
putexcel F59 = `e(N)', hcenter bold overwritefmt
putexcel A60 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled  danger_sign_coun if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A87 = "PE/E danger sign counseling", bold overwritefmt
putexcel E88 = "N = ", right bold overwritefmt
putexcel F88 = `e(N)', hcenter bold overwritefmt
putexcel A89 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled  i.birth_readiness_cat if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A113 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel E114 = "N = ", right bold overwritefmt
putexcel F114 = `e(N)', hcenter bold overwritefmt
putexcel A115 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled  SWanc_nd_info_yn if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A142 = "Nutritional counseling", bold overwritefmt
putexcel E143 = "N = ", right bold overwritefmt
putexcel F143 = `e(N)', hcenter bold overwritefmt
putexcel A144 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled  SWanc_tt_inject if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A168 = "Tetanus injection", bold overwritefmt
putexcel E169 = "N = ", right bold overwritefmt
putexcel F169 = `e(N)', hcenter bold overwritefmt
putexcel A170 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 i.facility_skilled maternal_assess_all if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
mat list results
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A196 = "Received all 5 maternal assessments", bold overwritefmt
putexcel E197 = "N = ", right bold overwritefmt
putexcel F197 = `e(N)', hcenter bold overwritefmt
putexcel A198 = matrix(results), names nformat(number_d2)

;


svy: logistic SWdelivprob_convuls i.age_cat urban i.facility_skilled  i.anc_num_cat##preg_comp 

svy: logistic SWdelivprob_convuls i.age_cat urban i.wealthquintile i.education i.parity4 preg_comp i.facility_skilled  i.provider_code i.birth_readiness_cat maternal_assess_all
;
svy: logistic SWdelivprob_convuls birth_readiness i.facility_skilled i.age_cat urban preg_comp

svy: logistic SWdelivprob_convuls i.anc_num_cat i.facility_skilled i.age_cat urban if preg_comp==1  //unintended_preg married i.education

/*

gen pe_no_e=0 if preg_comp==1 & SWdelivprob_convuls==1 // pe & e
replace pe_no_e=1 if preg_comp==1 & SWdelivprob_convuls==0 // pe but no e

*/


log close
