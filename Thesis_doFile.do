*=========================================================
* THESIS .DO FILE *
*=========================================================

clear
clear matrix
clear mata
capture log close
set maxvar 15000
set more off
numlabel, add


*=========================================================
* SET UP MAROS *
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
use "/Users/Ellie/Desktop/THESIS/Data/Cohort1_6W_Merged_25Mar2021.dta", clear

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
* DEMOGRAPHIC INFORMATION *
*=========================================================

* Replace 6w data with baseline for 5-9 weeks postpartum women
foreach var in delivery_place who_assisted_delivery anc_hew_place anc_phcp_place ///
				anc_delivery_place anc_delivery_skilled anc_emergency_transport ///
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
gen age_cat=0 if FQ_age<20
replace age_cat=1 if FQ_age>=20 & FQ_age<35
replace age_cat=2 if FQ_age>=35
label define agel 0 "15-19" 1 "20-34" 2 "35-49"
label value age_cat agel

* Generate categorical variable for parity 
egen parity4=cut(total_births), at(0, 1, 2, 4, 30) icodes

gen parity_cat=0 if parity4==0
replace parity_cat=1 if parity4==1
replace parity_cat=2 if parity4==2
replace parity_cat=3 if parity4==3
lab def parity_catl 0 "0 children" 1 "1 children" 2 "2-3 children" 3 "4+ children"
lab val parity_cat parity_catl
lab var parity_cat "Parity" 

* Group secondary, technical & vocational, and higher education 
gen education=school
replace education=2 if school>=3
lab def edul 0 "No education" 1 "Primary" 2 "Secondary+" 
lab val education edul
lab var education "Education level" 

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

* Facility delivery and skilled birth attendant combined variable
gen facility_skilled=0 if facility_deliv!=. & sba!=.
replace facility_skilled=1 if (facility_deliv==0 & sba==0)
replace facility_skilled=2 if (facility_deliv==0 & sba==1)
replace facility_skilled=3 if facility_deliv==1
label define facility_skilledl 1 "Home delivery with no SBA" 2 "Home delivery with SBA" 3 "Facility delivery with SBA"
label val facility_skilled facility_skilledl
tab facility_skilled, m

* PE/E-related Pregnancy complications binary variable 
gen preg_comp=0
replace preg_comp=1 if (SWpregprob_hbp==1 | SWpregprob_edema==1 |  SWpregprob_migraine==1) 
label var preg_comp "Had PE/E-related complications during pregnancy" 
label val preg_comp yesno


*=========================================================
* ANC COVERAGE AND QUALITY INDICATORS *
*=========================================================

/*
		NOTE: 
			1. "0" include women with no ANC 
			2. All the variables that did not come directly from the questionnaire do not have any missing
				-88, -99, and . are all treated as 0 in generated variables 
*/

* Recode missing variable to 0 for women with no ANC 

foreach var in SWanc_delivery_place SWanc_delivery_skilled SWanc_emergency_transport ///
				SWanc_danger_place SWanc_danger_migraine SWanc_danger_hbp SWanc_danger_edema ///
				SWanc_danger_convuls SWanc_danger_bleeding SWanc_nd_info_yn ///
				SWanc_bp SWanc_weight SWanc_urine SWanc_blood SWanc_stool ///
				SWanc_syph_test  SWanc_hiv_test SWanc_tt_inject SWanc_nd_info_iron {
				
				recode `var' (. -88 -99 = 0)
				
				}
				
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

* ANC by provider type
gen provider_code=0 if anyanc==0
replace provider_code=1 if SWanc_hew_yn==1 & SWanc_phcp_yn!=1 & anyanc==1 
replace provider_code=2  if SWanc_phcp_yn==1 & SWanc_hew_yn!=1 & anyanc==1
replace provider_code=3 if SWanc_phcp_yn==1 & SWanc_hew_yn==1 & anyanc==1
label define providerl 0 "No ANC" 1 "HEW only" 2 "PHCP only" 3 "Both" 
label val provider_code providerl

* Place of ANC
gen anc_place=0 if anyanc==0
replace anc_place=1 if anyanc==1 & ///
						provider_code==1 | ///
						(SWanc_phcp_place=="nowhere" & provider_code==3) | ///
						((SWanc_phcp_place=="her_home" | SWanc_phcp_place=="other_home") & ///
						(provider_code==2 | provider_code==3))
						
/* 	Women who reported receiving ANC from both HEW and PHCP but reported "nowhere" or "her home" for the place of ANC
	were coded to have received care from HEW.
	Since very few women received care from HEW at home only, HEW at home and HEW at HF are grouped together.

*/
			
replace anc_place=2 if anyanc==1 & ///
						(SWanc_phcp_place=="gov_hc" | SWanc_phcp_place=="gov_hc gov_hp" | ///
						SWanc_phcp_place=="gov_hc ngo_hf" | SWanc_phcp_place=="gov_hc other " | ///
						SWanc_phcp_place=="gov_hc other_priv" | SWanc_phcp_place=="gov_hc other_pub" | ///
						SWanc_phcp_place=="gov_hp" | SWanc_phcp_place=="ngo_hf" | ///
						SWanc_phcp_place=="other_priv" | SWanc_phcp_place=="other_pub") & ///
						(provider_code==2 | provider_code==3)
		
replace anc_place=3 if anyanc==1 & ///	
						(SWanc_phcp_place=="gov_hc gov_hp priv_hosp" | SWanc_phcp_place=="gov_hc priv_hosp" | ///
						SWanc_phcp_place=="gov_hc priv_hosp other" | SWanc_phcp_place=="gov_hosp" | ///
						SWanc_phcp_place=="gov_hosp gov_hc" | SWanc_phcp_place=="gov_hosp gov_hc gov_hp" | ///
						SWanc_phcp_place=="gov_hosp gov_hc other" | SWanc_phcp_place=="gov_hosp gov_hc priv_hosp" | ///
						SWanc_phcp_place=="gov_hosp gov_hc priv_hosp ngo_hf" | SWanc_phcp_place=="gov_hosp ngo_hf" | ///
						SWanc_phcp_place=="gov_hosp other_priv" | SWanc_phcp_place=="gov_hosp other_pub" | ///
						SWanc_phcp_place=="gov_hosp priv_hosp" | SWanc_phcp_place=="gov_hp priv_hosp" | ///
						SWanc_phcp_place=="priv_hosp" | SWanc_phcp_place=="priv_hosp ngo_hf" | SWanc_phcp_place=="priv_hosp other_priv") & ///
						(provider_code==2 | provider_code==3)
													
label define anc_placel 0 "No ANC" 1 "HEW" 2 "PHCP at HF" 3 "PHCP at hospital"
label val anc_place anc_placel
tab anc_place, m

* Generate GA at first visit variable
recode SWanc_phcp_timing(0 = 1)
recode SWanc_hew_timing(0 = 1)
gen ga_first_anc=SWanc_hew_timing
replace ga_first_anc=SWanc_phcp_timing if SWanc_phcp_timing<anc_hew_timing & SWanc_phcp_timing!=.
replace ga_first_anc=. if SWanc_hew_timing==. & SWanc_phcp_timing==.
replace ga_first_anc=. if anyanc==0
label var ga_first_anc "GA (in months) at first ANC"

* Create categories 
gen ga_first_anc_cat=0 if anyanc==0
replace ga_first_anc_cat=1 if ga_first_anc<4 & ga_first_anc>=1 & anyanc==1
replace ga_first_anc_cat=2 if ga_first_anc>=4 & ga_first_anc<=6 & anyanc==1
replace ga_first_anc_cat=3 if ga_first_anc>6 & ga_first_anc<=9 & anyanc==1
replace ga_first_anc_cat=4 if ga_first_anc==. & anyanc==1
label define ga_anc_l 0 "No ANC" 1 "First trimester" 2 "Second trimester" 3 "Third trimester" 4 "Missing ANC timing"
label val ga_first_anc_cat ga_anc_l


* Received care in the first trimester 
gen first_tri=0 if anyanc==0
replace first_tri=1 if ga_first_anc_cat>1 & anyanc==1
replace first_tri=2 if ga_first_anc_cat==1 & anyanc==1
label define gal 0 "No ANC" 1 "Received ANC after first trimester" 2 "Received ANC in the first trimester"
label val first_tri gal
tab first_tri

* PE/E danger sign counseling
gen danger_sign_coun=0 if anyanc==0
replace danger_sign_coun=1 if (SWanc_danger_migraine!=1 & anc_danger_hbp!=1 & anc_danger_edema!=1 & SWanc_danger_convuls!=1) & anyanc==1
replace danger_sign_coun=2 if (SWanc_danger_migraine==1 | anc_danger_hbp==1 | anc_danger_edema==1 | SWanc_danger_convuls==1) & anyanc==1
label var danger_sign_coun "Have at least one PE/E danger sign counseling at ANC"
label define dangerl 0 "No ANC" 1 "Did not receive any danger sign counseling" 2 "Received danger sign counseling at ANC"
label val danger_sign_coun dangerl
tab danger_sign_coun

gen danger_sign_coun1=0 
replace danger_sign_coun1=1 if (SWanc_danger_migraine==1 | anc_danger_hbp==1 | anc_danger_edema==1 | SWanc_danger_convuls==1) & anyanc==1

* Blood pressure
gen bp=0 if anyanc==0
replace bp=1 if anyanc==1 & SWanc_bp==0
replace bp=2 if anyanc==1 & SWanc_bp==1
label define bpl 0 "No ANC" 1 "No BP measurement at ANC" 2 "BP measured at ANC"
label val bp bpl

* Weight
gen weight=0 if anyanc==0
replace weight=1 if anyanc==1 & SWanc_weight==0
replace weight=2 if anyanc==1 & SWanc_weight==1
label define weightl 0 "No ANC" 1 "No weight measurement at ANC" 2 "Weight measured at ANC"
label val weight weightl

* Urine sample
gen urine=0 if anyanc==0
replace urine=1 if anyanc==1 & SWanc_urine==0
replace urine=2 if anyanc==1 & SWanc_urine==1
label define urinel 0 "No ANC" 1 "No urine sample taken at ANC" 2 "Urine sample taken at ANC"
label val urine urinel

* Blood sample 
gen blood=0 if anyanc==0
replace blood=1 if anyanc==1 & SWanc_blood==0
replace blood=2 if anyanc==1 & SWanc_blood==1
label define bloodl 0 "No ANC" 1 "No blood sample taken at ANC" 2 "Blood sample taken at ANC"
label val blood bloodl
tab blood

* Tetanus
gen tt=0 if anyanc==0
replace tt=1 if anyanc==1 & SWanc_tt_inject==0
replace tt=2 if anyanc==1 & SWanc_tt_inject==1
label define ttl 0 "No ANC" 1 "No tt injection ANC" 2 "Had tt injection at ANC"
label val tt ttl
tab tt

* Iron counseling 
gen iron=0 if anyanc==0
replace iron=1 if anyanc==1 & SWanc_nd_info_iron==0
replace iron=2 if anyanc==1 & SWanc_nd_info_iron==1
label define ironl 0 "No ANC" 1 "No counseling on iron at ANC" 2 "Had counseling on iron at ANC"
label val iron ironl
tab iron

* Receiving all content and content score variables 
gen all_content=0 if anyanc==0
replace all_content=1 if anyanc==1
replace all_content=2 if (SWanc_tt_inject==1 & SWanc_bp==1 &  SWanc_blood==1 &  ///
							SWanc_urine==1 & danger_sign_coun1==1 & SWanc_nd_info_iron==1) & ///
							anyanc==1

*=========================================================
* OUTCOME VARIABLE *
*=========================================================
							
gen eclampsia=0
replace eclampsia=1 if SWdelivprob_convuls==1 | SWpostdelivprob_convuls==1 

*=========================================================
* DESCRIBE MISSINGNESS *
*=========================================================
			
preserve 
keep age urban school wealthquintile total_births parity_cat pregnancy_desired ///
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
* TABLE 1 *
*=========================================================

quietly svy: logistic eclampsia i.age_cat urban i.parity_cat i.facility_skilled preg_comp i.tt i.anc_num_cat i.anc_place i.iron
gen select = e(sample)
tab select

* By outcome
foreach var in age_cat urban education wealthquintile region married ///
				parity_cat preg_comp facility_skilled ///
				anc_num_cat anc_place first_tri bp blood urine danger_sign_coun tt iron {

	svy: tab `var' eclampsia if select, col
	
}

*=====================================================================
* MULTIPLE LOGISTIC REGRESSOIN  *
*=====================================================================

*==================================*
* Among all women * 
*==================================*

*** ANC Frequency *** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_num_cat if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) replace
putexcel A1 = "Adjusted logistic regression results among all women", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)

*** ANC provider type and location *** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_place if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A30 = "ANC Provider Type and Location", bold overwritefmt
putexcel E31 = "N = ", right bold overwritefmt
putexcel F31 = `e(N)', hcenter bold overwritefmt
putexcel A32 = matrix(results), names nformat(number_d2)

*** ANC in the first trimester *** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.first_tri if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A60 = "Received ANC in the first trimester", bold overwritefmt
putexcel E61 = "N = ", right bold overwritefmt
putexcel F61 = `e(N)', hcenter bold overwritefmt
putexcel A62 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.bp if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A90 = "BP measurement at ANC", bold overwritefmt
putexcel E91 = "N = ", right bold overwritefmt
putexcel F91 = `e(N)', hcenter bold overwritefmt
putexcel A92= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.urine if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A120 = "Urine test at ANC", bold overwritefmt
putexcel E121 = "N = ", right bold overwritefmt
putexcel F121 = `e(N)', hcenter bold overwritefmt
putexcel A122= matrix(results), names nformat(number_d2)

*** Blood sample *** 
svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.blood if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A150 = "Blood sample taken at ANC", bold overwritefmt
putexcel E151 = "N = ", right bold overwritefmt
putexcel F151 = `e(N)', hcenter bold overwritefmt
putexcel A152= matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.danger_sign_coun if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A180 = "PE/E danger sign counseling", bold overwritefmt
putexcel E181 = "N = ", right bold overwritefmt
putexcel F181 = `e(N)', hcenter bold overwritefmt
putexcel A182 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.tt if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A210 = "Tetanus injection", bold overwritefmt
putexcel E211 = "N = ", right bold overwritefmt
putexcel F211 = `e(N)', hcenter bold overwritefmt
putexcel A212 = matrix(results), names nformat(number_d2)

*** Counseling on iron ** 
svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.iron if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A240 = "Iron counseling", bold overwritefmt
putexcel E241 = "N = ", right bold overwritefmt
putexcel F241 = `e(N)', hcenter bold overwritefmt
putexcel A242 = matrix(results), names nformat(number_d2)

*=====================================================*
* Among rural women * 
*=====================================================*

*** ANC Frequency *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.anc_num_cat if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A1 = "Adjusted logistic regression results among all women", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)

*** ANC provider type and location *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.anc_place if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A28 = "ANC Provider Type and Location", bold overwritefmt
putexcel E29 = "N = ", right bold overwritefmt
putexcel F29 = `e(N)', hcenter bold overwritefmt
putexcel A30 = matrix(results), names nformat(number_d2)

*** ANC in the first trimester *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.first_tri if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A55 = "Received ANC in the first trimester", bold overwritefmt
putexcel E56 = "N = ", right bold overwritefmt
putexcel F56 = `e(N)', hcenter bold overwritefmt
putexcel A57 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.bp if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A80 = "BP measurement at ANC", bold overwritefmt
putexcel E81 = "N = ", right bold overwritefmt
putexcel F81 = `e(N)', hcenter bold overwritefmt
putexcel A82= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.urine if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A105 = "Urine test at ANC", bold overwritefmt
putexcel E106 = "N = ", right bold overwritefmt
putexcel F106 = `e(N)', hcenter bold overwritefmt
putexcel A107 = matrix(results), names nformat(number_d2)

*** Blood sample *** 
svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.blood if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A130 = "Blood sample taken at ANC", bold overwritefmt
putexcel E131 = "N = ", right bold overwritefmt
putexcel F131 = `e(N)', hcenter bold overwritefmt
putexcel A132= matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.danger_sign_coun if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A155 = "PE/E danger sign counseling", bold overwritefmt
putexcel E156 = "N = ", right bold overwritefmt
putexcel F156 = `e(N)', hcenter bold overwritefmt
putexcel A157 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.tt if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A180 = "Tetanus injection", bold overwritefmt
putexcel E181 = "N = ", right bold overwritefmt
putexcel F181 = `e(N)', hcenter bold overwritefmt
putexcel A182 = matrix(results), names nformat(number_d2)

*** Counseling on iron ** 
svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.iron if select & urban==0
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A205 = "Iron counseling", bold overwritefmt
putexcel E206 = "N = ", right bold overwritefmt
putexcel F206 = `e(N)', hcenter bold overwritefmt
putexcel A207 = matrix(results), names nformat(number_d2)


*=====================================================*
* Among urban women * 
*=====================================================*

*** ANC Frequency *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.anc_num_cat if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A1 = "Adjusted logistic regression results among all women", bold overwritefmt
putexcel A3 = "ANC Frequency", bold overwritefmt
putexcel E4 = "N = ", right bold overwritefmt
putexcel F4 = `e(N)', hcenter bold overwritefmt
putexcel A5 = matrix(results), names nformat(number_d2)

*** ANC provider type and location *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.anc_place if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A28 = "ANC Provider Type and Location", bold overwritefmt
putexcel E29 = "N = ", right bold overwritefmt
putexcel F29 = `e(N)', hcenter bold overwritefmt
putexcel A30 = matrix(results), names nformat(number_d2)

*** ANC in the first trimester *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.first_tri if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A55 = "Received ANC in the first trimester", bold overwritefmt
putexcel E56 = "N = ", right bold overwritefmt
putexcel F56 = `e(N)', hcenter bold overwritefmt
putexcel A57 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.bp if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A80 = "BP measurement at ANC", bold overwritefmt
putexcel E81 = "N = ", right bold overwritefmt
putexcel F81 = `e(N)', hcenter bold overwritefmt
putexcel A82= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.urine if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A105 = "Urine test at ANC", bold overwritefmt
putexcel E106 = "N = ", right bold overwritefmt
putexcel F106 = `e(N)', hcenter bold overwritefmt
putexcel A107 = matrix(results), names nformat(number_d2)

*** Blood sample *** 
svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.blood if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A130 = "Blood sample taken at ANC", bold overwritefmt
putexcel E131 = "N = ", right bold overwritefmt
putexcel F131 = `e(N)', hcenter bold overwritefmt
putexcel A132= matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.danger_sign_coun if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A155 = "PE/E danger sign counseling", bold overwritefmt
putexcel E156 = "N = ", right bold overwritefmt
putexcel F156 = `e(N)', hcenter bold overwritefmt
putexcel A157 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.tt if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A180 = "Tetanus injection", bold overwritefmt
putexcel E181 = "N = ", right bold overwritefmt
putexcel F181 = `e(N)', hcenter bold overwritefmt
putexcel A182 = matrix(results), names nformat(number_d2)

*** Counseling on iron ** 
svy: logistic eclampsia i.age_cat i.parity_cat i.facility_skilled preg_comp i.iron if select & urban==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-2) modify
putexcel A205 = "Iron counseling", bold overwritefmt
putexcel E206 = "N = ", right bold overwritefmt
putexcel F206 = `e(N)', hcenter bold overwritefmt
putexcel A207 = matrix(results), names nformat(number_d2)



log close
