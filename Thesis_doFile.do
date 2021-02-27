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
replace preg_comp=1 if (SWpregprob_hbp==1 | SWpregprob_edema==1 | SWpregprob_convuls==1 | SWpregprob_vision==1) 
label var preg_comp "Had PE/E-related complications during pregnancy" 
label val preg_comp yesno


*=========================================================
* ANC-related covariates *
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
						((SWanc_phcp_place=="nowhere" | SWanc_phcp_place=="her_home") & provider_code==3)
						
/* 	Women who reported receiving ANC from both HEW and PHCP but reported "nowhere" or "her home" for the place of ANC
	were coded to have received care from HEW.
	Since very few women received care from HEW at home only, HEW at home and HEW at HF are grouped together.

*/

/*
replace anc_place=2 if anyanc==1 & ///
						(SWanc_hew_place=="gov_hp" | SWanc_hew_place =="gov_hp other" | ///
						SWanc_hew_place =="gov_hp other_hf" | SWanc_hew_place =="home gov_hp" | SWanc_hew_place == "home gov_hp other_hf" | ///
						SWanc_hew_place =="other_hf") & ///
						provider_code==1 */
				
replace anc_place=2 if anyanc==1 & ///
						(SWanc_phcp_place=="gov_hc" | SWanc_phcp_place=="gov_hc gov_hp" | SWanc_phcp_place=="her_home gov_hc" | ///
						SWanc_phcp_place=="gov_hc ngo_hf" | SWanc_phcp_place=="gov_hc other " | SWanc_phcp_place=="gov_hc other_priv" | SWanc_phcp_place=="gov_hc other_pub" | ///
						SWanc_phcp_place=="gov_hp" | SWanc_phcp_place=="ngo_hf" | SWanc_phcp_place=="other_priv" | SWanc_phcp_place=="other_pub") & ///
						(provider_code==2 | provider_code==3)
		
replace anc_place=3 if anyanc==1 & ///	
						(SWanc_phcp_place=="gov_hc gov_hp priv_hosp" | SWanc_phcp_place=="gov_hc priv_hosp" | ///
						SWanc_phcp_place=="gov_hc priv_hosp other" | SWanc_phcp_place=="gov_hosp" | SWanc_phcp_place=="gov_hosp gov_hc" | ///
						SWanc_phcp_place=="gov_hosp gov_hc gov_hp" | SWanc_phcp_place=="gov_hosp gov_hc other" | SWanc_phcp_place=="gov_hosp gov_hc priv_hosp" | ///
						SWanc_phcp_place=="gov_hosp gov_hc priv_hosp ngo_hf" | SWanc_phcp_place=="gov_hosp gov_hp" | SWanc_phcp_place=="gov_hosp ngo_hf" | ///
						SWanc_phcp_place=="gov_hosp other_priv" | SWanc_phcp_place=="gov_hosp other_pub" | SWanc_phcp_place=="gov_hosp priv_hosp" | ///
						SWanc_phcp_place=="gov_hp priv_hosp" | SWanc_phcp_place=="priv_hosp" | SWanc_phcp_place=="priv_hosp ngo_hf" | SWanc_phcp_place=="priv_hosp other_priv") & ///
						(provider_code==2 | provider_code==3)
													
label define anc_placel 0 "No ANC" 1 "HEW (home or HF)" 2 "PHCP at HF" 3 "PHCP at hospital"
label val anc_place anc_placel
tab anc_place

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
tab danger_sign_coun, m

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
replace all_content=2 if (SWanc_tt_inject==1 & SWanc_bp==1 &  SWanc_blood==1 &  SWanc_urine==1 & danger_sign_coun==1 & SWanc_iron==1) & anyanc==1

gen content_score=0 if anyanc==0
replace  content_score=SWanc_tt_inject+SWanc_bp+SWanc_blood+SWanc_urine+danger_sign_coun+SWanc_iron if anyanc==1
tab content_score 
tab all_content

*=========================================================
* Describe missingness *
*=========================================================
			
preserve 
keep age urban school wealthquintile total_births parity_cat pregnancy_desired unintended_preg ///
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
/*
tabout SWdelivprob_convuls age_cat [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls education [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls urban [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls married [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls parity_cat [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
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
for var parity_cat:  tabulate X SWdelivprob_convuls, chi2 exact column
for var education:  tabulate X SWdelivprob_convuls, chi2 exact column
for var anc_num_cat:  tabulate X SWdelivprob_convuls, chi2 exact column


graph box FQ_age, name(age_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
	marker(1,mlab(FQ_age)) t1("Age by provider type") scale(1.2) ytitle("Age (years)") 
	graph export graphs/age_provider.png, replace

graph box total_birth, name(parity_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
	marker(1,mlab(total_birth)) t1("Parity by provider type") scale(1.2) ytitle("Number of children") 
	graph export graphs/parity_provider.png, replace



*=========================================================
* Table 1 *
*=========================================================

* Run the biggest model to keep sample size consistent 
quietly svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_tt_inject i.anc_num_cat 
gen select = e(sample)
*/

quietly svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.tt i.anc_num_cat i.anc_place i.iron
gen select = e(sample)
tab select

* By outcome
foreach var in age_cat urban education wealthquintile region married ///
				parity_cat unintended_preg preg_comp facility_skilled ///
				anc_num_cat anc_place first_tri bp weight blood urine danger_sign_coun tt iron {

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

* ANC provider type and location 
svy: logistic SWdelivprob_convuls i.anc_place if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = HEW PHCP_HF PHCP_hos _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A8 = "ANC Provider Type and Location", bold overwritefmt
putexcel E9 = "N = ", right bold overwritefmt
putexcel F9 = `e(N)', hcenter bold overwritefmt
putexcel A10 = matrix(results), names nformat(number_d2)

* ANC timing 
svy: logistic SWdelivprob_convuls i.ga_first_anc_cat if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = first_trimester second_trimester third_trimester missing _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A16 = "Timing at first ANC", bold overwritefmt
putexcel E17 = "N = ", right bold overwritefmt
putexcel F17 = `e(N)', hcenter bold overwritefmt
putexcel A18 = matrix(results), names nformat(number_d2)

* ANC in the first trimester 
svy: logistic SWdelivprob_convuls i.first_tri if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = AfterFristTri InFristTri _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A25 = "Received ANC in the first trimester", bold overwritefmt
putexcel E26 = "N = ", right bold overwritefmt
putexcel F26 = `e(N)', hcenter bold overwritefmt
putexcel A27 = matrix(results), names nformat(number_d2)

* BP measurement 
svy: logistic SWdelivprob_convuls i.bp if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A33 = "BP measurement at ANC", bold overwritefmt
putexcel E34 = "N = ", right bold overwritefmt
putexcel F34 = `e(N)', hcenter bold overwritefmt
putexcel A35 = matrix(results), names nformat(number_d2)

* Weight measurement 
svy: logistic SWdelivprob_convuls i.weight if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A41 = "Weight measurement at ANC", bold overwritefmt
putexcel E42 = "N = ", right bold overwritefmt
putexcel F42 = `e(N)', hcenter bold overwritefmt
putexcel A43 = matrix(results), names nformat(number_d2)

* Urine test 
svy: logistic SWdelivprob_convuls i.urine if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A49 = "Urine tested at ANC", bold overwritefmt
putexcel E50 = "N = ", right bold overwritefmt
putexcel F50 = `e(N)', hcenter bold overwritefmt
putexcel A51 = matrix(results), names nformat(number_d2)

* Blood sample
svy: logistic SWdelivprob_convuls i.blood if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A57 = "Blood sample taken at ANC", bold overwritefmt
putexcel E58 = "N = ", right bold overwritefmt
putexcel F58 = `e(N)', hcenter bold overwritefmt
putexcel A60 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls i.danger_sign_coun if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A65 = "PE-related danger sign counseling", bold overwritefmt
putexcel E66 = "N = ", right bold overwritefmt
putexcel F66 = `e(N)', hcenter bold overwritefmt
putexcel A67 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls i.tt if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A73 = "Tetanus injection at ANC", bold overwritefmt
putexcel E74 = "N = ", right bold overwritefmt
putexcel F74 = `e(N)', hcenter bold overwritefmt
putexcel A75 = matrix(results), names nformat(number_d2)

* Iron counseling at ANC 
svy: logistic SWdelivprob_convuls i.iron if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A81 = "Iron counseling at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82= `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)


*** Crude effect among women with signs of preeclampsia ***

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A1 = "ANC Frequency", bold overwritefmt
putexcel E2 = "N = ", right bold overwritefmt
putexcel F2 = `e(N)', hcenter bold overwritefmt
putexcel A3 = matrix(results), names nformat(number_d2)

* ANC provider type and location 
svy: logistic SWdelivprob_convuls i.anc_place if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = HEW PHCP_HF PHCP_hos _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A8 = "ANC Provider Type and Location", bold overwritefmt
putexcel E9 = "N = ", right bold overwritefmt
putexcel F9 = `e(N)', hcenter bold overwritefmt
putexcel A10 = matrix(results), names nformat(number_d2)

* ANC timing 
svy: logistic SWdelivprob_convuls i.ga_first_anc_cat if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = first_trimester second_trimester third_trimester missing _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A16 = "Timing at first ANC", bold overwritefmt
putexcel E17 = "N = ", right bold overwritefmt
putexcel F17 = `e(N)', hcenter bold overwritefmt
putexcel A18 = matrix(results), names nformat(number_d2)

* ANC in the first trimester 
svy: logistic SWdelivprob_convuls i.first_tri if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = AfterFirstTri InFirstTri _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A25 = "Received ANC in the first trimester", bold overwritefmt
putexcel E26 = "N = ", right bold overwritefmt
putexcel F26 = `e(N)', hcenter bold overwritefmt
putexcel A27 = matrix(results), names nformat(number_d2)

* BP measurement 
svy: logistic SWdelivprob_convuls i.bp if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A33 = "BP measurement at ANC", bold overwritefmt
putexcel E34 = "N = ", right bold overwritefmt
putexcel F34 = `e(N)', hcenter bold overwritefmt
putexcel A35 = matrix(results), names nformat(number_d2)

* Weight measurement 
svy: logistic SWdelivprob_convuls i.weight if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A41 = "Weight measurement at ANC", bold overwritefmt
putexcel E42 = "N = ", right bold overwritefmt
putexcel F42 = `e(N)', hcenter bold overwritefmt
putexcel A43 = matrix(results), names nformat(number_d2)

* Urine test 
svy: logistic SWdelivprob_convuls i.urine if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A49 = "Urine test at ANC", bold overwritefmt
putexcel E50 = "N = ", right bold overwritefmt
putexcel F50 = `e(N)', hcenter bold overwritefmt
putexcel A51 = matrix(results), names nformat(number_d2)

* Blood sample
svy: logistic SWdelivprob_convuls i.blood if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A57 = "Blood sample taken at ANC", bold overwritefmt
putexcel E58 = "N = ", right bold overwritefmt
putexcel F58 = `e(N)', hcenter bold overwritefmt
putexcel A60 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls i.danger_sign_coun if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A65 = "PE-related danger sign counseling", bold overwritefmt
putexcel E66 = "N = ", right bold overwritefmt
putexcel F66 = `e(N)', hcenter bold overwritefmt
putexcel A67 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls i.tt if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A73 = "Tetanus injection at ANC", bold overwritefmt
putexcel E74 = "N = ", right bold overwritefmt
putexcel F74 = `e(N)', hcenter bold overwritefmt
putexcel A75 = matrix(results), names nformat(number_d2)

* Iron counseling at ANC 
svy: logistic SWdelivprob_convuls i.iron if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A81 = "Iron counseling at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82= `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)

*=====================================================================
* adjusted-all results *
*=====================================================================

*==================================*
* Among all women * 
*==================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_num_cat if select
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

*** ANC provider type and location *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_place if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A23 = "ANC Provider Type and Location", bold overwritefmt
putexcel E24 = "N = ", right bold overwritefmt
putexcel F24 = `e(N)', hcenter bold overwritefmt
putexcel A25 = matrix(results), names nformat(number_d2)

*** ANC in the first trimester *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.first_tri if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A44 = "Received ANC in the first trimester", bold overwritefmt
putexcel E45 = "N = ", right bold overwritefmt
putexcel F45 = `e(N)', hcenter bold overwritefmt
putexcel A46 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.bp if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A66 = "BP measurement at ANC", bold overwritefmt
putexcel E67 = "N = ", right bold overwritefmt
putexcel F67 = `e(N)', hcenter bold overwritefmt
putexcel A68= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.weight if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A88 = "Weight measurement at ANC", bold overwritefmt
putexcel E89 = "N = ", right bold overwritefmt
putexcel F89 = `e(N)', hcenter bold overwritefmt
putexcel A90= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.urine if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A110 = "Urine test at ANC", bold overwritefmt
putexcel E111 = "N = ", right bold overwritefmt
putexcel F111 = `e(N)', hcenter bold overwritefmt
putexcel A112= matrix(results), names nformat(number_d2)

*** Blood sample *** 
svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.blood if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A132 = "Blood sample taken at ANC", bold overwritefmt
putexcel E133 = "N = ", right bold overwritefmt
putexcel F133 = `e(N)', hcenter bold overwritefmt
putexcel A134= matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.danger_sign_coun if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A154 = "PE/E danger sign counseling", bold overwritefmt
putexcel E155 = "N = ", right bold overwritefmt
putexcel F155 = `e(N)', hcenter bold overwritefmt
putexcel A156 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.tt if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A176 = "Tetanus injection", bold overwritefmt
putexcel E177 = "N = ", right bold overwritefmt
putexcel F177 = `e(N)', hcenter bold overwritefmt
putexcel A178 = matrix(results), names nformat(number_d2)

*** Counseling on iron ** 
svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.iron if select
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel A198 = "Iron counseling", bold overwritefmt
putexcel E199 = "N = ", right bold overwritefmt
putexcel F199 = `e(N)', hcenter bold overwritefmt
putexcel A200 = matrix(results), names nformat(number_d2)

*=====================================================*
* Among all women with signs of preeclampsia * 
*=====================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.anc_num_cat if select & preg_comp==1
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

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.anc_place if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A23 = "ANC Provider Type and Location", bold overwritefmt
putexcel E24 = "N = ", right bold overwritefmt
putexcel F24 = `e(N)', hcenter bold overwritefmt
putexcel A25 = matrix(results), names nformat(number_d2)

*** ANC in the first trimester *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.first_tri if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A44 = "Received ANC in the first trimester", bold overwritefmt
putexcel E45 = "N = ", right bold overwritefmt
putexcel F45 = `e(N)', hcenter bold overwritefmt
putexcel A46 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.bp if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A66 = "BP measurement at ANC", bold overwritefmt
putexcel E67 = "N = ", right bold overwritefmt
putexcel F67 = `e(N)', hcenter bold overwritefmt
putexcel A68= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.weight if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A88 = "Weight measurement at ANC", bold overwritefmt
putexcel E89 = "N = ", right bold overwritefmt
putexcel F89 = `e(N)', hcenter bold overwritefmt
putexcel A90= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.urine if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A110 = "Urine test at ANC", bold overwritefmt
putexcel E111 = "N = ", right bold overwritefmt
putexcel F111 = `e(N)', hcenter bold overwritefmt
putexcel A112= matrix(results), names nformat(number_d2)

*** Blood sample *** 
svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.blood if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A132 = "Blood sample taken at ANC", bold overwritefmt
putexcel E133 = "N = ", right bold overwritefmt
putexcel F133 = `e(N)', hcenter bold overwritefmt
putexcel A134= matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.danger_sign_coun if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A154 = "PE/E danger sign counseling", bold overwritefmt
putexcel E155 = "N = ", right bold overwritefmt
putexcel F155 = `e(N)', hcenter bold overwritefmt
putexcel A156 = matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.tt if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A176 = "Tetanus injection", bold overwritefmt
putexcel E177 = "N = ", right bold overwritefmt
putexcel F177 = `e(N)', hcenter bold overwritefmt
putexcel A178 = matrix(results), names nformat(number_d2)

*** Counseling on iron ** 
svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.iron if select & preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-1) modify
putexcel A198 = "Iron counseling", bold overwritefmt
putexcel E199 = "N = ", right bold overwritefmt
putexcel F199 = `e(N)', hcenter bold overwritefmt
putexcel A200 = matrix(results), names nformat(number_d2)


log close
