import pandas as pd
from datetime import datetime
# Read the excel file
df = pd.read_excel(r"C:\Users\user\OneDrive - American University of Beirut\AUB\Fall 25-26\VIPP 401\cleaned_data.xlsx")

# Function to calculate age from date of birth
def calculate_age(dob):
    if pd.isna(dob):
        return None
    today = datetime.today()
    return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))

# Function to check if the date is valid
def is_valid_date(date):
    invalid_dates = ['Not applicable', 'Unknown', 'No information in chart', '']
    return (pd.notna(date) and date not in invalid_dates)

# Function to check if a value is valid
def is_valid_value(value):
    invalid_values = ['Not applicable', 'No information in chart', 'Unknown','Not reported','NONE','Not yet chosen','no DMT','None (withhold)', 'none ','not reported','Not tested','x','X', '','na','none','Not applicable (this is the diagnosis spine MRI)', 'MA']
    return pd.notna(value) and value not in invalid_values

def get_duration_in_months(start_date, end_date):
    if pd.isna(start_date) or pd.isna(end_date):
        return None
    return (end_date - start_date).days / 30.44  # Average days in a month

#function to create the paragraph text
def create_first_paragraph(row):
    sentences = []
    # Sentence 1: Demographic Information
    dob = row['Date of birth']
    gender = row['Gender']
    diagnosis = row['Diagnosis']
    other_diagnosis = row['If other diagnosis, specify']
    date_of_diagnosis = row['Date of diagnosis']

    # Check if the date of diagnosis is valid
    # date_of_diagnosis_valid = is_valid_date(date_of_diagnosis)

    # Calculate age
    age = calculate_age(dob)

    # Determine the correct diagnosis
    if diagnosis == 'Other diagnosis, specify below' and is_valid_value(other_diagnosis):
        diagnosis = other_diagnosis

    # Construct the demographic sentence
    parts = []
    if is_valid_value(age):
        parts.append(f"Age:{age}. ")
    if is_valid_value(gender):
        parts.append(f"Gender:{gender.lower()}. ")
    if is_valid_value(diagnosis):
        parts.append(f"Diagnosis:{diagnosis.lower()}")

    if parts:
        demographic_sentence = ' '.join(parts) + '.'
        sentences.append(demographic_sentence)

    # Sentence 2: Conversion to SPMS
    converted_to_spms = row['Converted to SPMS?']

    # Check if the date of conversion is valid
    conversion_sentence = None
    if diagnosis == 'SPMS' or converted_to_spms in ['Yes', 'Probably']:
       conversion_sentence = "Converted to SPMS. "
    else:
        conversion_sentence = "Has not converted to SPMS."

    if conversion_sentence:
        sentences.append(conversion_sentence)


    # Sentence 3: Initial Symptoms
    date_of_first_symptoms = row['Date of first symptoms']
    initial_presentations = [
        ('Sensory', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Sensory)']),
        ('Motor weakness', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Motor weakness)']),
        ('Optic neuritis', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Optic neuritis)']),
        ('Brainstem', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Brainstem)']),
        ('Cerebellar', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Cerebellar)']),
        ('Bowel/bladder dysfunction', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Bowel/bladder dysfunction)']),
        ('Cognitive dysfunction', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Cognitive dysfunction)']),
        ('Incidental MRI findings', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Incidental MRI findings)']),
        ('Other', row['Initial presentation (provide the details of relapses in the Relapses repository instrument) (choice=Other, specify below)'])
    ]
    other_initial_presentation = row['If other initial presentation type, specify']
    progressive_symptoms = row['Progressive symptoms since onset (such as in PPMS)']
    number_of_attacks = row['Number of attacks until the first visit']


    # Construct the initial symptoms sentence
    initial_parts = []
    checked_presentations = [presentation for presentation, checked in initial_presentations if checked == 'Checked']

    if pd.notna(progressive_symptoms) and progressive_symptoms in ['Yes', 'No']:
        if progressive_symptoms == 'Yes':
            initial_parts.append("Has experienced progressive symptoms since onset.")
        if progressive_symptoms == 'No' and is_valid_value(number_of_attacks):
            initial_parts.append(f"Number of attacks until the first visit:{number_of_attacks}.")
    if initial_parts:
        initial_symptoms_sentence = ' '.join(initial_parts)
        sentences.append(initial_symptoms_sentence)

    # Sentence 4: CSF Information
    csf_sampling = row['Any CSF sampling performed?']
    if csf_sampling == 'Yes':
        csf_parts=[]

        ocbs_in_csf = row['OCBs in CSF']
        if pd.notna(ocbs_in_csf) and ocbs_in_csf in ['Negative', 'Positive']:
            csf_parts.append(f"OCBs in CSF:{ocbs_in_csf.lower()}.")

        igg_index_in_csf = row['IgG index in CSF']
        if is_valid_value(igg_index_in_csf):
            csf_parts.append(f"IgG index in CSF:{igg_index_in_csf}.")

        igg_index_category_in_csf = row['IgG index category in CSF']
        if is_valid_value(igg_index_category_in_csf):
            csf_parts.append(f"IgG index category in CSF:{igg_index_category_in_csf}.")


        if csf_parts:
            csf_sentence = ' '.join(csf_parts)
            sentences.append(csf_sentence)
        # Sentence 5: Second CSF Information
        second_csf_sampling = row['Any second (#2) CSF sampling performed?']
        if second_csf_sampling == 'Yes':

            second_csf_parts = []
            ocbs_in_second_csf = row['OCBs in second (#2) CSF']
            if pd.notna(ocbs_in_second_csf) and ocbs_in_second_csf in ['Negative', 'Positive']:
                second_csf_parts.append(f"OCBs in second CSF:{ocbs_in_second_csf.lower()}.")

            igg_index_in_second_csf = row['IgG index in second (#2) CSF']
            if is_valid_value(igg_index_in_second_csf):
                second_csf_parts.append(f"IgG index in second:{igg_index_in_second_csf}.")

            igg_index_category_in_second_csf = row['IgG index category in second (#2) CSF']
            if is_valid_value(igg_index_category_in_second_csf):
                second_csf_parts.append(f"IgG index category in second CSF:{igg_index_category_in_second_csf}.")

            if second_csf_parts:
                second_csf_sentence = ' '.join(second_csf_parts)
                sentences.append(second_csf_sentence)
            # Sentence 6: Third CSF Information
            third_csf_sampling = row['Any third (#3) CSF sampling performed?']
            if third_csf_sampling == 'Yes':
               
                third_csf_parts = []

                ocbs_in_third_csf = row['OCBs in  third (#3) CSF']
                if pd.notna(ocbs_in_third_csf) and ocbs_in_third_csf in ['Negative', 'Positive']:
                    third_csf_parts.append(f"OCBs in third CSF:{ocbs_in_third_csf.lower()}.")

                igg_index_in_third_csf = row['IgG index in  third (#3) CSF']
                if is_valid_value(igg_index_in_third_csf):
                    third_csf_parts.append(f"IgG index in third CSF:{igg_index_in_third_csf}.")

                igg_index_category_in_third_csf = row['IgG index category in  third (#3) CSF']
                if is_valid_value(igg_index_category_in_third_csf):
                    third_csf_parts.append(f"IgG index category in third CSF:{igg_index_category_in_third_csf}.")

                if third_csf_parts:
                    third_csf_sentence = ' '.join(third_csf_parts)
                    sentences.append(third_csf_sentence)

    # Sentence 7: Remarks about CSF
    remarks_about_csf = row['Any remarks about CSF?']
    if is_valid_value(remarks_about_csf):
        sentences.append(f"Remarks about CSF: {remarks_about_csf}.")

    # Sentence 8: Comorbidities
    comorbidity_sentences = []
    for i in range(1, 16):
        any_comorbidity = row[f'Any #{i} co-morbidity?']
        if any_comorbidity == 'Yes':
            if i == 1:
                first_comorbidity_present = True
            
            comorbidity_sentence = []

            name_of_comorbidity = row[f'Name of co-morbidity #{i} or illness']
            if is_valid_value(name_of_comorbidity) and name_of_comorbidity!='Other disorder, specify below':
                comorbidity_sentence.append(f"Name of co-morbidity:{name_of_comorbidity}.")

            specify_comorbidity = row[f'Specify co-morbidity #{i}']
            if is_valid_value(specify_comorbidity) and name_of_comorbidity=='Other disorder, specify below':
                comorbidity_sentence.append(f"Name of co-morbidity:{specify_comorbidity}.")
            if is_valid_value(specify_comorbidity) and name_of_comorbidity!='Other disorder, specify below':
                comorbidity_sentence.append(f"More information on co-morbidity:{specify_comorbidity}.")

            outcome_of_comorbidity = row[f'Outcome of co-morbidity #{i}']
            if is_valid_value(outcome_of_comorbidity):
                comorbidity_sentence.append(f"Outcome:{outcome_of_comorbidity.lower()}.")

            comorbidity_sentences.append(' '.join(comorbidity_sentence))
        else:
            break
    # Additional Sentence for Personal Past Medical History if First Comorbidity was Present
    if row['Any #1 co-morbidity?']=='Yes':
        personal_past_medical_history = row['Other relevant personal past medical history']
        if is_valid_value(personal_past_medical_history) and personal_past_medical_history != 'NO' :
            sentences.append(f"Past medical history: {personal_past_medical_history}.")

    if comorbidity_sentences:
        sentences.append(' '.join(comorbidity_sentences))

    
    # Sentence 10: Visit Details
    visit_details = []
    
    diagnosis_at_visit = row['Diagnosis at visit']
    if is_valid_value(diagnosis_at_visit) and diagnosis_at_visit !='Other diagnosis, specify below':
        visit_details.append(f"Visit diagnosis:{diagnosis_at_visit.lower()}.")

    other_diagnosis_at_visit = row['If other diagnosis at visit, specify']
    if is_valid_value(other_diagnosis_at_visit) and diagnosis_at_visit =='Other diagnosis, specify below':
        visit_details.append(f"Visit diagnosis:{other_diagnosis_at_visit.lower()}.")

    total_number_of_relapses = row['TOTAL number of relapses UNTIL visit (cumulative)']
    if is_valid_value(total_number_of_relapses) and total_number_of_relapses not in[999, 9999]:
        visit_details.append(f"Total relapses until visit: {total_number_of_relapses}.")
    if is_valid_value(total_number_of_relapses) and total_number_of_relapses in [999, 9999]:
        visit_details.append(f"Several relapses until this visit.")
    new_relapse_since_last_visit = row['Any NEW relapse SINCE the previous visit? (if yes, please update the \'Relapses repository\' instrument)']
    if new_relapse_since_last_visit == 'Yes':
        visit_details.append("There were new relapses since the previous visit.")

    if visit_details:
        sentences.append(' '.join(visit_details))

   

    # Sentence 19: DMT Treatment
    currently_on_dmt = row['Currently treated with any DMT?']
    if is_valid_value(currently_on_dmt) and currently_on_dmt=='Yes':
        sentences.append(f"DMT taken.")
        dmt_initiation_date = row['Date of initiation of current DMT']

        if is_valid_date(dmt_initiation_date):
            current_dmt = row['Current DMT (please update and verify \'treatments repository\' instrument)']
            
            if is_valid_value(current_dmt) and current_dmt!= 'Other DMT; specify below':
                sentences.append(f"Current DMT:{current_dmt.lower()}")
            else:
                other_current_dmt = row['If other current DMT, specify']
                if is_valid_value(other_current_dmt):
                    sentences.append(f"Current DMT: {other_current_dmt.lower()} ")
            if dmt_initiation_date!= row['Date of visit']:
                duration_on_dmt = get_duration_in_months(dmt_initiation_date, row['Date of visit'])
                sentences.append(f", for the past {duration_on_dmt:.1f} months")
            else:
                sentences.append(f", changed during this visit.")
        else:
            current_dmt = row['Current DMT (please update and verify \'treatments repository\' instrument)']
            if is_valid_value(current_dmt) and current_dmt!= 'Other DMT; specify below':
                sentences.append(f"Current DMT: {current_dmt.lower()}.")
            else:
                other_current_dmt = row['If other current DMT, specify']
                if is_valid_value(other_current_dmt):
                    sentences.append(f"Current DMT: {other_current_dmt.lower()}.")

    # Sentence 20: Comments on Previous DMTs or Other Treatments
    previous_dmts_comments = row['Comments on previous DMTs or other treatments, if applicable']
    if is_valid_value(previous_dmts_comments):
        sentences.append(f"Comments on DMT: {previous_dmts_comments}.")


    # Sentence 30: Tobacco (any form) smoker?
    tobacco_smoker = row['Tobacco (any form) smoker?']
    if is_valid_value(tobacco_smoker):
      sentences.append(f"Tobacco smoker: {tobacco_smoker} ")
    if is_valid_value(tobacco_smoker) and tobacco_smoker=='current':
      # Sentence 31: If current smoker, type of tobacco used (cigarette or cigar)
      tobacco_type_cigarette_cigar = row['If current smoker, type of tobacco is used? (choice=cigarette or cigar)']
      if is_valid_value(tobacco_type_cigarette_cigar) and tobacco_type_cigarette_cigar=='Checked':
          sentences.append(f"Tobacco type: cigarette or cigar.")

      # Sentence 32: If current smoker, type of tobacco used (water-pipe (arghile))
      tobacco_type_water_pipe = row['If current smoker, type of tobacco is used? (choice=water-pipe (arghile))']
      if is_valid_value(tobacco_type_water_pipe) and tobacco_type_water_pipe=='Checked':
          sentences.append(f"Tobacco type: water-pipe.")

      # Sentence 33: Specify if other current tobacco use
      other_tobacco_use = row['Specify if other current tobacco use']
      if is_valid_value(other_tobacco_use):
          sentences.append(f"Tobacco type: {other_tobacco_use}.")
    # Sentence 34: Number of coffee cups per week
    coffee_cups_per_week = row['Number of coffee cups per week']
    if is_valid_value(coffee_cups_per_week) and coffee_cups_per_week != 0:
        sentences.append(f"Coffee cups per week: {coffee_cups_per_week}.")

    # Sentence 35: Number of alcohol drinks per week
    alcohol_drinks_per_week = row['Number of alcohol drinks per week']
    if is_valid_value(alcohol_drinks_per_week) and alcohol_drinks_per_week != 0:
        sentences.append(f"Alcohol drinsk per week: {alcohol_drinks_per_week}.")

    # Sentence 36: Number of hours of exercise per week
    hours_of_exercise_per_week = row['Number of hours of exercise per week']
    if is_valid_value(hours_of_exercise_per_week) and hours_of_exercise_per_week != 0:
        sentences.append(f"Exercise hrs per week:{hours_of_exercise_per_week}.")

    # Sentence 38: Vitamin D level at visit (ng/ml)
    vitamin_d_level = row['Vitamin D level at visit (ng/ml)']
    if is_valid_value(vitamin_d_level):
        sentences.append(f"Vit D at visit: {vitamin_d_level}.")


    # Sentence 56: Anti-MOG Abs seropositivity
    anti_mog_seropositivity = row['Anti-MOG Abs seropositivity']
    if is_valid_value(anti_mog_seropositivity):
        sentences.append(f"Anti-MOG Abs seropositivity: {anti_mog_seropositivity}.")

    # Sentence 57: If applicable, specify the anti-MOG Abs denominator titer (1:___ )
    anti_mog_titer = row['If applicable, specify the anti-MOG Abs denominator titer (1:___ )']
    if is_valid_value(anti_mog_titer):
        sentences.append(f"Anti-MOG Abs denominator titer: 1:{anti_mog_titer}.")

    

    # Sentence 64: Any MRI of the brain or spine performed for this visit?
    mri_performed = row['Any MRI of the brain or spine performed for this visit?']
    if is_valid_value(mri_performed) and mri_performed=='Yes':

        # Sentence 66: Received Gadolinium for Brain MRI?
        gadolinium_brain = row['Received Gadolinium for Brain MRI?']
        if is_valid_value(gadolinium_brain) and gadolinium_brain=='Yes':
            sentences.append(f"Received Gadolinium for Brain MRI.")

        # Sentence 67: Presence of new lesions on BRAIN MRI at visit?
        new_lesions_brain = row['Presence of new lesions on BRAIN MRI at visit? ']
        if is_valid_value(new_lesions_brain) and new_lesions_brain=='Yes':
            sentences.append(f"Presence of new lesions on Brain MRI.")
        if is_valid_value(new_lesions_brain) and new_lesions_brain=='No':
            sentences.append(f"No new lesions on Brain MRI.")

        # Sentence 68: Date of comparison Brain MRI
        #comparison_brain_mri_date = row['Date of comparison Brain MRI']
        #if is_valid_value(comparison_brain_mri_date):
        #   sentences.append(f"Date of comparison Brain MRI: {comparison_brain_mri_date}.")

        # Sentence 69: Number of new T2 lesions in the Brain
        new_t2_lesions_brain = row['Number of new T2 lesions in the Brain']
        if is_valid_value(new_lesions_brain) and new_lesions_brain=='Yes':
          if is_valid_value(new_t2_lesions_brain):
              sentences.append(f"Number of new Brain T2 lesions: {new_t2_lesions_brain}.")

        # Sentence 70: Total number of T1 Gad-enhancing lesions in the Brain
        t1_gad_lesions_brain = row['Total number of T1 Gad-enhancing lesions in the Brain']
        if is_valid_value(gadolinium_brain) and gadolinium_brain=='Yes':
          if is_valid_value(t1_gad_lesions_brain):
              sentences.append(f"Total T1 Gad-enhancing lesions in Brain: {t1_gad_lesions_brain}.")

        # Sentence 71: Brain MRI lesion load
        brain_mri_lesion_load = row['Brain MRI lesion load']
        if is_valid_value(brain_mri_lesion_load):
            sentences.append(f"Brain MRI lesion load: {brain_mri_lesion_load}.")

        # Sentence 72: Date of Cervical spine MRI for this visit
        # cervical_mri_date = row['Date of Cervical spine MRI for this visit']
        # if is_valid_date(cervical_mri_date):
        #     sentences.append(f"Cervical spine MRI scheduled.")
        # else:
        #     sentences.append(f"Cervical spine MRI not scheduled.")

        # # Sentence 73: Received Gadolinium for Cervical spine MRI?
        gadolinium_cervical = row['Received Gadolinium for Cervical spine MRI?']
        if is_valid_value(gadolinium_cervical) and gadolinium_cervical=='Yes':
            sentences.append(f"Received Gadolinium for Cervical spine MRI.")

        # Sentence 74: CERVICAL Spine MRI- any NEW lesion (enhancing or not)
        new_lesions_cervical = row['CERVICAL Spine MRI- any NEW lesion (enhancing or not)']
        if is_valid_value(new_lesions_cervical):
            sentences.append(f"Cervical spine MRI: {new_lesions_cervical}.")

        # Sentence 75: Number of new T2 lesions in the CERVICAL SPINE
        new_t2_lesions_cervical = row['Number of new T2 lesions in the CERVICAL SPINE']
        if is_valid_value(new_lesions_cervical) and new_lesions_cervical in ['T2 lesions, not gad-enhancing', 'Both enhancing and nonenhancing lesions', 'T2 lesion, unknown enhancement (MRI without gad)']:
          if is_valid_value(new_t2_lesions_cervical):
              sentences.append(f"Number of new T2 lesions in Cervical spine: {new_t2_lesions_cervical}.")

        # Sentence 76: Total number of T1 Gad-enhancing lesions in the CERVICAL spine
        t1_gad_lesions_cervical = row['Total number of T1 Gad-enhancing lesions in the CERVICAL spine']
        if is_valid_value(gadolinium_cervical) and gadolinium_cervical=='Yes':
          if is_valid_value(t1_gad_lesions_cervical):
              sentences.append(f"Total T1 Gad-enhancing lesions in Cervical spine: {t1_gad_lesions_cervical}.")


        # Sentence 78: Received Gadolinium for Thoracic (dorsal) spine MRI?
        gadolinium_thoracic = row['Received Gadolinium for Thoracic (dorsal) spine MRI?']
        if is_valid_value(gadolinium_thoracic) and gadolinium_thoracic=='Yes':
            sentences.append(f"Received Gadolinium for Thoracic (dorsal) spine MRI.")

        # Sentence 79: DORSAL/THORACIC Spine MRI- any NEW lesion (enhancing or not)
        new_lesions_thoracic = row['DORSAL/THORACIC Spine MRI- any NEW lesion (enhancing or not)']
        if is_valid_value(new_lesions_thoracic):
            sentences.append(f"Thoracic spine MRI: {new_lesions_thoracic}.")

        # Sentence 80: Number of new T2 lesions in the DORSAL/THORACIC SPINE
        new_t2_lesions_thoracic = row['Number of new T2 lesions in the DORSAL/THORACIC SPINE']
        if is_valid_value(new_lesions_thoracic) and new_lesions_thoracic in ['T2 lesions, not gad-enhancing', 'Both enhancing and nonenhancing lesions', 'T2 lesion, unknown enhancement (MRI without gad)']:
          if is_valid_value(new_t2_lesions_thoracic):
              sentences.append(f"Number of new T2 lesions in  Thoracic spine: {new_t2_lesions_thoracic}.")

        # Sentence 81: Total number of T1 Gad-enhancing lesions in the DORSAL/THORACIC spine
        t1_gad_lesions_thoracic = row['Total number of T1 Gad-enhancing lesions in the DORSAL/THORACIC spine']
        if is_valid_value(gadolinium_thoracic) and gadolinium_thoracic=='Yes':
          if is_valid_value(t1_gad_lesions_thoracic):
              sentences.append(f"Total T1 Gad-enhancing lesions in Thoracic spine: {t1_gad_lesions_thoracic}.")

        # Sentence 82: Evidence of longitudinally extensive CERVICAL or DORSAL/THORACIC cord lesion (LETM)?
        letm_evidence = row['Evidence of longitudinally extensive CERVICAL or DORSAL/THORACIC cord lesion (LETM)?']
        if is_valid_value(letm_evidence) and letm_evidence=='Yes':
            sentences.append(f"Evidence of LETM.")

        # Sentence 83: Whole Spinal Cord (Cervical+Dorsal) lesion load on MRI (If only Cervical or Dorsal spine MRI done, refer to the field note below)
        whole_spinal_lesion_load = row['Whole Spinal Cord (Cervical+Dorsal) lesion load on MRI (If only Cervical or Dorsal spine MRI done, refer to the field note below)']
        if is_valid_value(whole_spinal_lesion_load):
            sentences.append(f"Whole spinal cord lesion load on MRI: {whole_spinal_lesion_load}.")

        # Sentence 84: Any remarks about MRI (optic nerves or meningeal enhancement, cortical lesions, spinal cord compression, etc...)
        mri_remarks = row['Any remarks about MRI (optic nerves or meningeal enhancement, cortical lesions, spinal cord compression, etc...?']
        if is_valid_value(mri_remarks):
            sentences.append(f"Remarks about MRI: {mri_remarks}.")

    # Sentence 104: Was the 25-foot walk test attempted at this visit?
    walk_test_attempted = row['Was the 25-foot walk test attempted at this visit?']
    if is_valid_value(walk_test_attempted) and walk_test_attempted=='yes':

        # Sentence 105: Any assistance required for 25-FWT?
        assistance_required = row['Any assistance required for 25-FWT?']
        if is_valid_value(assistance_required):
            sentences.append(f"Assistance required for 25-foot walk test: {assistance_required}.")

        # Sentence 106: Best 25-foot walk test time (in seconds)
        walk_test_time = row['Best 25-foot walk test time (in seconds)']
        if is_valid_value(walk_test_time):
            sentences.append(f"25-foot walk test time was {walk_test_time} sec.")

    # Sentence 107: Was the SDMT done at this visit?
    sdmt_done = row['Was the SDMT done at this visit?']
    if is_valid_value(sdmt_done) and sdmt_done=='yes':
        # sentences.append(f"SDMT done.")

        # Sentence 108: SDMT numerator/denominator
        sdmt_denominator = row['SDMT denominator']
        sdmt_numerator = row['SDMT numerator']
        if is_valid_value(sdmt_numerator):
            sentences.append(f"SDMT numerator/denominator: {sdmt_numerator}/{sdmt_denominator}.")

        # # Sentence 109: SDMT denominator
        # sdmt_denominator = row['SDMT denominator']
        # if is_valid_value(sdmt_denominator):
        #     sentences.append(f"SDMT denominator: {sdmt_denominator}.")

    # Sentence 110: Was the 9-HPT done at this visit?
    hpt_done = row['Was the 9-HPT done at this visit?']
    if is_valid_value(hpt_done) and hpt_done=='yes':
        # sentences.append(f"9-HPT done.")

        # Sentence 111: Was the patient able to perform the 9HPT?
        hpt_able = row['Was the patient able to perform the 9HPT?']
        if is_valid_value(hpt_able):
            sentences.append(f"9-HPT ability: {hpt_able}.")

        # Sentence 112: Dominant hand
        dominant_hand = row['Dominant hand']
        if is_valid_value(dominant_hand):
            sentences.append(f"Dominant hand:{dominant_hand}.")

        # Sentence 113: Dominant hand 9-hole peg test time (seconds), if able to perform the test
        dominant_hand_time = row['Dominant hand 9-hole peg test time (seconds), if able to perform the test']
        if is_valid_value(dominant_hand_time):
            sentences.append(f" 9-HPT dominant:{dominant_hand_time} sec.")

        # Sentence 114: Non-dominant hand 9-hole peg test time (seconds), if able to perform the test
        non_dominant_hand_time = row['Non-dominant hand 9-hole peg test time (seconds), if able to perform the test']
        if is_valid_value(non_dominant_hand_time):
            sentences.append(f"9-HPT non-dominant:{non_dominant_hand_time} sec.")

    # Sentence 115: Was the EDSS performed for this visit?
    edss_performed = row['Was the EDSS performed for this visit?']
    if is_valid_value(edss_performed) and edss_performed=='Yes':
        # Sentence 124: EDSS at visit
        edss_at_visit = row['EDSS at visit']
        if is_valid_value(edss_at_visit):
            sentences.append(f"EDSS:{edss_at_visit}.")

    # Sentence 125: Any change in DMT or management at the end of visit?
    change_in_dmt_management = row['Any change in DMT or management at the end of visit?']
    if is_valid_value(change_in_dmt_management) and change_in_dmt_management=='yes- change of DMT was indicated':
        sentences.append(f"Change in DMT or management done.")

        # Sentence 126: Main reason for change in management and DMT at visit
        main_reason_for_change = row['Main reason for change in management and DMT at visit']
        if is_valid_value(main_reason_for_change) and main_reason_for_change!='other reasons':
            sentences.append(f"Reason for DMT change:{main_reason_for_change}.")

        # Sentence 127: If other reasons for change in DMT, specify
        other_reasons_for_change = row['If other reasons for change in DMT, specify']
        if is_valid_value(other_reasons_for_change) and main_reason_for_change=='other reasons':
            sentences.append(f"Reason for DMT change:{other_reasons_for_change}.")

        # Sentence 128: New DMT indicated
        new_dmt_indicated = row['New DMT indicated']
        if is_valid_value(new_dmt_indicated) and  new_dmt_indicated!='Other DMT; specify below':
            sentences.append(f"New DMT:{new_dmt_indicated}.")

        # Sentence 129: If other new DMT indicated, specify
        other_new_dmt_indicated = row['If other new DMT indicated, specify']
        if is_valid_value(other_new_dmt_indicated) and  new_dmt_indicated=='Other DMT; specify below':
            sentences.append(f"New DMT:{other_new_dmt_indicated}.")



    # Sentence 143: NEDA at visit?
    neda_at_visit = row['NEDA at visit?']
    reason_neda_loss = row['Reason of NEDA loss']
    if neda_at_visit =='No':
      if is_valid_value(reason_neda_loss):
        sentences.append(f"The patient is in EDA due to {reason_neda_loss}.")
      else:
        sentences.append(f"The patient is in EDA.")
    else:
        sentences.append("The patient is in NEDA.")
    # Combine all sentences into a single paragraph
    paragraph = ' '.join(sentences)

    return paragraph


#function to create the paragraph text
def create_second_paragraph(row):
    sentences = []
       
    # Sentence 10: Visit Details
    visit_details = []
    
    diagnosis_at_visit = row['Diagnosis at visit']
    if is_valid_value(diagnosis_at_visit) and diagnosis_at_visit !='Other diagnosis, specify below':
        visit_details.append(f"Visit diagnosis:{diagnosis_at_visit.lower()}.")

    other_diagnosis_at_visit = row['If other diagnosis at visit, specify']
    if is_valid_value(other_diagnosis_at_visit) and diagnosis_at_visit =='Other diagnosis, specify below':
        visit_details.append(f"Visit diagnosis:{other_diagnosis_at_visit.lower()}.")

    total_number_of_relapses = row['TOTAL number of relapses UNTIL visit (cumulative)']
    if is_valid_value(total_number_of_relapses) and total_number_of_relapses not in[999, 9999]:
        visit_details.append(f"Total relapses until visit: {total_number_of_relapses}.")
    if is_valid_value(total_number_of_relapses) and total_number_of_relapses in [999, 9999]:
        visit_details.append(f"Several relapses until this visit.")
    new_relapse_since_last_visit = row['Any NEW relapse SINCE the previous visit? (if yes, please update the \'Relapses repository\' instrument)']
    if new_relapse_since_last_visit == 'Yes':
        visit_details.append("There were new relapses since the previous visit.")

    if visit_details:
        sentences.append(' '.join(visit_details))

   

    # Sentence 19: DMT Treatment
    currently_on_dmt = row['Currently treated with any DMT?']
    if is_valid_value(currently_on_dmt) and currently_on_dmt=='Yes':
        sentences.append(f"DMT taken.")
        dmt_initiation_date = row['Date of initiation of current DMT']
        if is_valid_date(dmt_initiation_date):
            current_dmt = row['Current DMT (please update and verify \'treatments repository\' instrument)']
            if is_valid_value(current_dmt) and current_dmt!= 'Other DMT; specify below':
                sentences.append(f"Current DMT:{current_dmt.lower()} on {dmt_initiation_date.strftime('%B %d, %Y')}.")
            else:
                other_current_dmt = row['If other current DMT, specify']
                if is_valid_value(other_current_dmt):
                    sentences.append(f"Current DMT: {other_current_dmt.lower()} on {dmt_initiation_date.strftime('%B %d, %Y')}.")
        else:
            current_dmt = row['Current DMT (please update and verify \'treatments repository\' instrument)']
            if is_valid_value(current_dmt) and current_dmt!= 'Other DMT; specify below':
                sentences.append(f"Current DMT: {current_dmt.lower()}.")
            else:
                other_current_dmt = row['If other current DMT, specify']
                if is_valid_value(other_current_dmt):
                    sentences.append(f"Current DMT: {other_current_dmt.lower()}.")

    # Sentence 20: Comments on Previous DMTs or Other Treatments
    previous_dmts_comments = row['Comments on previous DMTs or other treatments, if applicable']
    if is_valid_value(previous_dmts_comments):
        sentences.append(f"Comments on DMT: {previous_dmts_comments}.")


    # Sentence 30: Tobacco (any form) smoker?
    tobacco_smoker = row['Tobacco (any form) smoker?']
    if is_valid_value(tobacco_smoker):
      sentences.append(f"Tobacco smoker: {tobacco_smoker} ")
    if is_valid_value(tobacco_smoker) and tobacco_smoker=='current':
      # Sentence 31: If current smoker, type of tobacco used (cigarette or cigar)
      tobacco_type_cigarette_cigar = row['If current smoker, type of tobacco is used? (choice=cigarette or cigar)']
      if is_valid_value(tobacco_type_cigarette_cigar) and tobacco_type_cigarette_cigar=='Checked':
          sentences.append(f"Tobacco type: cigarette or cigar.")

      # Sentence 32: If current smoker, type of tobacco used (water-pipe (arghile))
      tobacco_type_water_pipe = row['If current smoker, type of tobacco is used? (choice=water-pipe (arghile))']
      if is_valid_value(tobacco_type_water_pipe) and tobacco_type_water_pipe=='Checked':
          sentences.append(f"Tobacco type: water-pipe.")

      # Sentence 33: Specify if other current tobacco use
      other_tobacco_use = row['Specify if other current tobacco use']
      if is_valid_value(other_tobacco_use):
          sentences.append(f"Tobacco type: {other_tobacco_use}.")
    # Sentence 34: Number of coffee cups per week
    coffee_cups_per_week = row['Number of coffee cups per week']
    if is_valid_value(coffee_cups_per_week) and coffee_cups_per_week != 0:
        sentences.append(f"Coffee cups per week: {coffee_cups_per_week}.")

    # Sentence 35: Number of alcohol drinks per week
    alcohol_drinks_per_week = row['Number of alcohol drinks per week']
    if is_valid_value(alcohol_drinks_per_week) and alcohol_drinks_per_week != 0:
        sentences.append(f"Alcohol drinsk per week: {alcohol_drinks_per_week}.")

    # Sentence 36: Number of hours of exercise per week
    hours_of_exercise_per_week = row['Number of hours of exercise per week']
    if is_valid_value(hours_of_exercise_per_week) and hours_of_exercise_per_week != 0:
        sentences.append(f"Exercise hrs per week:{hours_of_exercise_per_week}.")

    # Sentence 38: Vitamin D level at visit (ng/ml)
    vitamin_d_level = row['Vitamin D level at visit (ng/ml)']
    if is_valid_value(vitamin_d_level):
        sentences.append(f"Vit D at visit: {vitamin_d_level}.")


    # Sentence 56: Anti-MOG Abs seropositivity
    anti_mog_seropositivity = row['Anti-MOG Abs seropositivity']
    if is_valid_value(anti_mog_seropositivity):
        sentences.append(f"Anti-MOG Abs seropositivity: {anti_mog_seropositivity}.")

    # Sentence 57: If applicable, specify the anti-MOG Abs denominator titer (1:___ )
    anti_mog_titer = row['If applicable, specify the anti-MOG Abs denominator titer (1:___ )']
    if is_valid_value(anti_mog_titer):
        sentences.append(f"Anti-MOG Abs denominator titer: 1:{anti_mog_titer}.")

    

    # Sentence 64: Any MRI of the brain or spine performed for this visit?
    mri_performed = row['Any MRI of the brain or spine performed for this visit?']
    if is_valid_value(mri_performed) and mri_performed=='Yes':
        # sentences.append(f"MRI of brain or spine done.")

        # Sentence 66: Received Gadolinium for Brain MRI?
        gadolinium_brain = row['Received Gadolinium for Brain MRI?']
        if is_valid_value(gadolinium_brain) and gadolinium_brain=='Yes':
            sentences.append(f"Received Gadolinium for Brain MRI.")

        # Sentence 67: Presence of new lesions on BRAIN MRI at visit?
        new_lesions_brain = row['Presence of new lesions on BRAIN MRI at visit? ']
        if is_valid_value(new_lesions_brain) and new_lesions_brain=='Yes':
            sentences.append(f"Presence of new lesions on Brain MRI.")
        if is_valid_value(new_lesions_brain) and new_lesions_brain=='No':
            sentences.append(f"No new lesions on Brain MRI.")

        # Sentence 68: Date of comparison Brain MRI
        #comparison_brain_mri_date = row['Date of comparison Brain MRI']
        #if is_valid_value(comparison_brain_mri_date):
        #   sentences.append(f"Date of comparison Brain MRI: {comparison_brain_mri_date}.")

        # Sentence 69: Number of new T2 lesions in the Brain
        new_t2_lesions_brain = row['Number of new T2 lesions in the Brain']
        if is_valid_value(new_lesions_brain) and new_lesions_brain=='Yes':
          if is_valid_value(new_t2_lesions_brain):
              sentences.append(f"Number of new Brain T2 lesions: {new_t2_lesions_brain}.")

        # Sentence 70: Total number of T1 Gad-enhancing lesions in the Brain
        t1_gad_lesions_brain = row['Total number of T1 Gad-enhancing lesions in the Brain']
        if is_valid_value(gadolinium_brain) and gadolinium_brain=='Yes':
          if is_valid_value(t1_gad_lesions_brain):
              sentences.append(f"Total T1 Gad-enhancing lesions in Brain: {t1_gad_lesions_brain}.")

        # Sentence 71: Brain MRI lesion load
        brain_mri_lesion_load = row['Brain MRI lesion load']
        if is_valid_value(brain_mri_lesion_load):
            sentences.append(f"Brain MRI lesion load: {brain_mri_lesion_load}.")

        # Sentence 72: Date of Cervical spine MRI for this visit
        cervical_mri_date = row['Date of Cervical spine MRI for this visit']
        # if is_valid_date(cervical_mri_date):
        #     sentences.append(f"Cervical spine MRI scheduled)")
        # else:
        #     sentences.append(f"Cervical spine MRI not scheduled.")

        # Sentence 73: Received Gadolinium for Cervical spine MRI?
        gadolinium_cervical = row['Received Gadolinium for Cervical spine MRI?']
        if is_valid_value(gadolinium_cervical) and gadolinium_cervical=='Yes':
            sentences.append(f"Received Gadolinium for Cervical spine MRI.")

        # Sentence 74: CERVICAL Spine MRI- any NEW lesion (enhancing or not)
        new_lesions_cervical = row['CERVICAL Spine MRI- any NEW lesion (enhancing or not)']
        if is_valid_value(new_lesions_cervical):
            sentences.append(f"Cervical spine MRI: {new_lesions_cervical}.")

        # Sentence 75: Number of new T2 lesions in the CERVICAL SPINE
        new_t2_lesions_cervical = row['Number of new T2 lesions in the CERVICAL SPINE']
        if is_valid_value(new_lesions_cervical) and new_lesions_cervical in ['T2 lesions, not gad-enhancing', 'Both enhancing and nonenhancing lesions', 'T2 lesion, unknown enhancement (MRI without gad)']:
          if is_valid_value(new_t2_lesions_cervical):
              sentences.append(f"Number of new T2 lesions in Cervical spine: {new_t2_lesions_cervical}.")

        # Sentence 76: Total number of T1 Gad-enhancing lesions in the CERVICAL spine
        t1_gad_lesions_cervical = row['Total number of T1 Gad-enhancing lesions in the CERVICAL spine']
        if is_valid_value(gadolinium_cervical) and gadolinium_cervical=='Yes':
          if is_valid_value(t1_gad_lesions_cervical):
              sentences.append(f"Total T1 Gad-enhancing lesions in Cervical spine: {t1_gad_lesions_cervical}.")


        # Sentence 78: Received Gadolinium for Thoracic (dorsal) spine MRI?
        gadolinium_thoracic = row['Received Gadolinium for Thoracic (dorsal) spine MRI?']
        if is_valid_value(gadolinium_thoracic) and gadolinium_thoracic=='Yes':
            sentences.append(f"Received Gadolinium for Thoracic (dorsal) spine MRI.")

        # Sentence 79: DORSAL/THORACIC Spine MRI- any NEW lesion (enhancing or not)
        new_lesions_thoracic = row['DORSAL/THORACIC Spine MRI- any NEW lesion (enhancing or not)']
        if is_valid_value(new_lesions_thoracic):
            sentences.append(f"Thoracic spine MRI: {new_lesions_thoracic}.")

        # Sentence 80: Number of new T2 lesions in the DORSAL/THORACIC SPINE
        new_t2_lesions_thoracic = row['Number of new T2 lesions in the DORSAL/THORACIC SPINE']
        if is_valid_value(new_lesions_thoracic) and new_lesions_thoracic in ['T2 lesions, not gad-enhancing', 'Both enhancing and nonenhancing lesions', 'T2 lesion, unknown enhancement (MRI without gad)']:
          if is_valid_value(new_t2_lesions_thoracic):
              sentences.append(f"Number of new T2 lesions in  Thoracic spine: {new_t2_lesions_thoracic}.")

        # Sentence 81: Total number of T1 Gad-enhancing lesions in the DORSAL/THORACIC spine
        t1_gad_lesions_thoracic = row['Total number of T1 Gad-enhancing lesions in the DORSAL/THORACIC spine']
        if is_valid_value(gadolinium_thoracic) and gadolinium_thoracic=='Yes':
          if is_valid_value(t1_gad_lesions_thoracic):
              sentences.append(f"Total T1 Gad-enhancing lesions in Thoracic spine: {t1_gad_lesions_thoracic}.")

        # Sentence 82: Evidence of longitudinally extensive CERVICAL or DORSAL/THORACIC cord lesion (LETM)?
        letm_evidence = row['Evidence of longitudinally extensive CERVICAL or DORSAL/THORACIC cord lesion (LETM)?']
        if is_valid_value(letm_evidence) and letm_evidence=='Yes':
            sentences.append(f"Evidence of LETM.")

        # Sentence 83: Whole Spinal Cord (Cervical+Dorsal) lesion load on MRI (If only Cervical or Dorsal spine MRI done, refer to the field note below)
        whole_spinal_lesion_load = row['Whole Spinal Cord (Cervical+Dorsal) lesion load on MRI (If only Cervical or Dorsal spine MRI done, refer to the field note below)']
        if is_valid_value(whole_spinal_lesion_load):
            sentences.append(f"Whole spinal cord lesion load on MRI: {whole_spinal_lesion_load}.")

        # Sentence 84: Any remarks about MRI (optic nerves or meningeal enhancement, cortical lesions, spinal cord compression, etc...)
        mri_remarks = row['Any remarks about MRI (optic nerves or meningeal enhancement, cortical lesions, spinal cord compression, etc...?']
        if is_valid_value(mri_remarks):
            sentences.append(f"Remarks about MRI: {mri_remarks}.")

    # Sentence 104: Was the 25-foot walk test attempted at this visit?
    walk_test_attempted = row['Was the 25-foot walk test attempted at this visit?']
    if is_valid_value(walk_test_attempted) and walk_test_attempted=='yes':

        # Sentence 105: Any assistance required for 25-FWT?
        assistance_required = row['Any assistance required for 25-FWT?']
        if is_valid_value(assistance_required):
            sentences.append(f"Assistance required for 25-foot walk test: {assistance_required}.")

        # Sentence 106: Best 25-foot walk test time (in seconds)
        walk_test_time = row['Best 25-foot walk test time (in seconds)']
        if is_valid_value(walk_test_time):
            sentences.append(f"25-foot walk test time was {walk_test_time} sec.")

    # Sentence 107: Was the SDMT done at this visit?
    sdmt_done = row['Was the SDMT done at this visit?']
    if is_valid_value(sdmt_done) and sdmt_done=='yes':

        # Sentence 108: SDMT numerator
        # sdmt_numerator = row['SDMT numerator']
        # if is_valid_value(sdmt_numerator):
        #     sentences.append(f"SDMT numerator: {sdmt_numerator}.")

        # Sentence 109: SDMT numerator/denominator
        sdmt_numerator = row['SDMT numerator']
        sdmt_denominator = row['SDMT denominator']
        if is_valid_value(sdmt_denominator):
            sentences.append(f"SDMT numerator/denominator:{sdmt_numerator}/{sdmt_denominator}.")
        
    # Sentence 110: Was the 9-HPT done at this visit?
    hpt_done = row['Was the 9-HPT done at this visit?']
    if is_valid_value(hpt_done) and hpt_done=='yes':

        # Sentence 111: Was the patient able to perform the 9HPT?
        hpt_able = row['Was the patient able to perform the 9HPT?']
        if is_valid_value(hpt_able):
            sentences.append(f"9-HPT ability: {hpt_able}.")

        # Sentence 112: Dominant hand
        dominant_hand = row['Dominant hand']
        if is_valid_value(dominant_hand):
            sentences.append(f"Dominant hand:{dominant_hand}.")

        # Sentence 113: Dominant hand 9-hole peg test time (seconds), if able to perform the test
        dominant_hand_time = row['Dominant hand 9-hole peg test time (seconds), if able to perform the test']
        if is_valid_value(dominant_hand_time):
            sentences.append(f" 9-HPT dominant:{dominant_hand_time} sec.")

        # Sentence 114: Non-dominant hand 9-hole peg test time (seconds), if able to perform the test
        non_dominant_hand_time = row['Non-dominant hand 9-hole peg test time (seconds), if able to perform the test']
        if is_valid_value(non_dominant_hand_time):
            sentences.append(f"9-HPT non-dominant:{non_dominant_hand_time} sec.")

    # Sentence 115: Was the EDSS performed for this visit?
    edss_performed = row['Was the EDSS performed for this visit?']
    if is_valid_value(edss_performed) and edss_performed=='Yes':
        # Sentence 124: EDSS at visit
        edss_at_visit = row['EDSS at visit']
        if is_valid_value(edss_at_visit):
            sentences.append(f"EDSS:{edss_at_visit}.")

    # Sentence 125: Any change in DMT or management at the end of visit?
    change_in_dmt_management = row['Any change in DMT or management at the end of visit?']
    if is_valid_value(change_in_dmt_management) and change_in_dmt_management=='yes- change of DMT was indicated':
        sentences.append(f"Change in DMT or management done.")

        # Sentence 126: Main reason for change in management and DMT at visit
        main_reason_for_change = row['Main reason for change in management and DMT at visit']
        if is_valid_value(main_reason_for_change) and main_reason_for_change!='other reasons':
            sentences.append(f"Reason for DMT change:{main_reason_for_change}.")

        # Sentence 127: If other reasons for change in DMT, specify
        other_reasons_for_change = row['If other reasons for change in DMT, specify']
        if is_valid_value(other_reasons_for_change) and main_reason_for_change=='other reasons':
            sentences.append(f"Reason for DMT change:{other_reasons_for_change}.")

        # Sentence 128: New DMT indicated
        new_dmt_indicated = row['New DMT indicated']
        if is_valid_value(new_dmt_indicated) and  new_dmt_indicated!='Other DMT; specify below':
            sentences.append(f"New DMT:{new_dmt_indicated}.")

        # Sentence 129: If other new DMT indicated, specify
        other_new_dmt_indicated = row['If other new DMT indicated, specify']
        if is_valid_value(other_new_dmt_indicated) and  new_dmt_indicated=='Other DMT; specify below':
            sentences.append(f"New DMT:{other_new_dmt_indicated}.")



    # Sentence 143: NEDA at visit?
    neda_at_visit = row['NEDA at visit?']
    reason_neda_loss = row['Reason of NEDA loss']
    if neda_at_visit =='No':
      if is_valid_value(reason_neda_loss):
        sentences.append(f"The patient is in EDA due to {reason_neda_loss}.")
      else:
        sentences.append(f"The patient is in EDA.")
    else:
        sentences.append("The patient is in NEDA.")
    # Combine all sentences into a single paragraph
    paragraph = ' '.join(sentences)

    return paragraph

#function to create the paragraph text
def create_last_paragraph(row):
    sentences = []
    neda_at_visit = row['NEDA at visit?']
    reason_neda_loss = row['Reason of NEDA loss']
    if neda_at_visit =='No':
        sentences.append(f"The patient is in EDA.")
    else:
        sentences.append("The patient is in NEDA.")
    # Combine all sentences into a single paragraph
    paragraph = ' '.join(sentences)

    return paragraph

# Apply the function to generate paragraph text for each row
df['paragraph_first'] = df.apply(create_first_paragraph, axis=1)
df['paragraph_between']=df.apply(create_second_paragraph, axis=1)
df['paragraph_out']=df.apply(create_last_paragraph, axis=1)
# Convert the 'Date of visit' column to datetime
df['Date of visit'] = pd.to_datetime(df['Date of visit'], errors='coerce')

# Extract the date part
df['Date of visit'] = df['Date of visit'].dt.date
# Select columns for the output file
columns_to_keep = ['MSC research database ID', 'Repeat Instance', 'Date of visit', 'NEDA at visit?', 'paragraph_first', 'paragraph_between', 'paragraph_out']
output_df = df[columns_to_keep]

# Save the DataFrame to a new Excel file
output_df.to_excel("cleaned_Lynn.xlsx", index=False)