import pandas as pd
import random

# Define more unique South African and English names and surnames
first_names_unique = [
    'Thulani', 'Anele', 'Siyabonga', 'Khanyisile', 'Sipho', 'Lindiwe', 'Bheki', 'Phumelele', 'Nomfundo', 'Themba', 'Lerato', 'Kagiso',
    'Mpho', 'Ayanda', 'Zinhle', 'Mandisa', 'Sibusiso', 'Nokuthula', 'Lwazi', 'Ntombi', 'John', 'Sarah', 'Michael', 'Emily', 'David', 
    'Jessica', 'Mark', 'Sophia', 'James', 'Grace', 'William', 'Emma', 'Ethan', 'Olivia', 'Daniel', 'Ava', 'Matthew', 'Isabella', 
    'Henry', 'Charlotte'
]

surnames_unique = [
    'Mthembu', 'Ndlovu', 'Zuma', 'Dlamini', 'Shabalala', 'Gumede', 'Ngcobo', 'Nkosi', 'Zondo', 'Khumalo', 'Motsepe', 'Moloi', 'Mabaso', 
    'Mkhize', 'Zulu', 'Mdletshe', 'Buthelezi', 'Mokoena', 'Maponya', 'Mahlangu', 'Smith', 'Johnson', 'Brown', 'Williams', 'Jones', 
    'Taylor', 'Davis', 'Miller', 'Wilson', 'Moore', 'Anderson', 'Thomas', 'Jackson', 'White', 'Harris', 'Martin', 'Thompson', 'Garcia', 
    'Clark', 'Lewis'
]

# Ensure names and surnames don't repeat by randomly shuffling and using each only once
random.shuffle(first_names_unique)
random.shuffle(surnames_unique)

# Define possible IT skills, soft skills, and program choices
it_skills = ['Python', 'SQL', 'JavaScript', 'C++', 'PowerShell', 'Linux', 'HTML', 'Docker', 'CSS', 'Java']
soft_skills = ['Leadership', 'Communication', 'Problem Solving', 'Teamwork', 'Adaptability', 'Creativity', 'Flexibility']
program_choices = {
    1: ['Cyber', 'Data Analytics', 'SAP'],
    2: ['Software Development', 'DevOps', 'Cloud Computing'],
    3: ['AI/ML', 'Salesforce', 'Software Testing'],
    4: ['IT Support', 'IT Technician', 'Networking'],
    5: ['Data', 'Robotics', 'Web Development'],
    6: ['BI', 'Cloud Computing', 'SAP'],
    7: ['Business Analytics', 'Cloud Computing', 'Systems Analyst'],
    8: ['Web Development', 'UI/UX', 'React']
}
risk_statuses = ['Non-attendance', 'Behaviour', 'Tech Learning', 'Personal', 'Not at risk']
placeabilities = ['Ready', 'Will be ready', 'Potentially unplaceable']
hackathon_results = ['Winner', 'Second place', 'Third place', 'Participated', 'Not participated']

# Generate a random candidate row with unique names and surnames
def generate_unique_candidate(candidate_key, program_number, first_name, surname):
    qualification = random.choice(['Degree', 'Diploma'])
    previous_program = random.choice(['Yes', 'No'])
    
    # Scores
    candidate_manager_score = random.randint(30, 90)
    tech_score = random.randint(30, 90)

    # Recommendations based on scores
    if candidate_manager_score < 50:
        candidate_manager_recommendation = 'Negative'
    elif 50 <= candidate_manager_score < 65:
        candidate_manager_recommendation = 'Neutral'
    elif 65 <= candidate_manager_score < 75:
        candidate_manager_recommendation = 'Positive'
    else:
        candidate_manager_recommendation = 'Outstanding'
    
    if tech_score < 50:
        tech_mentor_recommendation = 'Negative'
    elif 50 <= tech_score < 65:
        tech_mentor_recommendation = 'Neutral'
    elif 65 <= tech_score < 75:
        tech_mentor_recommendation = 'Positive'
    else:
        tech_mentor_recommendation = 'Outstanding'

    # Skills
    random_it_skills = random.sample(it_skills, 5)
    random_soft_skills = random.sample(soft_skills, 3)
    
    # Choices
    first_choice = random.choice(program_choices[program_number])
    second_choice = random.choice([choice for choice in program_choices[program_number] if choice != first_choice])
    third_choice = random.choice([choice for choice in program_choices[program_number] if choice not in [first_choice, second_choice]])
    
    # Hackathon results
    hackathon_1 = random.choice(hackathon_results)
    hackathon_2 = random.choice(hackathon_results)

    # Risk and Placeability
    risk_status = random.choice(risk_statuses)
    placeability = random.choice(placeabilities)

    return [
        candidate_key, first_name, surname, qualification, previous_program, candidate_manager_score, candidate_manager_recommendation, 
        tech_score, tech_mentor_recommendation, random_it_skills[0], random_it_skills[1], random_it_skills[2], random_it_skills[3], 
        random_it_skills[4], random_soft_skills[0], random_soft_skills[1], random_soft_skills[2], first_choice, second_choice, 
        third_choice, hackathon_1, hackathon_2, risk_status, placeability
    ]

# Ensure unique names for all candidates
unique_candidate_data = []
for program_number in range(1, 9):
    for i in range(50):  # 50 candidates per program
        candidate_key = 1000 * program_number + i + 1
        first_name = first_names_unique.pop(0)
        surname = surnames_unique.pop(0)
        unique_candidate_data.append(generate_unique_candidate(candidate_key, program_number, first_name, surname))

# Create dataframes for each sheet
columns = ['Candidate_Key', 'First Name', 'Surname', 'Qualification', 'Previous Programme Completed', 
           'Overall Candidate Manager Score', 'Candidate Manager Recommendation', 'Overall Tech Score', 
           'Tech Mentor Recommendation', 'IT Skill 1', 'IT Skill 2', 'IT Skill 3', 'IT Skill 4', 'IT Skill 5', 
           'Soft Skill 1', 'Soft Skill 2', 'Soft Skill 3', '1st Choice', '2nd Choice', '3rd Choice', 
           'Hackathon 1', 'Hackathon 2', 'Risk Status', 'Placeability']

df_candidates_unique = pd.DataFrame(unique_candidate_data, columns=columns)

# Save to Excel
with pd.ExcelWriter('score_cards_NF.xlsx', engine='xlsxwriter') as writer:
    df_candidates_unique[['Candidate_Key', 'First Name', 'Surname', 'Qualification', 'Previous Programme Completed']].to_excel(writer, sheet_name='Candidates', index=False)
    df_candidates_unique[['Candidate_Key', 'Overall Candidate Manager Score', 'Candidate Manager Recommendation', 'Overall Tech Score', 'Tech Mentor Recommendation']].to_excel(writer, sheet_name='Performance', index=False)
    df_candidates_unique[['Candidate_Key', 'IT Skill 1', 'IT Skill 2', 'IT Skill 3', 'IT Skill 4', 'IT Skill 5']].to_excel(writer, sheet_name='IT Skills', index=False)
    df_candidates_unique[['Candidate_Key', 'Soft Skill 1', 'Soft Skill 2', 'Soft Skill 3']].to_excel(writer, sheet_name='Soft Skills', index=False)
    df_candidates_unique[['Candidate_Key', '1st Choice', '2nd Choice', '3rd Choice']].to_excel(writer, sheet_name='Program Choices', index=False)
    df_candidates_unique[['Candidate_Key', 'Hackathon 1', 'Hackathon 2']].to_excel(writer, sheet_name='Hackathon Participation', index=False)
    df_candidates_unique[['Candidate_Key', 'Risk Status']].to_excel(writer, sheet_name='Risk Status', index=False)
    df_candidates_unique[['Candidate_Key', 'Placeability']].to_excel(writer, sheet_name='Placeability', index=False)

print("Excel file generated successfully!")
