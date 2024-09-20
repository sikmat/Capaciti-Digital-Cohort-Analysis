import pandas as pd
import random

# Function to generate random scores
def generate_score(min_val, max_val):
    return random.randint(min_val, max_val)

# Function to generate IT and soft skills
it_skills = ['Python', 'Java', 'SQL', 'C++', 'JavaScript', 'HTML', 'CSS', 'Docker', 'Kubernetes', 'AWS', 'Azure', 'Linux', 'PowerBI', 'Excel']
soft_skills = ['Communication', 'Problem-solving', 'Teamwork', 'Leadership', 'Adaptability', 'Time management', 'Critical thinking', 'Creativity']

# Function to generate recommendations based on scores
def generate_recommendation(score, thresholds, recommendations):
    for i, threshold in enumerate(thresholds):
        if score <= threshold:
            return recommendations[i]
    return recommendations[-1]

# Defining recommendation thresholds and categories
tech_thresholds = [49, 64, 74]
tech_recommendations = ['Negative', 'Neutral', 'Positive', 'Outstanding']
manager_thresholds = [49, 64, 74]
manager_recommendations = ['Negative', 'Neutral', 'Positive', 'Outstanding']

# Dummy data for Hackathons
hackathon_outcomes = ['Winner', 'Second place', 'Third place', 'Not participated', 'Participated']

# Programme choices based on Graduate Programmes
choices = {
    1: ['Cyber', 'Data Analytics', 'SAP'],
    2: ['Software Development', 'DevOps', 'Cloud Computing'],
    3: ['AI/ML', 'Salesforce', 'Software Testing'],
    4: ['IT Support', 'IT Technician', 'Networking'],
    5: ['Data', 'Robotics', 'Web Development'],
    6: ['BI', 'Cloud Computing', 'SAP'],
    7: ['Business Analytics', 'Cloud Computing', 'Systems Analyst'],
    8: ['Web Development', 'UI/UX', 'React']
}

# Risk status and Placeability options
risk_status = ['Non-attendance', 'Behaviour', 'Tech Learning', 'Personal', 'Not at risk']
placeability = ['Ready', 'Will be ready', 'Potentially unplaceable']

# Function to generate a random hackathon outcome
def random_hackathon_result():
    return random.choices(hackathon_outcomes, weights=[5, 5, 5, 15, 20])[0], random.choices(hackathon_outcomes, weights=[5, 5, 5, 20, 15])[0]

# Generate unique candidate keys, names, surnames
candidate_keys = random.sample(range(1000, 9999), 50)
names = random.sample([f'Name{i}' for i in range(1, 100)], 50)
surnames = random.sample([f'Surname{i}' for i in range(1, 100)], 50)
qualifications = ['BSc Computer Science', 'BEng Software Engineering', 'BSc Information Technology', 'BSc Data Science']

# Function to create a graduate programme table
def create_programme_table(programme_num, num_candidates=50):
    data = []
    
    for i in range(num_candidates):
        # Random scores and recommendations
        tech_score = generate_score(30, 100)
        manager_score = generate_score(30, 100)
        
        tech_rec = generate_recommendation(tech_score, tech_thresholds, tech_recommendations)
        manager_rec = generate_recommendation(manager_score, manager_thresholds, manager_recommendations)
        
        # Random IT skills and soft skills
        it_skill_sample = random.sample(it_skills, 5)
        soft_skill_sample = random.sample(soft_skills, 3)
        
        # Hackathon results
        hackathon1, hackathon2 = random_hackathon_result()
        
        # Random qualifications, risk, placeability
        qualification = random.choice(qualifications)
        risk = random.choice(risk_status)
        placeability_status = random.choice(placeability)
        
        # Choices
        first_choice = random.choice(choices[programme_num])
        remaining_choices = [choice for choice in choices[programme_num] if choice != first_choice]
        second_choice, third_choice = random.sample(remaining_choices, 2)
        
        # Previous Programme completed (randomly assign relevance)
        previous_programme = random.choice([qualification if random.random() > 0.7 else None])
        
        # Append data
        data.append([
            candidate_keys[i], names[i], surnames[i], 
            manager_score, tech_score,
            it_skill_sample[0], it_skill_sample[1], it_skill_sample[2], it_skill_sample[3], it_skill_sample[4],
            soft_skill_sample[0], soft_skill_sample[1], soft_skill_sample[2],
            manager_rec, tech_rec,
            first_choice, second_choice, third_choice, qualification,
            hackathon1, hackathon2, previous_programme, risk, placeability_status
        ])
    
    # Create DataFrame
    columns = [
        'Candidate_Key', 'Names', 'Surname', 'Overall Candidate Manager score', 'Overall tech score',
        'IT Skill 1', 'IT Skill 2', 'IT Skill 3', 'IT Skill 4', 'IT Skill 5',
        'Soft Skill 1', 'Soft Skill 2', 'Soft Skill 3',
        'Candidate Manager Recommendation', 'Tech Mentor Recommendation',
        '1st Choice', '2nd Choice', '3rd Choice', 'Qualification',
        'Hackathon 1', 'Hackathon 2', 'Previous programme completed', 'Risk status', 'Placeability'
    ]
    
    return pd.DataFrame(data, columns=columns)

# Creating 8 tables for 8 IT graduate programmes
tables = {}
for i in range(1, 9):
    tables[f'score_cards{i}'] = create_programme_table(i)

# Saving the Excel file with 8 sheets
file_path = 'score_cards.xlsx'
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    for table_name, table_data in tables.items():
        table_data.to_excel(writer, sheet_name=table_name, index=False)

file_path
