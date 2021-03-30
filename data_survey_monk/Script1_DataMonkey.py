# -*- coding: utf-8 -*-
"""
Created on Wed Mar 24 11:41:46 2021

@author: bgourdon
"""

import pandas as pd
import os

pwd = os.getcwd()

dataset = pd.read_excel(pwd + "/Data - Survey Monkey Output.xlsx", sheet_name = "edited")
#print(dataset)

#créer une version test pour revenir en arrière
dataset_modified = dataset.copy()
dataset_modified

#regarder les headers
dataset_modified.columns
#print(dataset_modified.columns)

#pour drop une liste
#créer la liste
columns_to_drop = ['Start Date', 'End Date', 'Email Address',
       'First Name', 'Last Name', 'Custom Data 1']

#reprendre le dataset modifié et ajouté la colonne à droper
dataset_modified = dataset_modified.drop(columns = columns_to_drop)
#print(dataset_modified)

# melt methodo
# commencer sur la fonction melt par les id_vars, etc...
# créer des listes pour les id à pivoter et les valeurs à ne pas pivoter
# syntaxe start 0 à 8 et 8 à infini
id_vars = list(dataset_modified.columns)[0:8]
value_vars = list(dataset_modified.columns)[8:]
#test sur le premier
#print(id_vars)


#dataset_modified = pd.dataset_modified.melt()
dataset_melted = dataset_modified.melt(id_vars = id_vars,value_vars = value_vars, var_name="Question+subquestion",value_name="Answers")
#print(dataset_melted)

#utilisation du join 
# jointure sur la feuille excel Question
questions_import = pd.read_excel(pwd + "/Data - Survey Monkey Output.xlsx", sheet_name= "Question")
#print(questions_import)

#inplace = perform on datat right now on the data
questions = questions_import.copy()
questions.drop(columns=['RQuestion','RSubQuestion','Subquestion' ], inplace=True)


#on questions drop NAN
questions.dropna(inplace=True)
print(questions)

dataset_merged = pd.merge(left=dataset_melted, right=questions, how="left",left_on="Question+subquestion",right_on="Question+subquestion")
# connaitre si le join est bien executé et ressemble à l'original
# ici 17028 rows
print("Original Data",len(dataset_melted))
print("Merged Data", len(dataset_merged))
dataset_merged

#combien de personnes ont répondus aux questions ?
#Answer
#dataset_merged.groupby("Question")["Answers"].count().reset_index()

#combien de personnes DISTINCTES ont répondus aux questions ?
respondents = dataset_merged[dataset_merged["Answers"].notna()]
respondents = respondents.groupby("Question")["Respondent ID"].nunique().reset_index()
print("respondents", len(respondents))
#créer new dataframe pour merged les deux dataset_merged pour cela on va rename Respon ID
respondents.rename(columns={"Respondent ID":"Respondents"}, inplace=True)

#second merged
dataset_merged_two = pd.merge(left=dataset_merged, right=respondents, how="left",left_on="Question",right_on="Question")
# connaitre si le join est bien executé et ressemble à l'original
# ici 17028 rows
print("Original Data",len(dataset_merged))
print("Merged Data 2", len(dataset_merged_two))
dataset_merged_two

#same_answer per customer
same_answer = dataset_merged #[dataset_merged["Answers"].notna()]
same_answer = same_answer.groupby(["Question+subquestion","Answers"])["Respondent ID"].nunique().reset_index()
#créer new dataframe pour merged les deux dataset_merged pour cela on va rename Respon ID
same_answer.rename(columns={"Respondent ID":"Same Answer"}, inplace=True)
same_answer

#3rd merged
dataset_merged_three = pd.merge(left=dataset_merged_two, right=same_answer, how="left",left_on=["Question+subquestion","Answers"],right_on=["Question+subquestion","Answers"])
print("Original Data",len(dataset_merged_two))
print("Merged Data 3", len(dataset_merged_three))


#extraire le fichier
#rename columns
#print(dataset_merged_three.columns)
output = dataset_merged_three.copy()
output.rename(columns={ "Identify which division you work in.-Response":"Division Primary",
       "Identify which division you work in.-Other (please specify)":"Division Secondary",
       "Which of the following best describes your position level?-Response":"Position",
       "Which generation are you apart of?-Response":"Generation",
       "Please select the gender in which you identify.-Response":"Gender",
       "Which duration range best aligns with your tenure at your company?-Response":"Tenure",
       "Which of the following best describes your employment type?-Response":"Employment Type"},inplace=True
)
print(output)

#extraire vers le fichier directory
output.to_excel(pwd + "//Final_Output.xls",index=False)

