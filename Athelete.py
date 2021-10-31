import plotly.express as px
import pandas as pd
import numpy as np  # linear algebra
import pandas as pd  # data processing, CSV file I/O (e.g. pd.read_csv)
import openpyxl

import os
for dirname, _, filenames in os.walk('C:\\Users\\Jaswanth Jeethu\\OneDrive\\Documents\\Project\\Data'):
    for filename in filenames:
        print(os.path.join(dirname, filename))

entries_gender = pd.read_excel(
    'C:\\Users\\Jaswanth Jeethu\\OneDrive\\Documents\\Project\\Data\\EntriesGender.xlsx')
teams = pd.read_excel(
    'C:\\Users\\Jaswanth Jeethu\\OneDrive\\Documents\\Project\\Data\\Teams.xlsx')
athletes = pd.read_excel(
    'C:\\Users\\Jaswanth Jeethu\\OneDrive\\Documents\\Project\\Data\\Athletes.xlsx')
coaches = pd.read_excel(
    'C:\\Users\\Jaswanth Jeethu\\OneDrive\\Documents\\Project\\Data\\Coaches.xlsx')
medals = pd.read_excel(
    'C:\\Users\\Jaswanth Jeethu\\OneDrive\\Documents\\Project\\Data\\Medals.xlsx')

athletes.head()
coaches.head()
teams.head()
medals.head()
entries_gender.head()

print("The number of athletes that participated in Olympics 2020 is " +
      str(athletes.shape[0]))
print("The number of sports categories in Olympics 2020 is " +
      str(athletes.Discipline.nunique()))
print("The number of countries that participated in Olympics 2020 is " +
      str(athletes.NOC.nunique()))
print("The number of coaches present in Olympics 2020 is " +
      str(coaches.shape[0]))

# 1
print("**Distribution of Athletes across different Sports**")
num_athletes_sportwise = athletes.pivot_table(
    index='Discipline', values='Name', aggfunc=pd.Series.nunique)
num_athletes_sportwise.reset_index(inplace=True)
num_athletes_sportwise.rename(
    columns={'Name': 'Count of Athletes'}, inplace=True)
num_athletes_sportwise.sort_values(
    by=['Count of Athletes'], ascending=False, inplace=True)
fig = px.bar(num_athletes_sportwise, x='Discipline', y='Count of Athletes')
fig.show()

# 2
print("Distribution of Athletes across Countries")
num_athletes_countrywise = athletes.pivot_table(
    index='NOC', values='Name', aggfunc=pd.Series.nunique)
num_athletes_countrywise.reset_index(inplace=True)
num_athletes_countrywise.rename(
    columns={'Name': 'Count of Athletes'}, inplace=True)
num_athletes_countrywise.sort_values(
    by=['Count of Athletes'], ascending=False, inplace=True)
fig = px.bar(num_athletes_countrywise, x='NOC', y='Count of Athletes')
fig.show()

# 3
print("Distribution of Coaches across different Sport Categories")
num_coaches_sportwise = coaches.pivot_table(
    index='Discipline', values='Name', aggfunc=pd.Series.nunique)
num_coaches_sportwise.reset_index(inplace=True)
num_coaches_sportwise.rename(
    columns={'Name': 'Count of Coaches'}, inplace=True)
num_coaches_sportwise.sort_values(
    by=['Count of Coaches'], ascending=False, inplace=True)
fig = px.bar(num_coaches_sportwise, x='Discipline',
             y='Count of Coaches', width=800, height=400)
fig.show()

# 4
print("Distribution of Coaches across different Countries")
num_coaches_countrywise = coaches.pivot_table(
    index='NOC', values='Name', aggfunc=pd.Series.nunique)
num_coaches_countrywise.reset_index(inplace=True)
num_coaches_countrywise.rename(
    columns={'Name': 'Count of Coaches'}, inplace=True)
num_coaches_countrywise.sort_values(
    by=['Count of Coaches'], ascending=False, inplace=True)
fig = px.bar(num_coaches_countrywise, x='NOC', y='Count of Coaches')
fig.show()

# To analyse further, I created a variable sport_country signifying the team on the basis of sports and country. There are 400 sport teams in total. These teams include men and women participants.
teams['sport_country'] = teams['Discipline']+"_"+teams['NOC']

print("Number of events participated by a sports team")
df = teams['sport_country'].value_counts().to_frame()
df.rename(columns={"sport_country": "Number of events"}, inplace=True)
fig = px.histogram(df, x="Number of events", width=600, height=400,
                   title="Number of events participated by sports team of a country")
fig.update_layout(bargap=0.2, yaxis_title="Count of Teams")
fig.show()

# Does every Athlete belong to a team?

# I calculated the sport_country variable for every athlete. If the sport_country variable for an athlete is present in the teams table then the athlete belongs to a team.
sport_country = teams.sport_country.unique()
athletes['sport_country'] = athletes['Discipline']+"_"+athletes['NOC']
athletes['in_a_team'] = [
    'yes' if x in sport_country else 'no' for x in athletes['sport_country']]

df = athletes.in_a_team.value_counts().to_frame()
df.reset_index(inplace=True)
fig = px.pie(df, values='in_a_team', names='index',
             title='Athlete belongs to a team?', width=400, height=400)
fig.show()

# The sport_country variable was also calculated for coaches to see which teams have coaches.
coaches['sport_country'] = coaches['Discipline']+"_"+coaches['NOC']

coach_sport_country = coaches.sport_country.unique()
teams['has_coach'] = [
    True if x in coach_sport_country else False for x in teams['sport_country']]
teams.head()

df = teams.has_coach.value_counts().to_frame()
df.reset_index(inplace=True)
fig = px.pie(df, values='has_coach', names='index',
             title='Team has a coach?', width=400, height=400)
fig.show()

# Distribution of Teams across Sports
df = pd.get_dummies(teams, columns=['has_coach'])
df = df.pivot_table(index='Discipline', values=[
                    'has_coach_True', 'has_coach_False'], aggfunc=pd.Series.sum)
df.reset_index(inplace=True)
fig = px.bar(df, x='Discipline', y=['has_coach_False', 'has_coach_True'], color_discrete_sequence=[
             '#636EFA', '#B6E880'], title="Count of Teams")
fig.update_layout(
    xaxis_title="Sports",
    yaxis_title="Count of Teams")
fig.show()

# To classify the event as men, women or mixed, a variable event_type was created.


def event_type(string):
    if ("women" in string.lower()):
        return "women"
    elif ("men" in string.lower()):
        return "men"
    else:
        return "mixed"


teams['event_type'] = teams.Event.apply(lambda x: event_type(x))
teams.head()

# Distribution of Event Types
df = teams.event_type.value_counts().to_frame()
df.reset_index(inplace=True)
df.head()
fig = px.pie(df, values='event_type', names='index', width=400,
             height=400, title="Distribution of Event Types")
fig.show()

# Distribution of event types across countries
df = pd.get_dummies(teams, columns=['event_type'])
df.reset_index(inplace=True)
df = df.pivot_table(index='NOC', values=[
                    'event_type_men', 'event_type_mixed', 'event_type_women'], aggfunc=pd.Series.sum)
df.reset_index(inplace=True)
df['total'] = df['event_type_men'] + \
    df['event_type_mixed']+df['event_type_women']
df.sort_values(by=['total'], ascending=False, inplace=True)
fig = px.bar(df, x='NOC', y=['event_type_men', 'event_type_mixed', 'event_type_women'], color_discrete_sequence=[
             '#636EFA', '#B6E880', '#FF6692'], title="Distribution of Event Types Across Countries")
fig.update_layout(
    xaxis_title="NOC",
    yaxis_title="Number of Events")
fig.show()


# Distribution of event types across sports
df = pd.get_dummies(teams, columns=['event_type'])
df.reset_index(inplace=True)
df = df.pivot_table(index='Discipline', values=[
                    'event_type_men', 'event_type_mixed', 'event_type_women'], aggfunc=pd.Series.sum)
df.reset_index(inplace=True)
df['total'] = df['event_type_men'] + \
    df['event_type_mixed']+df['event_type_women']
df.sort_values(by=['total'], ascending=False, inplace=True)
fig = px.bar(df, x='Discipline', y=['event_type_men', 'event_type_mixed', 'event_type_women'], color_discrete_sequence=[
             '#636EFA', '#B6E880', '#FF6692'], title="Distribution of Event Types Across Sports")
fig.update_layout(
    xaxis_title="Sport",
    yaxis_title="Number of Events")
fig.show()

# Distribution of Medal Across Countries
medals.sort_values(by='Total', ascending=False, inplace=True)
fig = px.bar(medals, x='Team/NOC', y=['Gold', 'Silver', 'Bronze'], color_discrete_sequence=[
             'gold', 'silver', 'brown'], title="Distribution of Medals Across Countries")
fig.update_layout(yaxis_title='Count of Medals')
fig.show()


# Sport Participation by Gender
entries_gender.sort_values(by='Total', ascending=False, inplace=True)
fig = px.bar(entries_gender, x='Discipline', y=['Male', 'Female'], color_discrete_sequence=[
             '#636EFA', '#FF6692'], title="Participation by Gender")
fig.update_layout(yaxis_title='Number of Athletes')
fig.show()
