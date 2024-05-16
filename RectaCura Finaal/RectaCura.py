#!/usr/bin/env python3.10
# coding: utf-8

# In[1]:
#MET DATA WEG TE SCHRIJVEN
import re
from flask import Flask
from pywebio.platform.flask import webio_view
from pywebio.input import *
from pywebio.output import *
from pywebio.pin import *
from pywebio import session
import pandas as pd
from numpy import *

app = Flask(__name__)
app.secret_key = "secretkey11111"

def main():

    class Questionnaire:
        def __init__(self, excel_file_path, vragen_sheet_path, categorien_sheet_path, adviessheet_path, output_excel_path):
        # Lees de Excel-sheets
            self.vragen_df = pd.read_excel(excel_file_path, sheet_name=vragen_sheet_path)
            self.categorien_df = pd.read_excel(excel_file_path, sheet_name=categorien_sheet_path)
            self.advice_df = pd.read_excel(excel_file_path, sheet_name=adviessheet_path)
            self.output_excel_path = output_excel_path
            self.user_responses = {}
            self.set_percentages = {}

        def start(self):
            self.present_set_1_general()
            self.process_user_responses()
            self.present_selected_sets()
            self.calculate_percentages()
            self.print_advice()

        def present_set_1_general(self):
            # Presenteer Set 1 Algemeen
            set_name = "Set 1 Algemeen"
            set_vragen = self.vragen_df[self.vragen_df["Setnaam"] == set_name]

            # Stel de eerste acht vragen van Set 1 Algemeen voor
            for _, row in set_vragen.iterrows():
                vraag = row["Vraagtekst"]
                antwoordopties = row["Antwoordopties"]
                if pd.isnull(antwoordopties):
                    antwoordopties = row["Antwoordopties"]
                else:
                    antwoordopties = antwoordopties.split(';')
                soort = row['Soort vraag']

                # Vraag 9 is de laatste vraag, dus sla het over in de eerste lus
                if vraag == "Van welke klachten heb je zoal last?":
                    continue

                # Print vraag en Vraag de gebruiker om een antwoord
                if soort == 'slider':
                    keuze = slider(vraag, min_value=1, max_value=10)
                if soort == 'input':
                    keuze = input(vraag)
                if soort == 'radio':
                    keuze = radio(vraag, [i for i in antwoordopties] )
                if soort == 'checkbox':
                    keuze = checkbox(vraag, [i for i in antwoordopties])
                # Controleer of de keuze binnen het bereik van antwoordopties valt
                # Sla het antwoord op
                self.user_responses[f"{set_name} - {vraag}"] = keuze

            # Stel vraag 9 als laatste
            vraag_9_row = set_vragen[set_vragen["Vraagtekst"] == "Van welke klachten heb je zoal last?"].iloc[0]
            vraag_9 = vraag_9_row["Vraagtekst"]
            antwoordopties_9 = vraag_9_row["Antwoordopties"].split(";")

            # Stel de vraag en laat de gebruiker al zijn klachten aanduiden obv een checkbox
            chosen_categories = checkbox(vraag_9, [i for i in antwoordopties_9] )
            self.user_responses["chosen_categories"] = chosen_categories

            # Processbar om aan te tonen dat nieuwe vragenset geladen wordt obv voorgaande antwoorden
            put_processbar('bar')
            for i in range(1, 11):
                set_processbar('bar', i / 10)
            put_markdown("We verwerken je antwoorden en vragen dynamisch verder...")

        def process_user_responses(self):
            # Haal de gekozen categorieën op
            chosen_categories = self.user_responses["chosen_categories"]

            #Filter de categorieën op basis van de gekozen categorieën
            filtered_categorien = self.categorien_df[self.categorien_df["Categorie"].isin(chosen_categories)]

            # Zoek de hoogste urgentiescore van de gekozen categorieën
            max_urgency_score = filtered_categorien["Urgentiescore"].max()

            # Selecteer sets op basis van de urgentiescore
            selected_sets = set()
            for _, row in filtered_categorien.iterrows():
                urgency_score = row["Urgentiescore"]
                set_name = row["Vragenset"]

                # Voeg sets toe als de urgentiescore binnen het bereik valt
                if abs(urgency_score - max_urgency_score) <= 2:
                    selected_sets.add(set_name)

            # Sla de geselecteerde sets op
            self.user_responses["selected_sets"] = list(selected_sets)

        def present_selected_sets(self):
            # Haal de geselecteerde sets op
            selected_sets = self.user_responses["selected_sets"]

            # Presenteer de vragen van de geselecteerde sets aan de gebruiker
            for set_name in selected_sets:
                # Filter de vragen voor de opgegeven set
                set_vragen = self.vragen_df[self.vragen_df["Setnaam"] == set_name]

                # Stel de vragen voor
                for _, row in set_vragen.iterrows():
                    vraag = row["Vraagtekst"]
                    antwoordopties = row["Antwoordopties"].split(";")
                    soort = row['Soort vraag']
                    # Stel de vraag en vraag de gebruiker om een antwoord
                    # Print vraag en Vraag de gebruiker om een antwoord
                    if soort == 'slider':
                        keuze = slider(vraag, min_value=1, max_value=10)
                    if soort == 'input':
                        keuze = input(vraag)
                    if soort == 'radio':
                        keuze = radio(vraag, [i for i in antwoordopties] )
                    if soort == 'checkbox':
                        keuze = checkbox(vraag, [i for i in antwoordopties])
                    # Sla het antwoord op
                    self.user_responses[f"{set_name} - {vraag}"] = keuze

                # Controleer of de gebruiker naar Set 2.8 moet worden doorverwezen
                if set_name == "Set 1.8: Depressie/Donkere gedachten" and vraag == "Heb je soms donkere gedachten gehad over zelfmoord?":
                    if keuze == 'Ja':  # Ja antwoord (keuze 1)
                        # Voeg Set 2.8 toe aan de geselecteerde sets
                        if "Set 2.8: Zelfmoord/Zelfbeschadiging" not in selected_sets:
                            selected_sets.append("Set 2.8: Zelfmoord/Zelfbeschadiging")
                            self.user_responses["chosen_categories"].append('Zelfbeschadiging')
                            # Breek de loop af na Set 1.8 omdat we mogelijk naar Set 2.8 moeten gaan
                            break
                if set_name == "Set 1.1: Trauma en misbruik" and vraag == "Heb je ooit gedacht aan zelfbeschadiging of zelfmoord als gevolg van het misbruik of trauma?":
                    if keuze == 'Ja':
                        # Voeg Set 2.8 toe aan de geselecteerde sets
                        if "Set 2.8: Zelfmoord/Zelfbeschadiging" not in selected_sets:
                            selected_sets.append("Set 2.8: Zelfmoord/Zelfbeschadiging")
                            self.user_responses["chosen_categories"].append('Zelfbeschadiging')
                            # Breek de loop af na Set 1.8 omdat we mogelijk naar Set 2.8 moeten gaan
                            break

            # Presenteer Set 2.8 als het antwoord "Ja" was op een vraag over zelfmoordgedachten
            if "Set 2.8: Zelfmoord/Zelfbeschadiging" in selected_sets:
                # Filter de vragen voor Set 2.8
                set_vragen = self.vragen_df[self.vragen_df["Setnaam"] == "Set 2.8: Zelfmoord/Zelfbeschadiging"]

                # Stel de vragen voor
                for _, row in set_vragen.iterrows():
                    vraag = row["Vraagtekst"]
                    antwoordopties = row["Antwoordopties"].split(";")
                    keuze = row['Soort vraag']
                    # Stel vraag en laat de gebruiker de vraag beantwoorden
                    # Print vraag en Vraag de gebruiker om een antwoord
                    if soort == 'slider':
                        keuze = slider(vraag, min_value=1, max_value=10)
                    if soort == 'input':
                        keuze = input(vraag)
                    if soort == 'radio':
                        keuze = radio(vraag, [i for i in antwoordopties] )
                    if soort == 'checkbox':
                        keuze = checkbox(vraag, [i for i in antwoordopties])
                # Controleer of de keuze binnen het bereik van antwoordopties valt
                    # Sla het antwoord op
                    self.user_responses[f"Set 2.8: Zelfmoord/Zelfbeschadiging - {vraag}"] = keuze


        def calculate_percentages(self):
            # Bereken de percentages voor elke categorie
            for category in self.user_responses.get("chosen_categories", []):

                category_sets = self.categorien_df[self.categorien_df["Categorie"] == category]["Vragenset"].iloc[0].split(",")
                total_score = 0
                max_score = 0

                # Bereken de totale score van de gebruiker en de hoogst haalbare score voor de categorie
                for set_name in category_sets:
                    set_vragen = self.vragen_df[self.vragen_df["Setnaam"] == set_name]
                    for _, row in set_vragen.iterrows():
                        vraag = row["Vraagtekst"]
                        antwoordopties = row["Antwoordopties"].split(";")

                        # Haal het antwoord van de gebruiker op
                        user_answer = self.user_responses.get(f"{set_name} - {vraag}", None)

                        # Zoek de index van het antwoord van de gebruiker
                        if user_answer in antwoordopties:
                            user_index = antwoordopties.index(user_answer)
                            total_score += user_index+1 # De index + 1 is de score voor dat antwoord

                        # De hoogst haalbare score is de lengte van de antwoordopties
                        max_score += len(antwoordopties)

                # Bereken het percentage
                if max_score > 0:
                    percentage = (total_score / max_score) * 100
                    self.set_percentages[category] = percentage
                    print(percentage)

        def print_advice(self):
            # Haal de gekozen categorieën op
            chosen_categories = self.user_responses.get("chosen_categories", [])

            # Verkrijg de volgorde van categorieën uit Vraag 9 van Set 1 Algemeen
            question_9_categories = self.categorien_df["Categorie"].tolist()

            # Sorteer de gekozen categorieën op basis van hun volgorde in Vraag 9 van Set 1 Algemeen
            sorted_chosen_categories = sorted(chosen_categories,key=lambda category: question_9_categories.index(category))

            # Controleer de categorieën van hoog naar laag op basis van de percentages
            reverse=sorted_chosen_categories[::-1]
            for category in reverse:
                # Controleer of de categorie een percentage heeft
                #if category in self.set_percentages:
                percentage = self.set_percentages[category]
                # Als het percentage 40% of hoger is, geef dan advies voor deze categorie
                if percentage >= 40:
                    self.print_advice_for_category(category)
                    return  # We zijn klaar, advies gegeven voor deze categorie

                    # Als het percentage lager is dan 40%, ga door naar de volgende categorie
                else:
                    continue

            # Als geen enkele categorie een percentage van 40% of hoger had, geef advies voor de hoogste categorie (laatste in de gesorteerde lijst)
            highest_category = sorted_chosen_categories[-1]
            self.print_advice_for_category(highest_category)

        def print_advice_for_category(self, category):
            # Zoek het berekende percentage voor deze categorie
            percentage = self.set_percentages[category]
            # Zoek het advies voor de gekozen categorie
            advice_row = self.advice_df[self.advice_df["Categorie"] == category]
            for _, row in advice_row.iterrows():
                # Split het interval in de kolom "Adviespercentage" en converteer naar integers
                interval = row["Adviespercentage"].split("-")
                min_percentage = int(interval[0])
                max_percentage = int(interval[1])

                # Controleer of het berekende percentage binnen dit interval valt
                if min_percentage <= percentage <= max_percentage:
                    # Print het advies
                    put_text(f"Advies: {row['Advies']}")
                    # We maken de kolommen evenlang
                    # Print zelfhulpmodule als aanwezig
                    zelfhulpmodule = row.get("Zelfhulpmodule", "")
                    link_regex = r'(https?://\S+)'
                    links = re.findall(link_regex, zelfhulpmodule)
                    if zelfhulpmodule:
                        put_text("\nAanbevolen modules:")
                        for link in links:
                            put_link(link, url=link, new_window=True), put_text('\t')
                    # Print noodnummers als aanwezig
                    noodnummers = row.get("Noodnummers", "")
                    links2 = re.findall(link_regex, noodnummers)
                    if noodnummers:
                        put_text('In geval van nood, aarzel niet om contact op te nemen met een van deze noodnummers:')
                        for link in links2:
                            put_link(link, url=link, new_window=True), put_text('\t')
                    #Indien advies = zelfhulpmodules, wordt er niet gevraagd direct een afspraak te maken
                    if row['Advies'] == 'Volgens RectaCura kan je momenteel het beste verdergeholpen worden met onze online zelfhulpmodules. Deze zijn ontworpen om je te ondersteunen en helpen je het hoofd te bieden tegen de problemen die voor jou van toepassing zijn. Als je het gevoel hebt dat je extra professionele begeleiding nodig hebt, moedigen we je aan om een afspraak te maken bij een lokale hulpverlener.  ':
                        return

                    antwoord = input('Indien je akkoord bent met onze privacy verklaring en wenst een afspraak te maken, gelieve dan hieronder je e-mailadres in te geven (hetzelfde als hetgeen dat gebruikt wordt voor je RectaCura-account).')
                    # Regex-patroon voor het controleren van een geldig e-mailadres
                    email_regex = r'^[\w\.-]+@[\w\.-]+\.[\w]+$'

                    # Controleren of het antwoord overeenkomt met het e-mailadrespatroon
                    if re.match(email_regex, antwoord):
                        stat = 'ok'
                    if stat == 'ok':
                        put_text('Bedankt voor het invullen van de oriëntatiemodule. Je kan nu een afspraak maken bij een hulpverlener in jouw buurt op: https://sofiebauwens8.wixsite.com/rectacura/book-online')
                        #for question, answer in self.user_responses.items():
                        data['Vraag en antwoord'].append(self.user_responses.items())
                        data['Advies'].append(row['Advies'])
                        data['Percentages'].append(self.set_percentages)
                        data['Emailadres'].append(antwoord)

                        # Schrijf naar Excel
                        with open(self.output_excel_path, 'a', encoding='utf-8') as f:
                            for i, question in enumerate(data['Vraag en antwoord']):
                                f.write(f"{question}\t{data['Vraag en antwoord'][i]}\t{data['Advies'][i]}\t{data['Percentages'][i]}\t{data['Emailadres']}\n")

                    return  # Stop de loop zodra het juiste advies is gevonden

            # Als er geen advies gevonden is, geef een melding
            put_text("Geen passend advies gevonden voor dit percentage.")


        def get_category_from_set_name(self, set_name):
            # Deze functie is een hulpfunctie om de categorie te vinden op basis van de setnaam
            # Neem aan dat de setnaam in het formaat "Set X.Y: Categorie Naam" is.
            # Splits de setnaam op basis van de dubbele punt (:) en neem het tweede deel als categorie.
            category = set_name.split(":")[1].strip()
            return category

    # Pad naar de Excel-sheets met vragen en categorieën
    excel_file_path = "mysite/data/Projectvak-data.xlsx"
    output_excel_path = "mysite/data/Output-data.xlsx"
    vragen_sheet_path = "Vragen"
    categorien_sheet_path = "Vraag 9 - Set 1 Algemeen"
    adviessheet_path = "Adviessheet"
    data = {
        'Vraag en antwoord': [],
        'Advies': [],
        'Percentages': [], 'Emailadres': []}
    # Maak een nieuwe vragenlijst
    questionnaire = Questionnaire(excel_file_path, vragen_sheet_path, categorien_sheet_path, adviessheet_path, output_excel_path)

    # Start de vragenlijst
    questionnaire.start()

app.add_url_rule('/', 'webio_view', webio_view(main), methods=['GET', 'POST'])

if __name__ == "__main__":
    #start_server(main,debug=True)
    #run the app and enable debugging
    app.run(debug=False)