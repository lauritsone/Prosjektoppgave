#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr 10 08:54:35 2025

@author: thomas
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime as dt
import time

#leser excelfilen og henter data:
data = pd.read_excel('support_uke_24.xlsx')


u_dag = arr = data["Ukedag"].to_numpy()
dictionary = {}
for item in u_dag:
    dictionary[item] = dictionary.get(item, 0) +1
print(dictionary)

print()

kl_slett = arr = data["Klokkeslett"].to_numpy()

print("Klokkeslett for innkommende samtaler: ")
print(kl_slett)

print()

print("Samtalenes varighet: ")

varighet = arr = data["Varighet"].to_numpy()
print(varighet)

print()

print("Tilfredshet:")
score = arr = data.dropna()["Tilfredshet"].to_numpy()
print(score)


print()
print()


#OBS! Denne plotter søylediagram i plots fanen
u_dag = arr = data["Ukedag"].to_numpy()
dictionary = {}
for item in u_dag:
    dictionary[item] = dictionary.get(item, 0) +1
Ukedag = np.array(list(dictionary.keys()))
Henvendelser = np.array(list(dictionary.values()))
plt.figure(1)
plt.bar(Ukedag, Henvendelser)
plt.show()


print()
print("Lengste og korteste samtale: ")


dictionary = {}
varighet = arr = data["Varighet"].to_numpy()
varig_høy = np.max(varighet)
varig_lav = np.min(varighet)
print(f"Korteste samtale varte i {varig_lav} minutter, og lengste samtale varte i {varig_høy} minutter")




print()
print()





#Henter ut og konverterer data fra kolonnen Varighet:
time_strings = np.array(data.Varighet)

#konverterer tid til sekunder:
def time_to_seconds(time_str):
    h, m, s = map(int, time_str.split(':'))
    return h * 3600 + m * 60 + s


time_floats = np.vectorize(time_to_seconds)(time_strings)

avg_time = np.mean(time_floats)

#Konverterer til tidsformat og printer resultat:
seconds = avg_time
time_obj = time.gmtime(seconds)
result_time = time.strftime("%H:%M:%S",time_obj)

print(f"Samtaletid for uke 24 har gjennomsnitt varighet på",result_time)

print()
print()






def kalkuler_nps(scores):
    promoters = scores[scores >= 9].count()
    passives = scores[(scores >= 7) & (scores <= 8)].count()
    detractors = scores[scores <= 6].count()
    total_responses = scores.count()

    if total_responses == 0:
        return 0, promoters, passives, detractors

    nps = ((promoters - detractors) / total_responses) * 100
    return round(nps, 2), promoters, passives, detractors

def main():
    

    # Leser excel og henter ut data til data
    data = pd.read_excel('support_uke_24.xlsx')


    scores = data['Tilfredshet']

    nps, promoters, passives, detractors = kalkuler_nps(scores)

    print(f"Promoters: {promoters}")
    print(f"Passives: {passives}")
    print(f"Detractors: {detractors}")
    print(f"Netto tilfredshet (NPS) er: {nps}")

if __name__ == "__main__":
    main()