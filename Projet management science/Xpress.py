import pandas as pd
from pulp import *

# Définition des deux sources de données
excel1 = "/Users/thomasdeclercq/Desktop/INGE12BA/Q2/Management science /Projet /Xpress/Copy of Origine_Destination.xlsx"
# excel2 = "/Users/thomasdeclercq/Downloads/Data_excel_xpress2023/g60n.xlsx"
excel2 = "/Users/thomasdeclercq/Desktop/INGE12BA/Q2/Management science /Projet /Xpress/Fichier final Xpress/g40nn(données).xlsx"
data_1 = pd.ExcelFile(excel1)
data_2 = pd.ExcelFile(excel2)
data_11 = pd.read_excel(data_1, 'Tableau_5', index_col=0) # Tableau 1
data_12 = pd.read_excel(data_1, 'Tableau_6', index_col=0) # Tableau 2
data_demand = pd.read_excel(data_2, 'Tableau_1',nrows=12, index_col=0)# Tableau 3
data_revenue = pd.read_excel(data_2, 'Tableau_2',nrows=12, index_col=0)# Tableau 4
data_contraintes = pd.read_excel(data_2, 'Tableau_3', nrows=7, index_col=0)# Tableau 5
data_capa = pd.read_excel(data_2, 'Tableau_4', nrows=12, index_col=0)# Tableau 6
# data_capa_mod = pd.read_excel(data_2, 'Tableau_5', nrows=12, index_col=0)# Tableau 6


lst_depart=[]
for i in range(len(data_11)):
    lst_i=[]
    for j in range(len(data_11.columns)):
        lst_i.append(data_11.iloc[i,j])
    lst_depart.append(lst_i)


lst_arrivee = []
for i in range(len(data_12)):
    lst_i=[]
    for j in range(len(data_12.columns)):
        lst_i.append(data_12.iloc[i,j])
    lst_arrivee.append(lst_i)


# Création d'une liste avec toutes les villes
villes = list(data_12.columns)
# Liste avec les trajets
traj = list(data_11.index)
# Liste des mois
mois = list(data_demand.columns)
# Liste des coûts
dem = []
for i in range(len(data_demand)):
    lst_i=[]
    for j in range(len(data_demand.columns)):
        lst_i.append(data_demand.iloc[i,j])
    dem.append(lst_i)
# Dictionnaire des demandes
demand = makeDict([traj, mois], dem, 0)

rev = []
for i in range(len(data_revenue)):
    lst_i=[]
    for j in range(len(data_revenue.columns)):
        lst_i.append(data_revenue.iloc[i,j])
    rev.append(lst_i)
# Dictionnaire des revenus
revenue = makeDict([traj, mois], rev, 0)

# Création d'un dictionnaire utile à la question 3

# Utile pour l'analyse de sensibilité
# cap_m = []
# for i in range(len(data_capa_mod)):
#     lst_i=[]
#     for j in range(len(data_capa_mod.columns)):
#         lst_i.append(data_capa_mod.iloc[i,j])
#     cap_m.append(lst_i)
# # Dictionnaire des revenus
# capa_mod = makeDict([traj, mois], cap_m, 0)

# Dictionnaire du nombre de conteneurs initiaux
cont_dep = dict(zip(villes, data_contraintes["Nombre conteneurs mai 2023"]))
# Dictionnaire des coûts de chargement/déchargement
cout_char = dict(zip(villes, data_contraintes["Coût chargement/déchargement"]))
# Dictionnaire du tarif leasing
c_leasing = dict(zip(villes, data_contraintes["Tarif leasing annuel conteneur"]))
# Dictionnaire du coût de chaque trajet
cout_traj = {}
for i in traj:
    for j in cout_char:
        for k in cout_char:
            if j != k:
                if j in i and k in i:
                    cout_traj[i]=cout_char[j]+cout_char[k]
c_stock = dict(zip(villes, data_contraintes["cout entrepot unitaire pour un mois"]))
stock_max = dict(zip(villes, data_contraintes["Capacité disponible pour stock"]))
capa = dict(zip(traj, data_capa["Capacité mensuelle"]))


prob = LpProblem("Xpress", LpMaximize)

# Je crée tous les trajets en supprimant ceux de la même ville à la même ville
# trajets = [(a,d)for a in villes for d in villes if est_couple(a,d,villes,lst_depart,lst_arrivee)]

villes_dep = []
for i in range(len(lst_depart)):
    for j in range(len(lst_depart[i])):
        if lst_depart[i][j] == 1:
            villes_dep.append(villes[j])

villes_ar = []
for i in range(len(lst_arrivee)):
    for j in range(len(lst_arrivee[i])):
        if lst_arrivee[i][j] == 1:
            villes_ar.append(villes[j])

routes = [(villes_dep[i], villes_ar[i]) for i in range(11)]

# Création d'une structure d'un dictionnaire avec la ville de départ et d'arrivée associée au trajet
traj_dep_arv = {}
for i in traj:
    for v in villes:
        if v in i:
            if i not in traj_dep_arv:
                traj_dep_arv[i]=[v]
            elif i in traj_dep_arv:
                traj_dep_arv[i].append(v)
for k in traj_dep_arv:
    for i in range(3):
        if k[i] != traj_dep_arv[k][0][i] and k[i] != traj_dep_arv[k][0][i]:
            temp = traj_dep_arv[k][0]
            traj_dep_arv[k][0] = traj_dep_arv[k][1]
            traj_dep_arv[k][1] = temp

# Creation d'un dictionnaire avec pour chaque ville toutes les liaisons qui arrivent à cette ville
arv = {}
for j in villes:
    for k in traj_dep_arv:
        if j == traj_dep_arv[k][1]:
            if j not in arv:
                arv[j]=[k]
            else:
                arv[j].append(k)
# Création d'un dictionnaire avec pour chaque ville toutes les liaisons qui partent de cette ville
dep = {}
for j in villes:
    for k in traj_dep_arv:
        if j == traj_dep_arv[k][0]:
            if j not in dep:
                dep[j]=[k]
            else:
                dep[j].append(k)

# Initialisation des variables
var_p = LpVariable.dicts("Plein", (traj,mois), 0 ,None,LpInteger)
var_v = LpVariable.dicts("Vide", (traj, mois), 0, None, LpInteger)
var_leasing = LpVariable.dicts("Leasing_Mai", villes, 0, None, LpInteger)
var_stock = LpVariable.dicts("Stock", (villes, mois), 0, None, LpInteger)


# Fonction objectif
prob += lpSum([var_p[i][m]*revenue[i][m] for i in traj for m in mois]) \
        - lpSum([(var_p[i][m]+var_v[i][m])*cout_traj[i]for i in traj for m in mois])\
        - lpSum([var_stock[i][m]*c_stock[i] for i in villes for m in mois]) \
        - lpSum([var_leasing[i]*c_leasing[i] for i in villes]) \
        # + lpSum([var_v[traj[2]][m]*cout_traj[traj[2]] for m in mois])\


# Contrainte de satisfaction de la demande
for i in traj:
    for m in mois:
        prob += var_p[i][m] >= 0.6*demand[i][m]
        prob += var_p[i][m] <= demand[i][m]

# Contrainte test pour la question 3
# for m in mois:
#     prob += var_p[traj[2]][m]<= 350000

# Contrainte de capacité mensuelle maximale de chaque ligne
for i in traj:
    for m in mois:
        prob += var_p[i][m]+var_v[i][m] <= capa[i]

# Version alternative test question 3
# for i in traj:
#     for m in mois:
#         prob += var_p[i][m]+var_v[i][m] <= capa_mod[i][m]

# Contrainte de capacité de stockage maximale de chaque port
for i in villes:
    for m in mois:
        prob += var_stock[i][m] <= stock_max[i]

# Contrainte de balance
for v in villes:
    for m in range(1,len(mois)):
        k = m-1
        prob += var_stock[v][mois[k]] + lpSum([var_p[t][mois[k]] for t in arv[v]]) + lpSum([var_v[t][mois[k]] for t in arv[v]])\
              == var_stock[v][mois[m]] + lpSum([var_p[t][mois[m]] for t in dep[v]])\
             + lpSum([var_v[t][mois[m]] for t in dep[v]])

# Il me reste à l'initialiser pour la relation entre t = 0 et t = 1


for v in villes:
    prob += cont_dep[v]+var_leasing[v] == var_stock[v][mois[0]] + lpSum([var_p[t][mois[0]]for t in dep[v]])\
         + lpSum([var_v[t][mois[0]]for t in dep[v]])


prob.solve()

print("Status :", LpStatus[prob.status])

for v in prob.variables():
    print(v.name, "=", v.varValue)

print(value(prob.objective))


dic_plein ={}
for m in mois:
    for t in traj:
        if m not in dic_plein:
            dic_plein[m] = [var_p[t][m].varValue]
        else:
            dic_plein[m].append(var_p[t][m].varValue)

Allocation_plein = pd.DataFrame(dic_plein, index = traj)


dic_vide = {}
for m in mois:
    for t in traj:
        if m not in dic_vide:
            dic_vide[m] = [var_v[t][m].varValue]
        else:
            dic_vide[m].append(var_v[t][m].varValue)

Allocation_vide = pd.DataFrame(dic_vide, index = traj)

dic_stock = {}
for m in mois:
    for v in villes:
        if m not in dic_stock:
            dic_stock[m]=[var_stock[v][m].varValue]
        else:
            dic_stock[m].append(var_stock[v][m].varValue)

Allocation_stock = pd.DataFrame(dic_stock, index = villes)

dic_leasing = {"Mai":[]}

for v in villes:
        dic_leasing["Mai"].append(var_leasing[v].varValue)

Allocation_leasing = pd.DataFrame(dic_leasing, index = villes)

with pd.ExcelWriter("Plan Xpress.xlsx") as writer :
    Allocation_vide.to_excel(writer, sheet_name="Vides")
    Allocation_plein.to_excel(writer, sheet_name="Pleins")
    Allocation_stock.to_excel(writer, sheet_name="Stock")
    Allocation_leasing.to_excel(writer, sheet_name = "Leasing")

# Reduced cost
# for v in prob.variables():
#     print(v.name, v.dj)

# Shadow price
# for (name, c) in list(prob.constraints.items()):
#     print(name, ":", c, "\t", c.pi)

# Pour la 2 : Pas de dual, uniquement analyser les résultats et le données

# Pour la 3 : Analyser le dual de la contrainte de capacité des lignes, subtilité avec les conteneurs VIDES
# On va donc regarder les lignes déjà full et regarder si amener des conteneurs via ces lignes permettrait
# De remplir des porte conteneurs avec gros profit qui ne sont pas remplis dans le port d'arrivée

# Pour la 4 : utiliser le dual de la contrainte de capacité de stockage

print(int(value(prob.objective)))

