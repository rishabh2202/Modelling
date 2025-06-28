# excel related and mathematical computation based libraries

import xlwings as xw
import numpy as np

#chart library
import matplotlib.pyplot as plt

#directory based library
import os

wb1 = os.getcwd() + "\\" + (xw.books.active.name)

wb = xw.Book(wb1)

inp = wb.sheets["Input"]

#assign the uncertain parameters with their name

inp["C20"].name = "CPI"
inp["C25"].name = "PPF"
inp["C37"].name = "COST"
inp["J14:J15"].name = "Performance"

#Probability distribution of CPI
CPI_exp = .02
CPI_std = .01

#Probability distribution of PPF
PPF_exp = 23
PPF_std = 3

#Probability distribution of Cost
Cost_exp = 250000
Cost_std = 50000


sims = 1000

results = np.empty((sims,2))

for i in range(sims):
    inp["CPI"].value = np.random.normal(CPI_exp, CPI_std)
    inp["PPF"].value = np.random.normal(PPF_exp, PPF_std)
    inp["Cost"].value = np.random.normal(Cost_exp, Cost_std)
    results[i] = inp["Performance"].value

sim_sheet = wb.sheets["Output"]
sim_sheet.range("A1").expand().clear_contents()
sim_sheet.range("A1").value = "Scenario Number"
sim_sheet.range("B1").value = "Investment Multiple"
sim_sheet.range("C1").value = "IRR"

sim_sheet.range("A2").options(transpose = True).value = np.arange(1,sims+1)
sim_sheet.range("B2").value = results
sim_sheet.range("A1").expand().autofit()

fig1 = plt.figure(figsize=(6,4))
plt.hist(results[:,0], bins=100)
plt.title("Investment Multiple Simulations")
plt.xlabel("Investment Multiple")
plt.ylabel("Number of Scenarios")

plot = sim_sheet.pictures.add(fig1, name = "Output Results", update = True)
plot.left = sim_sheet.range("F7").left
plot.top = sim_sheet.range("F7").top

sim_sheet.range("F3").options(transpose = True).value = np.mean(results, axis = 0)
sim_sheet.range("G3").options(transpose = True).value = np.median(results, axis = 0)
sim_sheet.range("F5").value = sum(results[:,0]<1)/sims
