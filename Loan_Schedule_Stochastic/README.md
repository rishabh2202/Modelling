# 🏢 Loan Schedule Stochastic Modeling for Real Estate Investment

This project models the financial performance of a real estate investment using Excel and Python (via `xlwings`). It simulates 1,000 scenarios to assess the risk and return profile of purchasing a commercial property, renting it out under a long-term lease, and selling it after 10 years.

---

## 📌 Project Overview

- **Asset Type**: Commercial Property
- **Purchase Price**: ₹100 million
- **Lease Agreement**: 20-year lease with a large corporate tenant
- **Investment Horizon**: 10 years
- **Exit Strategy**: Sell the property after 10 years; lease continues with new investor
- **Funding Structure**:
  - 60% via bank loan (mortgage)
  - 40% via equity
  - Loan repaid fully upon sale

---

## 🎯 Objective

To evaluate the investment's return profile by simulating:

- Annual rental income
- Loan amortization and interest payments
- Property value at exit (uncertain, modeled stochastically)
- Net proceeds from sale
- Overall return on equity (IRR, NPV, etc.)

---

## 🧪 Methodology

1. **Excel Model**:
   - Loan amortization schedule
   - Rental income and cash flow tracking
   - Integration with Python via `xlwings`

2. **Python Simulation**:
   - Generate 1,000 scenarios for property value at year 10
   - Use probability distributions (e.g., lognormal or normal with drift) to model price evolution
   - Calculate equity returns for each scenario
   - Visualize distribution of outcomes (histograms, percentiles, etc.)

3. **Key Metrics**:
   - Investment Multiple
   - Internal Rate of Return (IRR)
   - Shortfall Risk

---

## 🛠️ Tools & Technologies

| Tool       | Purpose                                  |
|------------|------------------------------------------|
| Excel      | Financial modeling and loan schedule     |
| Python     | Simulation and analysis                  |
| xlwings    | Bridge between Excel and Python          |
| NumPy      | Random number generation and math        |           
| Matplotlib | Visualization of results                 |

## Project Structure

- [Loan Schedule.xlsx](/Loan_Schedule_Stochastic/Loan%20Schedule.xlsx)
- [Stochastic_Model.py](/Loan_Schedule_Stochastic/Stochastic_Model.py)
- [README.md](/Loan_Schedule_Stochastic/README.md)

## Author
Developed by Rishabh