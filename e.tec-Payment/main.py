import tabula

def pdf_read(path):
    tables = tabula.read_pdf(path, pages='all', lattice=True, pandas_options={'header': None})
path = "F:\python_code\e.tec-Payment\B601アーステック様.pdf"
tables = tabula.read_pdf(path, pages='all', lattice=True, pandas_options={'header': None})

print(len(tables))
print(type(tables))



# job-Nomber
job_No = tables[0][1][1]
print(job_No)

# total-amount
total_amount = tables[0][1][2]
print(total_amount)

