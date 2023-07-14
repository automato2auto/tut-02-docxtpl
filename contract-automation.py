from docxtpl import DocxTemplate
import pandas as pd
from num2words import num2words
import os
from timeit import default_timer as Timer

if not os.path.isdir('./output'):

    os.mkdir('./output')

doc = DocxTemplate('sale-contract.docx')

df = pd.read_excel('clients_data.xlsx')

# وقت بدأ العملية
operation_start_time = Timer()

for record in df.to_dict(orient='records'):

    client_context = record

    client_context['sale_month'], client_context['sale_day'], client_context['sale_year'] = client_context['sale_date'].split(
        '/')

    client_context['total_price_word'] = num2words(
        client_context['total_price'])

    doc.render(client_context)

    client_name = client_context["client_name"].lower().replace(' ', '')

    doc.save(f'./output/{client_name}-sale-contract.docx')

# وقت انتهاء العملية
operation_end_time = Timer()

# مدة العملية بالثواني
operation_duration = operation_end_time-operation_start_time

print(f'Operation took {operation_duration} seconds to complete')
