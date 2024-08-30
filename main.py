from docxtpl import DocxTemplate
import pandas as pd

doc = DocxTemplate("formato_corrected.docx")
'''context = { 
           'name_work' : "leon de la cruz julio adrian",
           'curp' : "JALC3432432" ,
           'ocupation' : "Ingenier" ,
           'job' : "Developer" ,
           'razon' : "World company" ,
           'rfc' : "2324" ,
           'name_curso' : "dec re" ,
           'hour' : "20" ,
           'ini_year' : "2024" ,
           'i_m' : "08" ,
           'ini_d' : "21" ,
           'te_day' : "2024",
           'te_m' : "08",
           'te_d' : "22",
           'tematica' : "World company",
           'agent_name' : "jane smit",
           'instructor' : "jhon smit",
           'patron' : "linux",
           'representante' : "face" 
           }'''
df = pd.read_excel('datas.xlsx')

for index, row in df.iterrows():
    context = { 
           'name_work' : row['name'],
           'curp' : row['curp'] ,
           'ocupation' : row['ocupation'] ,
           'job' : row['job'] ,
           'razon' : row['razon'] ,
           'rfc' : row['rfc'] ,
           'name_curso' : row['name_curso'] ,
           'hour' : row['hour'] ,
           'ini_year' : row['iyear'] ,
           'i_m' : row['imonth'] ,
           'ini_d' : row['iday'] ,
           'te_day' : row['tyear'],
           'te_m' : row['tmonth'],
           'te_d' : row['tday'],
           'tematica' : row['tematica'],
           'agent_name' : row['agent'],
           'instructor' : row['ins'],
           'patron' : row['pa'],
           'representante' : row['re'] 
           }
    context.update()
    doc.render(context)
    doc.save(f"generated_doc_{index}.docx")